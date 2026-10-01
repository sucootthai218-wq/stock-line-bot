import os
import glob
import re
import json
import time
import threading
import datetime
import gc
import pytz
import io
import pandas as pd
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

app = Flask(__name__)

# ==========================================
# 📌 การตั้งค่า LINE และ Google Drive
# ==========================================
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

DRIVE_FOLDER_ID = "19DLipG-4_C0qWTOsFGXyJWfhLsNvR4V8"

HEADER_ROW = 4
CODE_B_INDEX = 1
CODE_C_INDEX = 2
DESC_COL_INDEX = 3        
NEW_COL_INDEX = 12
OLD_COL_INDEX = 13
MAINT_COL_INDEX = 60            
TOTAL_ONHAND_COL_INDEX = 61  
ON_HAND_COL_INDEX = 62       
BALANCE_COL_INDEX = 63

CACHE_DURATION = 28800  # 8 ชั่วโมง
last_download_time = 0
last_download_str = "-"

# Global Cache สำหรับเก็บข้อมูลในหน่วยความจำ
STOCK_CACHE = {}
cache_lock = threading.Lock()

# ==========================================
# 🛠️ ฟังก์ชัน Utility
# ==========================================
def clean_num(val):
    if pd.isna(val) or val is None:
        return 0.0
    s = str(val).replace(',', '').strip()
    if s in ["-", "_", "", "nan", "None"]:
        return 0.0
    try:
        return float(s)
    except:
        return 0.0

def normalize_code(code_str):
    if not code_str:
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(code_str).strip().upper())

def get_google_credentials():
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if google_creds_json:
        creds_dict = json.loads(google_creds_json)
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        return Credentials.from_service_account_file("credentials.json", scopes=SCOPES)

# ==========================================
# 🔄 การโหลดและประมวลผลไฟล์ Excel เก็บใน Memory
# ==========================================
def update_excel_cache(creds):
    global last_download_time, last_download_str, STOCK_CACHE

    existing_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
    for f in existing_xlsx:
        try:
            os.remove(f)
        except Exception:
            pass

    drive_service = build('drive', 'v3', credentials=creds, static_discovery=False)
    query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])

    if not files:
        print("⚠️ ไม่พบไฟล์ Excel ในโฟลเดอร์ Google Drive ที่กำหนด")
        return

    for file in files:
        file_id = file['id']
        file_name = file['name']
        try:
            request_file = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request_file)
            done = False
            while done is False:
                status, done = downloader.next_chunk()

            fh.seek(0)
            with open(file_name, 'wb') as f:
                f.write(fh.read())
            print(f"✅ ดาวน์โหลดไฟล์ {file_name} สำเร็จ")
        except Exception as e:
            print(f"❌ ดาวน์โหลดไฟล์ {file_name} ไม่สำเร็จ: {e}")

    new_stock_map = {}
    all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]

    for file_path in all_xlsx:
        try:
            df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
            num_cols = df_raw.shape[1]

            # กรองและสกัดหัวคอลัมน์การจอง โดยตัดคอลัมน์คำนวณภายใน/หักจองออก
            col_booking_meta = {}
            for c in range(BALANCE_COL_INDEX + 1, num_cols):
                txts = [str(df_raw.iloc[r, c]).strip() for r in range(0, min(7, len(df_raw))) if pd.notna(df_raw.iloc[r, c])]
                header_str = " ".join(txts)
                header_lower = header_str.lower()

                # คำที่บ่งบอกว่าเป็นคอลัมน์คำนวณภายใน/คอลัมน์ตัดยอด ไม่ใช่รายการจองงานจริง
                exclude_keywords = [
                    "total", "reserve", "maintenance", "pending", "import", 
                    "ek17", "น้ำหนัก", "คงเหลือ", "sale", "rent", 
                    "หักจอง", "หัก จอง", "balance"
                ]
                is_excluded = any(kw in header_lower for kw in exclude_keywords)

                is_booking_col = ("จอง" in header_lower or "po" in header_lower or "ใช้" in header_lower)

                if is_booking_col and not is_excluded:
                    # เก็บข้อความ วันที่จอง เวลา วันที่ใช้ และชื่อโครงการไว้ครบถ้วน
                    clean_tokens = [t for t in txts if t.lower() not in ["nan", "none", "c", "e", "null", ""]]
                    proj_name = " ".join(clean_tokens).strip()

                    # ต้องมีข้อความระบุรายละเอียดและไม่ใช่แค่ตัวเลขหรือคำสั้นๆ
                    if proj_name and len(proj_name) > 3:
                        col_booking_meta[c] = proj_name

            # อ่านข้อมูลแถวสินค้า
            for r in range(HEADER_ROW + 1, len(df_raw)):
                extracted_codes = []
                for col_idx in [CODE_B_INDEX, CODE_C_INDEX]:
                    if col_idx < num_cols:
                        c_val = df_raw.iloc[r, col_idx]
                        norm_c = normalize_code(c_val)
                        if norm_c and norm_c not in ["NAN", "NONE", "0", "CODE"]:
                            extracted_codes.append(norm_c)

                if not extracted_codes:
                    continue

                new_v = clean_num(df_raw.iloc[r, NEW_COL_INDEX]) if NEW_COL_INDEX < num_cols else 0.0
                old_v = clean_num(df_raw.iloc[r, OLD_COL_INDEX]) if OLD_COL_INDEX < num_cols else 0.0
                maint_v = clean_num(df_raw.iloc[r, MAINT_COL_INDEX]) if MAINT_COL_INDEX < num_cols else 0.0
                
                if TOTAL_ONHAND_COL_INDEX < num_cols:
                    total_onhand_v = clean_num(df_raw.iloc[r, TOTAL_ONHAND_COL_INDEX])
                else:
                    total_onhand_v = new_v + old_v + maint_v

                balance_v = clean_num(df_raw.iloc[r, BALANCE_COL_INDEX]) if BALANCE_COL_INDEX < num_cols else 0.0

                # ดึงเฉพาะยอดการจองที่มีค่ามากกว่า 0
                proj_bookings = {}
                for col_idx, proj_name in col_booking_meta.items():
                    val = clean_num(df_raw.iloc[r, col_idx])
                    if val > 0:
                        proj_bookings[proj_name] = int(val)

                item_info = {
                    'code': str(df_raw.iloc[r, CODE_B_INDEX]) if CODE_B_INDEX < num_cols and pd.notna(df_raw.iloc[r, CODE_B_INDEX]) else extracted_codes[0],
                    'desc': str(df_raw.iloc[r, DESC_COL_INDEX]) if DESC_COL_INDEX < num_cols and pd.notna(df_raw.iloc[r, DESC_COL_INDEX]) else "-",
                    'new': int(new_v) if int(new_v) != 0 else "-",
                    'old': int(old_v) if int(old_v) != 0 else "-",
                    'maintenance': int(maint_v) if int(maint_v) != 0 else "-",
                    'on_hand': int(total_onhand_v),
                    'balance': int(balance_v),
                    'bookings': proj_bookings
                }

                for norm_c in extracted_codes:
                    if norm_c not in new_stock_map:
                        new_stock_map[norm_c] = item_info

            del df_raw
        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {file_path}: {e}")

    with cache_lock:
        STOCK_CACHE = new_stock_map

    last_download_time = time.time()
    tz = pytz.timezone('Asia/Bangkok')
    last_download_str = datetime.datetime.now(tz).strftime('%d/%m/%Y เวลา %H:%M น.')

    gc.collect()
    print(f"[{datetime.datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')}] สร้าง Cache สำเร็จ: {len(new_stock_map)} รายการ (เวลา: {last_download_str})")

# ==========================================
# ⚙️ Background Threads & Endpoints
# ==========================================
def background_sync_loop():
    while True:
        try:
            print("กำลังตรวจสอบและอัปเดตข้อมูล Excel ในเบื้องหลัง...")
            creds = get_google_credentials()
            update_excel_cache(creds)
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการอัปเดต Cache เบื้องหลัง: {e}")
        time.sleep(CACHE_DURATION)

try:
    print("กำลังดาวน์โหลดและเตรียม Cache ครั้งแรก...")
    initial_creds = get_google_credentials()
    update_excel_cache(initial_creds)
except Exception as e:
    print(f"เกิดข้อผิดพลาดในการดาวน์โหลดเริ่มต้น: {e}")

threading.Thread(target=background_sync_loop, daemon=True).start()

@app.route("/cron-sync", methods=['GET'])
def cron_sync():
    try:
        print("ได้รับสัญญาณ Cron-Job ภายนอก กำลังอัปเดตไฟล์ Excel...")
        creds = get_google_credentials()
        update_excel_cache(creds)
        return f"Sync Success at {last_download_str}", 200
    except Exception as e:
        return f"Sync Error: {str(e)}", 500

# ==========================================
# 🔍 การประมวลผลคำสั่งเช็คสต็อก
# ==========================================
def process_order_and_get_summary(user_msg):
    lines = user_msg.strip().split('\n')
    parsed_requests = []

    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 1:
            raw_code = parts[0]
            norm_c = normalize_code(raw_code)
            if not norm_c:
                continue
            qty_val = clean_num(parts[1]) if len(parts) >= 2 else 1.0
            parsed_requests.append((raw_code, norm_c, qty_val))

    if not parsed_requests:
        return "❌ กรุณาระบุรหัสสินค้าที่ต้องการตรวจสอบ เช่น:\nST01 100\nST02"

    report_items = []

    with cache_lock:
        current_cache = STOCK_CACHE

    for raw_code, norm_c, qty_needed in parsed_requests:
        item = current_cache.get(norm_c)
        if item:
            bal = item['balance']
            shortage = qty_needed if bal < 0 else max(0, qty_needed - bal)
            
            report_items.append({
                'code': item['code'],
                'desc': item['desc'],
                'shortage': "มีของ" if shortage == 0 else int(shortage),
                'on_hand': item['on_hand'],
                'balance': item['balance'],
                'new': item['new'],
                'old': item['old'],
                'maintenance': item['maintenance'],
                'bookings': item['bookings']
            })
        else:
            report_items.append({
                'code': raw_code,
                'desc': "ไม่พบรหัสสินค้าในสต็อก",
                'shortage': int(qty_needed),
                'on_hand': "-",
                'balance': "-",
                'new': "-",
                'old': "-",
                'maintenance': "-",
                'bookings': {}
            })

    # ประกอบข้อความสรุปรายงาน
    summary_text = "📊 รายงานสรุปสต็อก:\n"
    for item in report_items:
        summary_text += f"\n📦 {item['code']} ({item['desc']})\n"
        summary_text += f"- On hand: {item['on_hand']}\n"
        summary_text += f"- สต็อก balance: {item['balance']}\n"
        summary_text += f"- ขาด: {item['shortage']}\n"
        summary_text += f"- ของใหม่: {item['new']}\n"
        summary_text += f"- ของเก่า: {item['old']}\n"
        summary_text += f"- maintenance: {item['maintenance']}\n"

        if item['bookings']:
            total_booked = sum(item['bookings'].values())
            summary_text += f"- ติดจอง (รวม {total_booked}):\n"
            for p, q in item['bookings'].items():
                summary_text += f"   • {p}: {q}\n"
        else:
            summary_text += "- ติดจอง: -\n"

    summary_text += f"\n🕒 (ข้อมูลอัปเดตล่าสุด: {last_download_str})"
    return summary_text

# ==========================================
# 📩 LINE Webhook Routes
# ==========================================
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    try:
        reply = process_order_and_get_summary(event.message.text)
    except Exception as e:
        reply = f"❌ เกิดข้อผิดพลาดในการค้นหา: {str(e)}"
    
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
