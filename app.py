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
import openpyxl
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
last_download_str = "กำลังโหลดข้อมูลเริ่มต้น..."
is_syncing = False

# Global Cache สำหรับเก็บสต็อกสินค้า และข้อมูลโครงการ
STOCK_CACHE = {}
PROJECT_CACHE = {}
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

def col2num(col_str):
    num = 0
    for c in col_str.upper():
        num = num * 26 + (ord(c) - ord('A')) + 1
    return num - 1

def extract_dynamic_booking_columns(file_path):
    """อ่านสูตรจาก Balance ใน Excel เพื่อดูว่าบวกคอลัมน์ไหนบ้าง"""
    valid_cols = set()
    try:
        wb = openpyxl.load_workbook(file_path, data_only=False, read_only=True)
        sheet = wb.active
        
        target_formula = ""
        for r in range(HEADER_ROW + 1, min(HEADER_ROW + 350, sheet.max_row or 350)):
            cell_val = str(sheet.cell(row=r, column=BALANCE_COL_INDEX + 1).value or '')
            if "+" in cell_val or "SUM" in cell_val.upper():
                target_formula = cell_val
                break
        wb.close()

        if target_formula:
            col_letters = re.findall(r'([A-Z]+)\d+', target_formula)
            for c in col_letters:
                c_idx = col2num(c)
                if c_idx > BALANCE_COL_INDEX:
                    valid_cols.add(c_idx)
            if valid_cols:
                return valid_cols
    except Exception as e:
        print(f"⚠️ แกะสูตร Dynamic ไม่สำเร็จ: {e}")

    return None

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
    global last_download_time, last_download_str, STOCK_CACHE, PROJECT_CACHE, is_syncing

    if is_syncing:
        print("⏳ การ Sync กำลังดำเนินการอยู่ ข้ามรอบนี้...")
        return
    is_syncing = True

    try:
        existing_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
        for f in existing_xlsx:
            try: os.remove(f)
            except Exception: pass

        drive_service = build('drive', 'v3', credentials=creds, static_discovery=False)
        query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
        results = drive_service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])

        valid_files = [f for f in files if f['name'].lower().endswith('.xlsx') and not f['name'].startswith('~$')]

        if not valid_files:
            print("⚠️ ไม่พบไฟล์ Excel (.xlsx) ใน Google Drive")
            is_syncing = False
            return

        for file in valid_files:
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
        new_project_map = {}
        all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]

        for file_path in all_xlsx:
            try:
                dynamic_booking_cols = extract_dynamic_booking_columns(file_path)
                df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
                num_cols = df_raw.shape[1]

                hard_exclude = [
                    "total reserve", "น้ำหนัก", "total maintenance", "lot", 
                    "eta", "หักจอง", "pr26"
                ]

                col_booking_meta = {}

                if dynamic_booking_cols:
                    for c in sorted(list(dynamic_booking_cols)):
                        if c < num_cols:
                            txts = [str(df_raw.iloc[r, c]).strip() for r in range(0, min(7, len(df_raw))) if pd.notna(df_raw.iloc[r, c])]
                            clean_tokens = [t for t in txts if t.lower() not in ["nan", "none", "c", "e", "null", ""]]
                            proj_name = " ".join(clean_tokens).strip()
                            proj_lower = proj_name.lower()

                            if proj_name and not any(kw in proj_lower for kw in hard_exclude):
                                col_booking_meta[c] = proj_name
                else:
                    for c in range(116, num_cols, 2):
                        txts = [str(df_raw.iloc[r, c]).strip() for r in range(0, min(7, len(df_raw))) if pd.notna(df_raw.iloc[r, c])]
                        clean_tokens = [t for t in txts if t.lower() not in ["nan", "none", "c", "e", "null", ""]]
                        proj_name = " ".join(clean_tokens).strip()
                        proj_lower = proj_name.lower()

                        if "total reserve" in proj_lower or "น้ำหนัก" in proj_lower:
                            break

                        if proj_name and not any(kw in proj_lower for kw in hard_exclude):
                            col_booking_meta[c] = proj_name

                unique_col_display_name = {}
                temp_name_count = {}
                for col_idx, proj_name in col_booking_meta.items():
                    if proj_name not in temp_name_count:
                        temp_name_count[proj_name] = 1
                        unique_col_display_name[col_idx] = proj_name
                    else:
                        temp_name_count[proj_name] += 1
                        unique_col_display_name[col_idx] = f"{proj_name} #{temp_name_count[proj_name]}"

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

                    item_code_str = str(df_raw.iloc[r, CODE_B_INDEX]) if CODE_B_INDEX < num_cols and pd.notna(df_raw.iloc[r, CODE_B_INDEX]) else extracted_codes[0]
                    item_desc_str = str(df_raw.iloc[r, DESC_COL_INDEX]) if DESC_COL_INDEX < num_cols and pd.notna(df_raw.iloc[r, DESC_COL_INDEX]) else "-"

                    proj_bookings = {}
                    for col_idx, disp_name in unique_col_display_name.items():
                        val = clean_num(df_raw.iloc[r, col_idx])
                        if val > 0:
                            proj_bookings[disp_name] = int(val)

                            orig_proj_name = col_booking_meta[col_idx]
                            proj_key = f"{file_path}_{col_idx}"
                            if proj_key not in new_project_map:
                                new_project_map[proj_key] = {
                                    'display_name': disp_name,
                                    'raw_name': orig_proj_name,
                                    'file': os.path.basename(file_path),
                                    'items': []
                                }
                            new_project_map[proj_key]['items'].append({
                                'code': item_code_str,
                                'desc': item_desc_str,
                                'qty': int(val)
                            })

                    total_booked_calc = sum(proj_bookings.values())

                    item_info = {
                        'code': item_code_str,
                        'desc': item_desc_str,
                        'new': int(new_v) if int(new_v) != 0 else "-",
                        'old': int(old_v) if int(old_v) != 0 else "-",
                        'maintenance': int(maint_v) if int(maint_v) != 0 else "-",
                        'on_hand': int(total_onhand_v),
                        'balance': int(balance_v),
                        'total_booked': int(total_booked_calc),
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
            PROJECT_CACHE = new_project_map

        last_download_time = time.time()
        tz = pytz.timezone('Asia/Bangkok')
        last_download_str = datetime.datetime.now(tz).strftime('%d/%m/%Y เวลา %H:%M น.')

        gc.collect()
        print(f"[{datetime.datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')}] สร้าง Cache สำเร็จ: สินค้า {len(new_stock_map)} รายการ, โครงการจอง {len(new_project_map)} คอลัมน์")
    finally:
        is_syncing = False

# ==========================================
# ⚙️ Background Threads & Endpoints
# ==========================================
def background_sync_loop():
    # ให้เริ่มรอบแรกหลังจากสตาร์ตเซิร์ฟเวอร์เสร็จทันที (ไม่บล็อกพอร์ต)
    time.sleep(2)
    while True:
        try:
            creds = get_google_credentials()
            update_excel_cache(creds)
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการอัปเดต Cache เบื้องหลัง: {e}")
        time.sleep(CACHE_DURATION)

# รัน Thread แยกเป็น Background ทันที ไม่ให้บล็อกการบูตของ Web Server
threading.Thread(target=background_sync_loop, daemon=True).start()

@app.route("/", methods=['GET'])
def index():
    return "Stock Bot Service is Live!", 200

@app.route("/cron-sync", methods=['GET'])
def cron_sync():
    def _run():
        try:
            creds = get_google_credentials()
            update_excel_cache(creds)
        except Exception as e:
            print(f"Cron Sync Error: {e}")

    threading.Thread(target=_run, daemon=True).start()
    return f"Triggered Sync successfully (Last updated: {last_download_str})", 200

# ==========================================
# 🔍 การประมวลผลคำสั่งค้นหาตามโครงการ (By Project)
# ==========================================
def process_project_query(query_keyword):
    clean_keyword = query_keyword.strip().lower()
    
    with cache_lock:
        current_projects = list(PROJECT_CACHE.values())
        curr_time_str = last_download_str

    if not current_projects:
        return f"⏳ ระบบกำลังดาวน์โหลดข้อมูลสต็อกเริ่มต้น กรุณารอสักครู่แล้วลองใหม่อีกครั้ง"

    matched_projects = []
    for proj in current_projects:
        if clean_keyword in proj['raw_name'].lower() or clean_keyword in proj['display_name'].lower():
            if proj['items']:
                matched_projects.append(proj)

    if not matched_projects:
        return f"❌ ไม่พบโครงการที่ตรงกับคำค้นหา: '{query_keyword}'\n🕒 (ข้อมูลอัปเดตล่าสุด: {curr_time_str})"

    report_text = f"📋 รายการจองโครงการ: '{query_keyword}'\n"
    for idx, proj in enumerate(matched_projects, 1):
        total_qty = sum(item['qty'] for item in proj['items'])
        report_text += f"\n📌 [{idx}] {proj['display_name']}\n"
        report_text += f"🔢 รวมจอง: {len(proj['items'])} รายการ ({total_qty:,} ชิ้น)\n"
        report_text += "-------------------------\n"
        for item in proj['items']:
            report_text += f"   • {item['code']} ({item['desc']}): {item['qty']:,} ชิ้น\n"

    report_text += f"\n🕒 (ข้อมูลอัปเดตล่าสุด: {curr_time_str})"
    return report_text

# ==========================================
# 🔍 การประมวลผลคำสั่งเช็คสต็อกสินค้า (By Product)
# ==========================================
def process_order_and_get_summary(user_msg):
    lines = user_msg.strip().split('\n')
    parsed_requests = []

    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 1:
            raw_code = parts[0]
            norm_c = normalize_code(raw_code)
            if not norm_c: continue
            qty_val = clean_num(parts[1]) if len(parts) >= 2 else 1.0
            parsed_requests.append((raw_code, norm_c, qty_val))

    if not parsed_requests:
        return "❌ กรุณาระบุรหัสสินค้าที่ต้องการตรวจสอบ เช่น:\nST01 100\nST02\n\nหรือค้นหาโครงการ เช่น:\nโครงการ T008\nโครงการ PO93"

    report_items = []

    with cache_lock:
        current_cache = STOCK_CACHE
        curr_time_str = last_download_str

    if not current_cache:
        return "⏳ ระบบกำลังดาวน์โหลดข้อมูลสต็อกเริ่มต้น กรุณารอสักครู่แล้วลองใหม่อีกครั้ง"

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
                'total_booked': item['total_booked'],
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
                'total_booked': "-",
                'bookings': {}
            })

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
            summary_text += f"- ติดจอง (รวม {item['total_booked']}):\n"
            for p, q in item['bookings'].items():
                summary_text += f"   • {p}: {q}\n"
        else:
            summary_text += "- ติดจอง: -\n"

    summary_text += f"\n🕒 (ข้อมูลอัปเดตล่าสุด: {curr_time_str})"
    return summary_text

# ==========================================
# 📩 LINE Webhook Routes & Router
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
    raw_text = event.message.text.strip()
    
    project_prefixes = ["โครงการ", "project", "pj", "งาน"]
    is_project_query = False
    keyword = ""

    for prefix in project_prefixes:
        pattern = rf"^{prefix}[:\s]+(.+)$"
        match = re.match(pattern, raw_text, re.IGNORECASE)
        if match:
            is_project_query = True
            keyword = match.group(1).strip()
            break

    try:
        if is_project_query and keyword:
            reply = process_project_query(keyword)
        else:
            reply = process_order_and_get_summary(raw_text)
    except Exception as e:
        reply = f"❌ เกิดข้อผิดพลาดในการประมวลผล: {str(e)}"
    
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
