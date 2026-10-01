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

SCOPES = ["https://www.googleapis.com/auth/drive"]
DRIVE_FOLDER_ID = "19DLipG-4_C0qWTOsFGXyJWfhLsNvR4V8"

CACHE_FILE = "stock_cache.json"

HEADER_ROW = 5  # Excel row index (1-based)
CODE_B_COL = 2  # Col B
CODE_C_COL = 3  # Col C
DESC_COL = 4    # Col D
NEW_COL = 13    # Col M
OLD_COL = 14    # Col N
MAINT_COL = 61  # Col BI
TOTAL_ONHAND_COL = 62  # Col BJ
BALANCE_COL = 64       # Col BL

CACHE_DURATION = 28800  # 8 ชั่วโมง
is_updating = False

# ==========================================
# 🛠️ ฟังก์ชัน Utility
# ==========================================
def clean_num(val):
    if val is None:
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

def col2num_1based(col_str):
    num = 0
    for c in col_str.upper():
        num = num * 26 + (ord(c) - ord('A')) + 1
    return num

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
# 🔄 โหลดไฟล์ Excel และบันทึกลง JSON Cache
# ==========================================
def update_excel_cache():
    global is_updating

    if is_updating:
        print("⚠️ กำลังมีกระบวนการอัปเดต Cache อยู่แล้ว ข้ามรอบนี้")
        return False
    
    is_updating = True
    try:
        creds = get_google_credentials()

        for f in glob.glob("*.xlsx"):
            if not os.path.basename(f).startswith("~$"):
                try: os.remove(f)
                except Exception: pass

        drive_service = build('drive', 'v3', credentials=creds, static_discovery=False)
        query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
        results = drive_service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])

        if not files:
            print("⚠️ ไม่พบไฟล์ Excel ใน Google Drive")
            return False

        for file in files:
            try:
                req = drive_service.files().get_media(fileId=file['id'])
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                fh.seek(0)
                with open(file['name'], 'wb') as f:
                    f.write(fh.read())
                print(f"✅ ดาวน์โหลด {file['name']} สำเร็จ")
            except Exception as e:
                print(f"❌ ดาวน์โหลดล้มเหลว: {e}")

        new_stock_map = {}
        hard_exclude = ["total reserve", "น้ำหนัก", "total maintenance", "lot", "eta", "หักจอง", "pr26", "po26", "รถ", "so26"]

        for file_path in glob.glob("*.xlsx"):
            if os.path.basename(file_path).startswith("~$"):
                continue

            try:
                target_booking_cols = set()
                wb_formula = openpyxl.load_workbook(file_path, data_only=False, read_only=True)
                ws_formula = wb_formula.active

                for r in range(HEADER_ROW, HEADER_ROW + 300):
                    cell_val = str(ws_formula.cell(row=r, column=BALANCE_COL).value or '')
                    if "+" in cell_val:
                        cols = re.findall(r'([A-Z]+)\d+', cell_val)
                        for c in cols:
                            col_idx = col2num_1based(c)
                            if col_idx > BALANCE_COL:
                                target_booking_cols.add(col_idx)
                        if target_booking_cols:
                            break
                wb_formula.close()

                wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
                ws = wb.active

                header_rows_data = []
                for r in range(1, HEADER_ROW + 1):
                    header_rows_data.append([cell.value for cell in ws[r]])

                col_meta = {}
                for col_idx in target_booking_cols:
                    parts = []
                    for r_idx in range(len(header_rows_data)):
                        if col_idx - 1 < len(header_rows_data[r_idx]):
                            val = header_rows_data[r_idx][col_idx - 1]
                            if val is not None:
                                s = str(val).strip()
                                if s.lower() not in ["nan", "none", "c", "e", "null", ""]:
                                    parts.append(s)
                    proj_name = " ".join(parts).strip()
                    if proj_name and not any(kw in proj_name.lower() for kw in hard_exclude):
                        col_meta[col_idx] = proj_name

                for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
                    if not row or len(row) < BALANCE_COL:
                        continue

                    code_b = row[CODE_B_COL - 1] if len(row) >= CODE_B_COL else None
                    code_c = row[CODE_C_COL - 1] if len(row) >= CODE_C_COL else None

                    extracted_codes = []
                    for c_val in [code_b, code_c]:
                        norm = normalize_code(c_val)
                        if norm and norm not in ["NAN", "NONE", "0", "CODE"]:
                            extracted_codes.append(norm)

                    if not extracted_codes:
                        continue

                    desc_val = str(row[DESC_COL - 1]).strip() if len(row) >= DESC_COL and row[DESC_COL - 1] else "-"
                    new_v = clean_num(row[NEW_COL - 1]) if len(row) >= NEW_COL else 0.0
                    old_v = clean_num(row[OLD_COL - 1]) if len(row) >= OLD_COL else 0.0
                    maint_v = clean_num(row[MAINT_COL - 1]) if len(row) >= MAINT_COL else 0.0
                    total_onhand_v = clean_num(row[TOTAL_ONHAND_COL - 1]) if len(row) >= TOTAL_ONHAND_COL else (new_v + old_v + maint_v)
                    balance_v = clean_num(row[BALANCE_COL - 1]) if len(row) >= BALANCE_COL else 0.0

                    proj_bookings = {}
                    for col_idx, proj_name in col_meta.items():
                        if col_idx - 1 < len(row):
                            v = clean_num(row[col_idx - 1])
                            if v > 0:
                                proj_bookings[proj_name] = int(v)

                    item_info = {
                        'code': str(code_b) if code_b else extracted_codes[0],
                        'desc': desc_val,
                        'new': int(new_v) if int(new_v) != 0 else "-",
                        'old': int(old_v) if int(old_v) != 0 else "-",
                        'maintenance': int(maint_v) if int(maint_v) != 0 else "-",
                        'on_hand': int(total_onhand_v),
                        'balance': int(balance_v),
                        'total_booked': sum(proj_bookings.values()),
                        'bookings': proj_bookings
                    }

                    for norm_c in extracted_codes:
                        if norm_c not in new_stock_map:
                            new_stock_map[norm_c] = item_info

                wb.close()
                del header_rows_data
            except Exception as e:
                print(f"❌ Error processing {file_path}: {e}")

        tz = pytz.timezone('Asia/Bangkok')
        now_str = datetime.datetime.now(tz).strftime('%d/%m/%Y เวลา %H:%M น.')

        # บันทึกข้อมูลลงเป็นไฟล์ JSON บนเครื่อง เพื่อแชร์ข้อมูลให้ทุก Worker
        cache_payload = {
            "last_updated": now_str,
            "data": new_stock_map
        }
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache_payload, f, ensure_ascii=False)

        gc.collect()
        print(f"✅ บันทึก Cache ลง JSON สำเร็จ: {len(new_stock_map)} รายการ ({now_str})")
        return True
    finally:
        is_updating = False

def load_stock_cache():
    """โหลดข้อมูลจากไฟล์ JSON ที่เซฟไว้"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return None

# ==========================================
# ⚙️ Background Thread & Endpoints
# ==========================================
def background_sync_loop():
    time.sleep(3)
    while True:
        try:
            update_excel_cache()
        except Exception as e:
            print(f"Background Sync Error: {e}")
        time.sleep(CACHE_DURATION)

# รัน Sync เบื้องหลัง
threading.Thread(target=background_sync_loop, daemon=True).start()

@app.route("/", methods=['GET'])
def index():
    return "Stock Bot Service is Live!", 200

@app.route("/cron-sync", methods=['GET'])
def cron_sync():
    threading.Thread(target=update_excel_cache).start()
    return "Triggered sync in background.", 200

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
            if not norm_c: continue
            qty_val = clean_num(parts[1]) if len(parts) >= 2 else 1.0
            parsed_requests.append((raw_code, norm_c, qty_val))

    if not parsed_requests:
        return "❌ กรุณาระบุรหัสสินค้าที่ต้องการตรวจสอบ เช่น:\nST01 100\nST02"

    cache_payload = load_stock_cache()
    if not cache_payload or not cache_payload.get("data"):
        return "⏳ บอทกำลังซิงค์ฐานข้อมูลสต็อกเริ่มต้น กรุณารอสักครู่แล้วลองพิมพ์ใหม่อีกครั้งครับ"

    stock_data = cache_payload["data"]
    last_update_str = cache_payload.get("last_updated", "-")
    report_items = []

    for raw_code, norm_c, qty_needed in parsed_requests:
        item = stock_data.get(norm_c)
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

    summary_text += f"\n🕒 (ข้อมูลอัปเดตล่าสุด: {last_update_str})"
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
