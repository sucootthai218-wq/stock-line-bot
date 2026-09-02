import os
import glob
import re
import json
import time
import threading
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import gspread
import pandas as pd
import gdown

app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

DRIVE_FOLDER_ID = "19DLipG-4_C0qWTOsFGXyJWfhLsNvR4V8"
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit?gid=0#gid=0"

HEADER_ROW = 4
CODE_B_INDEX = 1
CODE_C_INDEX = 2
DESC_COL_INDEX = 3      
NEW_COL_INDEX = 12
OLD_COL_INDEX = 13
MAINT_COL_INDEX = 60         # Maintenance อยู่ช่อง 60
TOTAL_ONHAND_COL_INDEX = 61  # Total Onhand (New+Old) + maintenance อยู่ช่อง 61
ON_HAND_COL_INDEX = 62       
BALANCE_COL_INDEX = 63

# ตัวแปรสำหรับเก็บ Cache ไฟล์ Excel และเวลาที่อัปเดต (กำหนดอายุ 4 ชั่วโมง = 14400 วินาที)
CACHE_DURATION = 4 * 3600
last_download_time = 0
last_download_str = "-"

def clean_num(val):
    if pd.isna(val) or val is None: return 0.0
    s = str(val).replace(',', '').strip()
    if s in ["-", "_", "", "nan", "None"]: return 0.0
    try: return float(s)
    except: return 0.0

def normalize_code(code_str):
    if not code_str: return ""
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

def update_excel_cache(creds):
    global last_download_time, last_download_str
    
    existing_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]

    # ลบไฟล์เก่าทิ้งก่อนดาวน์โหลดใหม่เสมอ
    for f in existing_xlsx:
        try: os.remove(f)
        except: pass

    drive_service = build('drive', 'v3', credentials=creds, static_discovery=False)
    query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    for file in files:
        gdown.download(id=file['id'], output=file['name'], quiet=True)
        
    last_download_time = time.time()
    last_download_str = time.strftime('%d/%m/%Y เวลา %H:%M น.')
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] อัปเดตไฟล์ Excel เข้า Cache เรียบร้อยแล้ว")

def background_sync_loop():
    """ วนลูปทำงานเบื้องหลังเพื่อดาวน์โหลดไฟล์ใหม่ทุกๆ 4 ชั่วโมง """
    while True:
        try:
            print("กำลังตรวจสอบและอัปเดตข้อมูล Excel ในเบื้องหลัง...")
            creds = get_google_credentials()
            update_excel_cache(creds)
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการอัปเดต Cache เบื้องหลัง: {e}")
        
        # รอเวลา 4 ชั่วโมง (14400 วินาที) ค่อยวนกลับมาโหลดใหม่
        time.sleep(CACHE_DURATION)

# เริ่มต้น Thread ทำงานเบื้องหลังทันทีที่แอปเปิดขึ้น
threading.Thread(target=background_sync_loop, daemon=True).start()

def process_order_and_get_summary(user_msg):
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(GSHEET_URL)

    ws_input = spreadsheet.worksheet("Input_Order")
    ws_input.clear()
    
    lines = user_msg.strip().split('\n')
    input_data = [["Code", "Qty"]]
    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 1:
            raw_code = parts[0]
            norm_c = normalize_code(raw_code)
            if not norm_c: continue
            qty = parts[1] if len(parts) >= 2 else "1"
            input_data.append([raw_code, qty])

    if len(input_data) <= 1:
        return "❌ กรุณาระบุรหัสสินค้าที่ต้องการตรวจสอบให้ถูกต้อง"
    ws_input.update(input_data, 'A1')

    # อ่านไฟล์ Excel ที่ถูก Cache ไว้ในเครื่องมาประมวลผลทันทีโดยไม่ต้องโหลดใหม่ตอนลูกคบทักมา
    all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
    global_code_map = {}

    for file_path in all_xlsx:
        df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
        for r in range(HEADER_ROW + 1, len(df_raw)):
            for col_idx in [CODE_B_INDEX, CODE_C_INDEX]:
                if col_idx < df_raw.shape[1]:
                    c_val = df_raw.iloc[r, col_idx]
                    norm_c = normalize_code(c_val)
                    if norm_c and norm_c not in ["NAN", "NONE", "0", "CODE"]:
                        if norm_c not in global_code_map:
                            global_code_map[norm_c] = (df_raw, r)

    df_input = pd.DataFrame(ws_input.get_all_records())
    report_items = []
    
    for _, row in df_input.iterrows():
        raw_code = str(row[df_input.columns[0]]).strip()
        norm_input = normalize_code(raw_code)
        qty_needed = clean_num(row[df_input.columns[1]])
        
        match_data = global_code_map.get(norm_input)
        if match_data:
            df_target, target_row = match_data
            
            new_v = clean_num(df_target.iloc[target_row, NEW_COL_INDEX]) if NEW_COL_INDEX < df_target.shape[1] else 0.0
            old_v = clean_num(df_target.iloc[target_row, OLD_COL_INDEX]) if OLD_COL_INDEX < df_target.shape[1] else 0.0
            maint_v = clean_num(df_target.iloc[target_row, MAINT_COL_INDEX]) if MAINT_COL_INDEX < df_target.shape[1] else 0.0
            total_onhand_v = clean_num(df_target.iloc[target_row, TOTAL_ONHAND_COL_INDEX]) if TOTAL_ONHAND_COL_INDEX < df_target.shape[1] else (new_v + old_v + maint_v)
            
            balance = clean_num(df_target.iloc[target_row, BALANCE_COL_INDEX]) if BALANCE_COL_INDEX < df_target.shape[1] else 0.0
            shortage = qty_needed if balance < 0 else max(0, qty_needed - balance)
            
            proj_bookings = {}
            for c in range(BALANCE_COL_INDEX + 1, df_target.shape[1]):
                txts = [str(df_target.iloc[r, c]).strip() for r in range(0, min(7, len(df_target))) if pd.notna(df_target.iloc[r, c])]
                header_str = " ".join(txts)
                val = clean_num(df_target.iloc[target_row, c])
                
                if val > 0:
                    header_lower = header_str.lower()
                    is_booking_col = "จอง" in header_lower or "po" in header_lower
                    exclude_keywords = ["total", "reserve", "maintenance", "pending", "import", "ek17", "น้ำหนัก", "คงเหลือ", "sale", "rent"]
                    is_excluded = any(kw in header_lower for kw in exclude_keywords)

                    if is_booking_col and not is_excluded:
                        valid_names = [t for t in txts if t.lower() not in ["nan", "none", "c", "e", "จอง", "เวลา", "ใช้", "ยืม", "วันที่", "พค", "มิย", "กค", "สค"]]
                        proj_name = " ".join(valid_names)
                        if proj_name:
                            proj_bookings[proj_name] = int(val)

            report_items.append({
                'code': str(df_target.iloc[target_row, CODE_B_INDEX]), 
                'desc': str(df_target.iloc[target_row, DESC_COL_INDEX]),
                'shortage': "มีของ" if shortage == 0 else int(shortage),
                'on_hand': int(total_onhand_v),
                'balance': int(balance),
                'new': int(new_v) if int(new_v) != 0 else "-",
                'old': int(old_v) if int(old_v) != 0 else "-",
                'maintenance': int(maint_v) if int(maint_v) != 0 else "-",
                'bookings': proj_bookings
            })
        else:
            report_items.append({'code': raw_code, 'desc': "ไม่พบรหัส", 'shortage': int(qty_needed), 'on_hand': "-", 'balance': "-", 'new': "-", 'old': "-", 'maintenance': "-", 'bookings': {}})

    summary_text = "📊 รายงานสรุปสต็อก:\n"
    for item in report_items:
        summary_text += f"\n📦 {item['code']} ({item['desc']})\n"
        summary_text += f"- On hand: {item['on_hand']}\n- สต็อก balance: {item['balance']}\n- ขาด: {item['shortage']}\n- ของใหม่: {item['new']}\n- ของเก่า: {item['old']}\n- maintenance: {item['maintenance']}\n"
        if item['bookings']:
            summary_text += "- ติดจอง:\n"
            for p, q in item['bookings'].items(): summary_text += f"  • {p}: {q}\n"
        else: summary_text += "- ติดจอง: -\n"

    # เพิ่มบรรทัดแสดงเวลาอัปเดตไฟล์ล่าสุดต่อท้ายรายงาน
    summary_text += f"\n🕒 (ข้อมูลจากไฟล์อัปเดตล่าสุดเมื่อ: {last_download_str})"

    return summary_text

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    try: handler.handle(request.get_data(as_text=True), signature)
    except InvalidSignatureError: abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    try: reply = process_order_and_get_summary(event.message.text)
    except Exception as e: reply = f"❌ เกิดข้อผิดพลาด: {str(e)}"
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
