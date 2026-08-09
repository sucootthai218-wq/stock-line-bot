import os
import glob
import re
import json
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

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

DRIVE_FOLDER_ID = "19DLipG-4_C0qWTOsFGXyJWfhLsNvR4V8"
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit?gid=0#gid=0"

HEADER_ROW = 4
CODE_B_INDEX = 1
CODE_C_INDEX = 2
DESC_COL_INDEX = 3     
NEW_COL_INDEX = 12
OLD_COL_INDEX = 13
BALANCE_COL_INDEX = 63

# --- กำหนดช่วงคอลัมน์โครงการ ---
PROJECT_START_COL = 14  
PROJECT_END_COL = 62    
# -----------------------------

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

def process_order_and_get_summary(user_msg):
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(GSHEET_URL)

    # จัดการชีต Input_Order
    try:
        ws_input = spreadsheet.worksheet("Input_Order")
    except:
        ws_input = spreadsheet.add_worksheet(title="Input_Order", rows="100", cols="10")
    
    ws_input.clear()
    
    lines = user_msg.strip().split('\n')
    input_data = [["Code", "Qty"]]
    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 2: input_data.append([parts[0], parts[1]])
        elif len(parts) == 1: input_data.append([parts[0], "1"])
    ws_input.update(input_data, 'A1')

    for f in glob.glob("*.xlsx"): 
        try: os.remove(f)
        except: pass

    drive_service = build('drive', 'v3', credentials=creds)
    query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    for file in results.get('files', []):
        gdown.download(id=file['id'], output=file['name'], quiet=True)

    all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
    global_code_map = {}
    column_headers = {}

    for file_path in all_xlsx:
        df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
        for c in range(df_raw.shape[1]):
            txts = [str(df_raw.iloc[r, c]).strip() for r in range(HEADER_ROW, HEADER_ROW + 2) if pd.notna(df_raw.iloc[r, c])]
            column_headers[c] = " ".join(txts)

        for r in range(HEADER_ROW + 1, len(df_raw)):
            c_val = df_raw.iloc[r, CODE_B_INDEX]
            norm_c = normalize_code(c_val)
            if norm_c and norm_c not in ["CODE", "รหัสสินค้า"]:
                if norm_c not in global_code_map:
                    global_code_map[norm_c] = (df_raw, r, column_headers)

    df_input = pd.DataFrame(ws_input.get_all_records())
    report_items = []

    for _, row in df_input.iterrows():
        norm_input = normalize_code(str(row.iloc[0]))
        qty_needed = clean_num(row.iloc[1])
        
        if norm_input in global_code_map:
            df_target, target_row, headers = global_code_map[norm_input]
            
            bookings = {}
            for col_idx in range(PROJECT_START_COL, PROJECT_END_COL):
                val = clean_num(df_target.iloc[target_row, col_idx])
                if val > 0:
                    proj_name = str(headers.get(col_idx, "")).strip()
                    # กรองเฉพาะชื่อโครงการที่มีความยาวไม่เกิน 4 ตัวอักษร/ตัวเลข (เช่น SKR1, ITA1, T008)
                    if proj_name and re.match(r'^[A-Z0-9]{1,4}$', proj_name.upper()):
                        bookings[proj_name] = int(val)
            
            balance = clean_num(df_target.iloc[target_row, BALANCE_COL_INDEX])
            shortage = "มีของ" if qty_needed <= balance else int(qty_needed - balance)
            
            report_items.append({
                'code': str(df_target.iloc[target_row, CODE_B_INDEX]), 
                'desc': str(df_target.iloc[target_row, DESC_COL_INDEX]),
                'shortage': shortage,
                'balance': int(balance),
                'new': int(clean_num(df_target.iloc[target_row, NEW_COL_INDEX])),
                'old': int(clean_num(df_target.iloc[target_row, OLD_COL_INDEX])),
                'bookings': bookings
            })

    summary_text = "📊 รายงานสรุปสต็อก:\n"
    table_data = [["รหัสสินค้า", "รายการ", "สต็อก Balance", "ขาด", "ของใหม่", "ของเก่า", "ติดจอง"]]
    
    for item in report_items:
        summary_text += f"\n📦 {item['code']} ({item['desc']})\n- สต็อกรวม: {item['balance']}\n- ขาด: {item['shortage']}\n- ของใหม่: {item['new']}\n- ของเก่า: {item['old']}\n"
        if item['bookings']:
            summary_text += "- ติดจอง:\n"
            for proj, qty in item['bookings'].items():
                summary_text += f"  • {proj}: {qty}\n"
        
        booking_str = ", ".join([f"{p}: {q}" for p, q in item['bookings'].items()]) if item['bookings'] else "-"
        table_data.append([item['code'], item['desc'], item['balance'], item['shortage'], item['new'], item['old'], booking_str])

    # บันทึกข้อมูลลง Google Sheets (ชีต Summary)
    try:
        try:
            ws_summary = spreadsheet.worksheet("Summary")
        except:
            ws_summary = spreadsheet.add_worksheet(title="Summary", rows="100", cols="10")
            
        ws_summary.clear()
        ws_summary.update(table_data, 'A1')
    except Exception as e:
        print(f"Error updating Summary sheet: {e}")

    return summary_text

@app.route("/callback", methods=['POST'])
def callback():
    body = request.get_data(as_text=True)
    try: handler.handle(body, request.headers['X-Line-Signature'])
    except: abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    reply_text = process_order_and_get_summary(event.message.text)
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
