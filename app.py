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

# ตั้งค่าพื้นฐาน
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
DRIVE_FOLDER_ID = "19DLipG-4_C0qWTOsFGXyJWfhLsNvR4V8"
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit?gid=0#gid=0"

HEADER_ROW = 4
CODE_B_INDEX = 1
DESC_COL_INDEX = 3     
NEW_COL_INDEX = 12
OLD_COL_INDEX = 13

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
    creds_dict = json.loads(os.environ.get('GOOGLE_CREDENTIALS_JSON'))
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)

def process_order_and_get_summary(user_msg):
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(GSHEET_URL)
    
    # ดึงค่า Input
    lines = user_msg.strip().split('\n')
    input_dict = {}
    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 2: input_dict[normalize_code(parts[0])] = clean_num(parts[1])
        elif len(parts) == 1: input_dict[normalize_code(parts[0])] = 1.0

    # โหลดไฟล์ Excel
    for f in glob.glob("*.xlsx"): 
        try: os.remove(f)
        except: pass
    drive_service = build('drive', 'v3', credentials=creds)
    results = drive_service.files().list(q=f"'{DRIVE_FOLDER_ID}' in parents and trashed=false", fields="files(id, name)").execute()
    for file in results.get('files', []):
        if file['name'].endswith('.xlsx'): gdown.download(id=file['id'], output=file['name'], quiet=True)

    # ประมวลผล
    report_items = []
    for file_path in glob.glob("*.xlsx"):
        df = pd.read_excel(file_path, header=None, engine='openpyxl')
        headers = df.iloc[HEADER_ROW, :].astype(str).tolist()
        
        # หาตำแหน่งคอลัมน์สำคัญอัตโนมัติ
        bal_idx = -1
        for i, h in enumerate(headers):
            if "Balance" in h or "สต็อกรวม" in h: bal_idx = i
        
        for r in range(HEADER_ROW + 1, len(df)):
            code = normalize_code(str(df.iloc[r, CODE_B_INDEX]))
            if code in input_dict:
                qty_needed = input_dict[code]
                balance = clean_num(df.iloc[r, bal_idx]) if bal_idx != -1 else 0.0
                
                # หาชื่อโครงการ (ขยายช่วง 14 ถึง 100)
                bookings = {}
                for c in range(14, 100):
                    p_name = str(headers[c]).strip()
                    if p_name and not p_name.isdigit() and len(p_name) <= 4:
                        val = clean_num(df.iloc[r, c])
                        if val > 0: bookings[p_name] = int(val)
                
                report_items.append({
                    'code': str(df.iloc[r, CODE_B_INDEX]),
                    'desc': str(df.iloc[r, DESC_COL_INDEX]),
                    'balance': int(balance),
                    'shortage': "มีของ" if balance >= qty_needed else int(qty_needed),
                    'new': int(clean_num(df.iloc[r, NEW_COL_INDEX])),
                    'old': int(clean_num(df.iloc[r, OLD_COL_INDEX])),
                    'bookings': bookings
                })

    # สร้างข้อความสรุป
    summary_text = "📊 รายงานสรุปสต็อก:\n"
    for item in report_items:
        summary_text += f"\n📦 {item['code']} ({item['desc']})\n- สต็อกรวม: {item['balance']}\n- ขาด: {item['shortage']}\n- ของใหม่: {item['new']}\n- ของเก่า: {item['old']}\n"
        if item['bookings']:
            summary_text += "- ติดจอง:\n"
            for p, q in item['bookings'].items(): summary_text += f"  • {p}: {q}\n"
    
    return summary_text

@app.route("/callback", methods=['POST'])
def callback():
    try: handler.handle(request.get_data(as_text=True), request.headers['X-Line-Signature'])
    except: abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=process_order_and_get_summary(event.message.text)))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
