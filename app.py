import os
import glob
import re
import json
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import gdown

app = Flask(__name__)

# ตั้งค่า Configuration
FILE_IDS = ["1M7YJzGsNRTSQswyxdHHiDDXAuX1h7U4a", "1HwUNEZ1wwwne2-ZG0ogluODTdmAAdTSg"]
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit?gid=0#gid=0"
HEADER_ROW = 4 # บรรทัดที่เป็นหัวตาราง
CODE_B_INDEX = 1
CODE_C_INDEX = 2
DESC_COL_INDEX = 3

# ฟังก์ชันจัดการ Credentials
def get_google_credentials():
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    creds_dict = json.loads(google_creds_json)
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    return Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])

# ฟังก์ชันทำความสะอาดเลข
def clean_num(val):
    if pd.isna(val) or val is None: return 0.0
    s = str(val).replace(',', '').strip()
    try: return float(s)
    except: return 0.0

def normalize_code(code_str):
    if not code_str: return ""
    return re.sub(r'[^A-Z0-9]', '', str(code_str).strip().upper())

def process_order_and_get_summary(user_msg):
    # 1. จัดการไฟล์: ถ้าไม่มีไฟล์ในโฟลเดอร์ค่อยโหลดใหม่
    existing_files = glob.glob("stock_data_*.xlsx")
    if len(existing_files) < len(FILE_IDS):
        for i, file_id in enumerate(FILE_IDS):
            filename = f"stock_data_{i}.xlsx"
            gdown.download(id=file_id, output=filename, quiet=False)
    
    # 2. อ่านข้อมูลโดยค้นหาคอลัมน์ Balance อัตโนมัติ
    global_code_map = {}
    for file_path in glob.glob("stock_data_*.xlsx"):
        df_raw = pd.read_excel(file_path, header=HEADER_ROW)
        
        # ค้นหาคอลัมน์ที่มีคำว่า 'Balance', 'Stock', 'คงเหลือ'
        bal_col_name = next((col for col in df_raw.columns if any(kw in str(col).lower() for kw in ['balance', 'stock', 'คงเหลือ'])), None)
        
        for _, row in df_raw.iterrows():
            code = str(row.iloc[0]) # รหัสสินค้า (คอลัมน์แรกหลัง Header)
            norm_c = normalize_code(code)
            if norm_c and norm_c not in global_code_map:
                global_code_map[norm_c] = {
                    'desc': row.iloc[2], # คำอธิบายสินค้า
                    'balance': clean_num(row[bal_col_name] if bal_col_name else 0)
                }

    # 3. คำนวณยอด
    lines = user_msg.strip().split('\n')
    report_items = []
    for line in lines:
        parts = line.split()
        code_input = normalize_code(parts[0])
        qty_needed = clean_num(parts[1]) if len(parts) > 1 else 1.0
        
        if code_input in global_code_map:
            data = global_code_map[code_input]
            shortage = max(0, qty_needed - data['balance'])
            report_items.append(f"- {parts[0]} ({data['desc']}): คงเหลือ {int(data['balance'])} | ขาด {int(shortage)}")
        else:
            report_items.append(f"- {parts[0]}: ไม่พบรหัสในสต็อก")

    return "📊 รายงานสรุปสต็อก:\n" + "\n".join(report_items)

# ... (ส่วนของ Flask Route และ Handler เหมือนเดิม)
