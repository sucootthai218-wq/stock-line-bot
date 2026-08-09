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

# ดึงค่า LINE Token และ Secret จาก Environment Variables บน Render
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
FILE_IDS = [
    "1M7YJzGsNRTSQswyxdHHiDDXAuX1h7U4a",
    "1HwUNEZ1wwwne2-ZG0ogluODTdmAAdTSg",
]
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit?gid=0#gid=0"

HEADER_ROW = 4
CODE_B_INDEX = 1
CODE_C_INDEX = 2
DESC_COL_INDEX = 3     
NEW_COL_INDEX = 12
OLD_COL_INDEX = 13
BALANCE_COL_INDEX = 63

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
    """ อ่านค่า JSON จาก Env Var และแก้ไขการขึ้นบรรทัดใหม่ใน Private Key ให้ถูกต้อง """
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if google_creds_json:
        creds_dict = json.loads(google_creds_json)
        # แก้ไขปัญหาเครื่องหมาย \n ในรหัสลับที่มักเพี้ยนเวลาวางใน Render
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        return Credentials.from_service_account_file("credentials.json", scopes=SCOPES)

def run_stock_sync():
    print("🔒 กำลังเชื่อมต่อกับ Google Sheets API...")
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(GSHEET_URL)

    print("--- ตรวจสอบและดาวน์โหลดไฟล์จาก Google Drive ---")
    existing_files = glob.glob("*.xlsx")
    if len(existing_files) < len(FILE_IDS):
        for file_id in FILE_IDS:
            gdown.download(id=file_id, quiet=False)

    all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
    global_code_map = {}
    project_col_map = {}
    combined_headers = []

    for file_path in all_xlsx:
        excel_file_obj = pd.ExcelFile(file_path)
        sheet_name = excel_file_obj.sheet_names[0]
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        num_cols = df_raw.shape[1]
        
        for r in range(HEADER_ROW + 1, len(df_raw)):
            for col_idx in [CODE_B_INDEX, CODE_C_INDEX]:
                c_val = df_raw.iloc[r, col_idx]
                if pd.notna(c_val):
                    norm_c = normalize_code(c_val)
                    if norm_c and norm_c not in ["NAN", "NONE", "0", "CODE", "รหัสสินค้า", "รหัส", "NO"]:
                        if norm_c not in global_code_map:
                            global_code_map[norm_c] = (df_raw, r)
                            
        for c in range(num_cols):
            txts = [str(df_raw.iloc[r, c]).strip() for r in range(HEADER_ROW, min(HEADER_ROW + 3, len(df_raw))) if pd.notna(df_raw.iloc[r, c])]
            combined_headers.append(" ".join(txts))

    ws_input = spreadsheet.worksheet("Input_Order")
    df_input = pd.DataFrame(ws_input.get_all_records())
    excluded_kw = ["no", "code", "รายการ", "order", "category", "weight", "quantity", "qty", "on hand", "balance", "ขาด", "old", "new", "broken", "po", "repair", "total", "dift"]

    sample_df = all_xlsx and pd.read_excel(all_xlsx[0], sheet_name=0, header=None, engine='openpyxl')
    num_cols_sample = sample_df.shape[1] if sample_df is not None else 0

    for c in range(num_cols_sample):
        if c in [CODE_B_INDEX, CODE_C_INDEX, DESC_COL_INDEX, NEW_COL_INDEX, OLD_COL_INDEX, BALANCE_COL_INDEX]: continue
        h_text = combined_headers[c].lower() if c < len(combined_headers) else ""
        if any(k in h_text for k in excluded_kw): continue
        
        if c < len(combined_headers):
            pcode_match = re.search(r'\b([A-Z0-9]{2,6})\b', combined_headers[c].upper())
            if pcode_match:
                pcode = pcode_match.group(1)
                if pcode not in ["NEW", "OLD", "QTY", "KG", "TOTAL", "PO", "BROKEN", "REPAIR"]:
                    if pcode not in project_col_map: project_col_map[pcode] = []
                    project_col_map[pcode].append(c)

    report_items = []
    found_count = 0
    input_code_col = next((c for c in df_input.columns if "code" in str(c).lower() or "รหัส" in str(c)), df_input.columns[0])
    input_qty_col = next((c for c in df_input.columns if "qty" in str(c).lower() or "จำนวน" in str(c)), df_input.columns[1])

    for _, row in df_input.iterrows():
        raw_code = str(row[input_code_col]).strip()
        norm_input = normalize_code(raw_code)
        qty_needed = clean_num(row[input_qty_col])
        
        match_data = global_code_map.get(norm_input)
        if match_data is not None:
            found_count += 1
            df_target, target_row = match_data
            balance = clean_num(df_target.iloc[target_row, BALANCE_COL_INDEX])
            shortage = qty_needed if balance < 0 else max(0, qty_needed - balance)
            
            report_items.append({
                'code': str(df_target.iloc[target_row, CODE_B_INDEX]), 
                'desc': str(df_target.iloc[target_row, DESC_COL_INDEX]),
                'shortage': "มีของ" if shortage == 0 else int(shortage),
                'balance': int(balance)
            })
        else:
            report_items.append({
                'code': raw_code, 
                'desc': "ไม่พบรหัส",
                'shortage': int(qty_needed),
                'balance': "-"
            })

    header = ["รหัสสินค้า", "รายการ", "สต็อก Balance", "ขาด"]
    table_data = [header] + [[i['code'], i['desc'], i['balance'], i['shortage']] for i in report_items]

    ws_summary = spreadsheet.worksheet("Summary") if "Summary" in [w.title for w in spreadsheet.worksheets()] else spreadsheet.add_worksheet(title="Summary", rows="100", cols="30")
    ws_summary.clear()
    ws_summary.update(table_data, 'A1')
    return found_count

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_msg = event.message.text.strip()
    if user_msg in ["อัปเดตสต็อก", "sync"]:
        try:
            count = run_stock_sync()
            reply_text = f"✅ อัปเดตข้อมูลสต็อกสำเร็จ! ประมวลผลสำเร็จ {count} รายการ"
        except Exception as e:
            reply_text = f"❌ เกิดข้อผิดพลาด: {str(e)}"
    else:
        reply_text = f"พิมพ์คำว่า 'อัปเดตสต็อก' เพื่อสั่งประมวลผลข้อมูลผ่านคลาวด์ได้เลยครับ"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
