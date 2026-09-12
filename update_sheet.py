import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import requests
import zipfile
import io
from datetime import datetime, timedelta
import os
import json

# 1. Credentials Setup
creds_json = os.environ.get('GCP_CREDENTIALS')
if not creds_json:
    print("ERROR: GCP_CREDENTIALS secret missing!")
    exit(1)

creds_dict = json.loads(creds_json)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# आपकी शीट की ID 
spreadsheet_id = "1-KDhV_vasXJeA96E7NRFGpaLlhZYkswzJPkp4mVbVGU"

# दोनों शीट्स को कनेक्ट करना
try:
    ws_volume = client.open_by_key(spreadsheet_id).worksheet("Top 250 Stocks")
    ws_turnover = client.open_by_key(spreadsheet_id).worksheet("Top 250 Turnover")
except Exception as e:
    print(f"Sheet Connection Error: {e}")
    exit(1)

# 2. New NSE UDiFF Data Fetcher
def fetch_bhavcopy_for_date(date_obj):
    date_str = date_obj.strftime("%Y%m%d")
    url = f"https://nsearchives.nseindia.com/content/cm/BhavCopy_NSE_CM_0_0_0_{date_str}_F_0000.csv.zip"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    print(f"--- तारीख {date_obj.strftime('%d-%m-%Y')} चेक कर रहे हैं ---")
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            print("फाइल मिल गई! अब इसे खोल रहे हैं...")
            with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                csv_filename = z.namelist()[0]
                with z.open(csv_filename) as f:
                    df = pd.read_csv(f)
                    
                    sym_col = 'TckrSymb' if 'TckrSymb' in df.columns else 'SYMBOL'
                    close_col = 'ClsPric' if 'ClsPric' in df.columns else 'CLOSE'
                    series_col = 'SctySrs' if 'SctySrs' in df.columns else 'SERIES'
                    
                    vol_col = 'TtlTradgVol'
                    for c in ['TtlTradgVol', 'TOTTRDQTY', 'TtlTrdQty', 'TotTrdQty']:
                        if c in df.columns:
                            vol_col = c
                            break
                            
                    turnover_col = 'TtlTrfVal'
                    for c in ['TtlTrfVal', 'TOTTRDVAL', 'TtlTrdVal', 'TotTrdVal']:
                        if c in df.columns:
                            turnover_col = c
                            break
                    
                    if series_col in df.columns:
                        df = df[df[series_col].astype(str).str.strip() == 'EQ']
                    
                    filter_keywords = 'BEES|ETF|GOLD|LIQUID|CASE|SILVER|LIQ'
                    df = df[~df[sym_col].astype(str).str.contains(filter_keywords, case=False, na=False)]
                    
                    # लिस्ट A: वॉल्यूम टॉप 250
                    df_vol = df.sort_values(by=vol_col, ascending=False).head(250)
                    data_vol = df_vol[[sym_col, vol_col, close_col]].values.tolist()
                    
                    # लिस्ट B: टर्नओवर टॉप 250
                    df_turnover = df.sort_values(by=turnover_col, ascending=False).head(250)
                    data_turnover = df_turnover[[sym_col, turnover_col, close_col]].values.tolist()
                    
                    return data_vol, data_turnover
        else:
            print(f"NSE सर्वर ने {response.status_code} रिस्पॉन्स दिया।")
            return None, None
    except Exception as e:
        print(f"Error: {e}")
        return None, None

# 3. Execution Logic
date = datetime.now()
data_vol_to_insert = None
data_turnover_to_insert = None
fetched_date_str = ""

for i in range(7):
    test_date = date - timedelta(days=i)
    if test_date.weekday() >= 5:
        continue
        
    data_vol, data_turnover = fetch_bhavcopy_for_date(test_date)
    if data_vol and data_turnover:
        data_vol_to_insert = data_vol
        data_turnover_to_insert = data_turnover
        fetched_date_str = test_date.strftime('%d-%b-%Y')
        break

# 4. Google Sheets और JSON दोनों में डेटा सेव करना
if data_vol_to_insert and data_turnover_to_insert:
    try:
        # A. Google Sheet अपडेट करें
        ws_volume.batch_clear(['A2:C251'])
        ws_volume.update('A2', data_vol_to_insert)
        
        ws_turnover.batch_clear(['A2:C251'])
        ws_turnover.update('A2', data_turnover_to_insert)
        
        ist_now = (datetime.utcnow() + timedelta(hours=5, minutes=30)).strftime('%d-%b %H:%M')
        status_msg = f"Data Date: {fetched_date_str} | Last Update: {ist_now} (IST)"
        
        ws_volume.update('K2', [[status_msg]])
        ws_turnover.update('K2', [[status_msg]])
        
        # B. हाई-स्पीड CDN के लिए 'stocks.json' फाइल बनाना
        json_output = {
            "status": "success",
            "last_updated": status_msg,
            "data_date": fetched_date_str,
            "update_time_ist": ist_now,
            "top_volume": [
                {"symbol": str(r[0]), "volume": r[1], "close": r[2]} for r in data_vol_to_insert
            ],
            "top_turnover": [
                {"symbol": str(r[0]), "turnover": r[1], "close": r[2]} for r in data_turnover_to_insert
            ]
        }
        
        with open('stocks.json', 'w', encoding='utf-8') as f:
            json.dump(json_output, f, ensure_ascii=False, indent=2)
            
        print("SUCCESS: Google Sheets और stocks.json दोनों सफलतापूर्वक अपडेट हो गए!")
    except Exception as e:
        print(f"अपडेट करने में एरर: {e}")
        exit(1)
else:
    print("FAILED: पिछले 7 दिनों में से किसी भी दिन की फाइल नहीं मिली।")
    exit(1)
