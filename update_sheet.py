import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import requests
import zipfile
import io
from datetime import datetime, timedelta
import os
import json
import time

# 1. Credentials Setup
creds_json = os.environ.get('GCP_CREDENTIALS')
if not creds_json:
    print("ERROR: GCP_CREDENTIALS secret missing!")
    exit(1)

creds_dict = json.loads(creds_json)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

spreadsheet_id = "1-KDhV_vasXJeA96E7NRFGpaLlhZYkswzJPkp4mVbVGU"

try:
    workbook = client.open_by_key(spreadsheet_id)
    ws_volume = workbook.worksheet("Top 250 Stocks")
    ws_turnover = workbook.worksheet("Top 250 Turnover")
except Exception as e:
    print(f"Sheet Connection Error: {e}")
    exit(1)

# 2. Bhavcopy Fetcher
def fetch_bhavcopy_for_date(date_obj):
    date_str = date_obj.strftime("%Y%m%d")
    url = f"https://nsearchives.nseindia.com/content/cm/BhavCopy_NSE_CM_0_0_0_{date_str}_F_0000.csv.zip"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    print(f"--- तारीख {date_obj.strftime('%d-%m-%Y')} चेक कर रहे हैं ---")
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            print("फाइल मिल गई! डेटा प्रोसेस हो रहा है...")
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
                    
                    df_vol = df.sort_values(by=vol_col, ascending=False).head(250)
                    data_vol = df_vol[[sym_col, vol_col, close_col]].values.tolist()
                    
                    df_turnover = df.sort_values(by=turnover_col, ascending=False).head(250)
                    data_turnover = df_turnover[[sym_col, turnover_col, close_col]].values.tolist()
                    
                    return data_vol, data_turnover
        return None, None
    except Exception as e:
        print(f"Error: {e}")
        return None, None

# सुरक्षित डेटा रिकॉर्ड फ़ंक्शन
def get_clean_records(worksheet):
    data = worksheet.get_all_values()
    if not data or len(data) < 2:
        return []
    headers = [str(h).strip() for h in data[0]]
    records = []
    for row in data[1:]:
        if not any(str(cell).strip() for cell in row):
            continue
        row_dict = {}
        for i, header in enumerate(headers):
            if header:
                row_dict[header] = row[i] if i < len(row) else ""
        records.append(row_dict)
    return records

# 3. Bhavcopy सर्च
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

# 4. Sheets अपडेट और stocks.json निर्माण
if data_vol_to_insert and data_turnover_to_insert:
    try:
        # A. बिना शीट क्लियर किए सीधे डेटा ओवरराइट करना (फ़ॉर्मूले सुरक्षित रहेंगे)
        ws_volume.update(values=data_vol_to_insert, range_name='A2')
        ws_turnover.update(values=data_turnover_to_insert, range_name='A2')
        
        ist_now = (datetime.utcnow() + timedelta(hours=5, minutes=30)).strftime('%d-%b %H:%M')
        status_msg = f"Data Date: {fetched_date_str} | Last Update: {ist_now} (IST)"
        
        ws_volume.update(values=[[status_msg]], range_name='K2')
        ws_turnover.update(values=[[status_msg]], range_name='K2')
        
        print("Google Sheet अपडेट हो गई। फॉर्मूलों के कैलकुलेशन के लिए 20 सेकंड रुक रहे हैं...")
        time.sleep(20)
        
        # B. फ़िल्टर शीट्स से डेटा निकालना
        ws_final_turnover = workbook.worksheet("Final List Turnover")
        ws_final_volume = workbook.worksheet("Final List Volume")
        
        turnover_records = get_clean_records(ws_final_turnover)
        volume_records = get_clean_records(ws_final_volume)
        
        print(f"चेक: शीट से {len(turnover_records)} टर्नओवर और {len(volume_records)} वॉल्यूम रिकॉर्ड मिले।")
        
        # C. सेफ़गार्ड: अगर शीट से सही रिकॉर्ड्स मिले हैं तभी फ़ाइल अपडेट करें
        if len(turnover_records) > 0 or len(volume_records) > 0:
            final_json_data = {
                "status": "success",
                "last_updated": status_msg,
                "data_date": fetched_date_str,
                "final_turnover": turnover_records,
                "final_volume": volume_records
            }
            with open('stocks.json', 'w', encoding='utf-8') as f:
                json.dump(final_json_data, f, ensure_ascii=False, indent=2)
            print(f"SUCCESS: {len(turnover_records)} टर्नओवर रिकॉर्ड्स stocks.json में सेव हो गए!")
        else:
            print("चेतावनी: शीट अभी कैलकुलेट हो रही थी, 0 रिकॉर्ड मिले। पुरानी stocks.json सुरक्षित रखी गई है।")
            
    except Exception as e:
        print(f"डेटा प्रोसेस/सेव करने में एरर: {e}")
        exit(1)
else:
    print("FAILED: Bhavcopy फाइल नहीं मिली।")
    exit(1)
