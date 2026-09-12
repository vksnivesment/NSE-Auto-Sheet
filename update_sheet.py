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

# 1. Credentials Setup (GitHub Secrets से लोड करना)
creds_json = os.environ.get('GCP_CREDENTIALS')
if not creds_json:
    print("ERROR: GCP_CREDENTIALS secret missing!")
    exit(1)

creds_dict = json.loads(creds_json)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# आपकी Google Sheet की ID 
spreadsheet_id = "1-KDhV_vasXJeA96E7NRFGpaLlhZYkswzJPkp4mVbVGU"

# Google Sheet वर्कबुक खोलना
try:
    workbook = client.open_by_key(spreadsheet_id)
    ws_volume = workbook.worksheet("Top 250 Stocks")
    ws_turnover = workbook.worksheet("Top 250 Turnover")
except Exception as e:
    print(f"Sheet Connection Error: {e}")
    exit(1)

# 2. NSE UDiFF Bhavcopy Data Fetcher
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
            print("फाइल मिल गई! डेटा प्रोसेस हो रहा है...")
            with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                csv_filename = z.namelist()[0]
                with z.open(csv_filename) as f:
                    df = pd.read_csv(f)
                    
                    sym_col = 'TckrSymb' if 'TckrSymb' in df.columns else 'SYMBOL'
                    close_col = 'ClsPric' if 'ClsPric' in df.columns else 'CLOSE'
                    series_col = 'SctySrs' if 'SctySrs' in df.columns else 'SERIES'
                    
                    # वॉल्यूम कॉलम ढूँढना
                    vol_col = 'TtlTradgVol'
                    for c in ['TtlTradgVol', 'TOTTRDQTY', 'TtlTrdQty', 'TotTrdQty']:
                        if c in df.columns:
                            vol_col = c
                            break
                            
                    # टर्नओवर कॉलम ढूँढना
                    turnover_col = 'TtlTrfVal'
                    for c in ['TtlTrfVal', 'TOTTRDVAL', 'TtlTrdVal', 'TotTrdVal']:
                        if c in df.columns:
                            turnover_col = c
                            break
                    
                    # केवल इक्विटी (EQ) सीरीज फ़िल्टर करना
                    if series_col in df.columns:
                        df = df[df[series_col].astype(str).str.strip() == 'EQ']
                    
                    # ETF, GOLD, LIQUID आदि हटाना
                    filter_keywords = 'BEES|ETF|GOLD|LIQUID|CASE|SILVER|LIQ'
                    df = df[~df[sym_col].astype(str).str.contains(filter_keywords, case=False, na=False)]
                    
                    # लिस्ट A: टॉप 250 वॉल्यूम
                    df_vol = df.sort_values(by=vol_col, ascending=False).head(250)
                    data_vol = df_vol[[sym_col, vol_col, close_col]].values.tolist()
                    
                    # लिस्ट B: टॉप 250 टर्नओवर
                    df_turnover = df.sort_values(by=turnover_col, ascending=False).head(250)
                    data_turnover = df_turnover[[sym_col, turnover_col, close_col]].values.tolist()
                    
                    return data_vol, data_turnover
        else:
            print(f"NSE सर्वर ने कोड {response.status_code} दिया।")
            return None, None
    except Exception as e:
        print(f"Error: {e}")
        return None, None

# सुरक्षित तरीके से शीट डेटा पढ़ने का फ़ंक्शन (बिना किसी हेडर एरर के)
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
            if header:  # खाली हेडर को छोड़ दें
                row_dict[header] = row[i] if i < len(row) else ""
        records.append(row_dict)
    return records

# 3. पिछले 7 दिनों में से ताज़ा डेटा खोजना
date = datetime.now()
data_vol_to_insert = None
data_turnover_to_insert = None
fetched_date_str = ""

for i in range(7):
    test_date = date - timedelta(days=i)
    if test_date.weekday() >= 5: # शनिवार और रविवार छोड़ना
        continue
        
    data_vol, data_turnover = fetch_bhavcopy_for_date(test_date)
    if data_vol and data_turnover:
        data_vol_to_insert = data_vol
        data_turnover_to_insert = data_turnover
        fetched_date_str = test_date.strftime('%d-%b-%Y')
        break

# 4. Sheets अपडेट करना और CDN के लिए stocks.json तैयार करना
if data_vol_to_insert and data_turnover_to_insert:
    try:
        # A. कच्चा डेटा दोनों इनपुट शीट्स में डालना
        ws_volume.batch_clear(['A2:C251'])
        ws_volume.update(values=data_vol_to_insert, range_name='A2')
        
        ws_turnover.batch_clear(['A2:C251'])
        ws_turnover.update(values=data_turnover_to_insert, range_name='A2')
        
        # टाइमस्टैम्प बनाना और K2 में डालना
        ist_now = (datetime.utcnow() + timedelta(hours=5, minutes=30)).strftime('%d-%b %H:%M')
        status_msg = f"Data Date: {fetched_date_str} | Last Update: {ist_now} (IST)"
        
        ws_volume.update(values=[[status_msg]], range_name='K2')
        ws_turnover.update(values=[[status_msg]], range_name='K2')
        
        # फ़ॉर्मूलों (200 DMA, CAR) को पूरा लोड होने के लिए 30 सेकंड का समय दें
        print("Google Sheet अपडेट हो गई। सभी 200 DMA फॉर्मूलों के कैलकुलेशन के लिए 30 सेकंड रुक रहे हैं...")
        time.sleep(30)
        
        # B. दोनों Final List शीट्स से फ़िल्टर किया हुआ डेटा निकालना
        ws_final_turnover = workbook.worksheet("Final List Turnover")
        ws_final_volume = workbook.worksheet("Final List Volume")
        
        turnover_records = get_clean_records(ws_final_turnover)
        volume_records = get_clean_records(ws_final_volume)
        
        # C. JSON फाइल तैयार करना (CDN के लिए)
        final_json_data = {
            "status": "success",
            "last_updated": status_msg,
            "data_date": fetched_date_str,
            "final_turnover": turnover_records,
            "final_volume": volume_records
        }
        
        with open('stocks.json', 'w', encoding='utf-8') as f:
            json.dump(final_json_data, f, ensure_ascii=False, indent=2)
            
        print(f"SUCCESS: {len(turnover_records)} टर्नओवर और {len(volume_records)} वॉल्यूम स्टॉक्स stocks.json में सेव हो गए!")
    except Exception as e:
        print(f"डेटा प्रोसेस/सेव करने में एरर: {e}")
        exit(1)
else:
    print("FAILED: पिछले 7 दिनों में से किसी भी दिन की फाइल नहीं मिली।")
    exit(1)
