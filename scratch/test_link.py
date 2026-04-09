import requests
import json
import datetime

url = "https://script.google.com/macros/s/AKfycbx10peXFFkDrBkWwf_qJIFO_QDVHie8Aq01LA6Sd5PXCmdQSao8lwKj_3LuJQdYU8LV/exec"
payload = {
    "mid": "DEVELOPER_TEST",
    "timestamp": datetime.datetime.now().isoformat(),
    "event": "TELEMETRY_LINKED",
    "details": {"msg": "EasyMatch Pro has been successfully linked to your monitoring sheet."}
}

try:
    r = requests.post(url, data=json.dumps(payload), headers={'Content-Type': 'application/json'}, timeout=10)
    print(f"Status: {r.status_code}")
    print(f"Response: {r.text}")
except Exception as e:
    print(f"Error: {e}")
