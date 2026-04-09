import requests
import threading
import json
import datetime
from utils.license_manager import LicenseManager

class TelemetryManager:
    @staticmethod
    def log_event(url, event_type, details=None):
        """Send a telemetry event to the remote webhook asynchronously."""
        if not url: return

        def task():
            try:
                payload = {
                    "mid": LicenseManager.get_machine_id(),
                    "timestamp": datetime.datetime.now().isoformat(),
                    "event": event_type,
                    "details": details or {}
                }
                # Silent post with timeout
                requests.post(url, data=json.dumps(payload), headers={'Content-Type': 'application/json'}, timeout=5)
            except Exception:
                # Silently fail to not disturb user
                pass

        threading.Thread(target=task, daemon=True).start()

    @staticmethod
    def test_ping(url):
        """Test connection to the webhook (synchronous)."""
        try:
            payload = {
                "mid": "TEST_ID",
                "timestamp": datetime.datetime.now().isoformat(),
                "event": "TEST_PING",
                "details": {"msg": "Checking connection"}
            }
            r = requests.post(url, data=json.dumps(payload), headers={'Content-Type': 'application/json'}, timeout=5)
            return r.status_code == 200
        except Exception as e:
            raise Exception(f"연결 실패: {e}")
