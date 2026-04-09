import hashlib
import subprocess
import sys
import os

class LicenseManager:
    SECRET_SALT = "EASYMATCH_SECURE_SALT_2024"

    @staticmethod
    def get_machine_id():
        """Get a stable hardware fingerprint (Motherboard/BIOS UUID)."""
        try:
            if sys.platform == "win32":
                # Windows: Get Motherboard UUID
                cmd = "wmic csproduct get uuid"
                output = subprocess.check_output(cmd, shell=True).decode().split('\n')
                if len(output) > 1:
                    raw_id = output[1].strip()
                else:
                    raw_id = os.environ.get('COMPUTERNAME', 'UnknownWin')
            elif sys.platform == "darwin":
                # Mac: Get IOPlatformUUID
                cmd = "ioreg -rd1 -c IOPlatformExpertDevice | grep IOPlatformUUID"
                output = subprocess.check_output(cmd, shell=True).decode()
                raw_id = output.split('"')[-2]
            else:
                raw_id = "GenericLinuxID"
                
            # Hash it to make it look professional and anonymized
            return hashlib.sha256(raw_id.encode()).hexdigest()[:16].upper()
        except Exception:
            # Fallback to a secondary ID
            import uuid
            node = str(uuid.getnode())
            return hashlib.sha256(node.encode()).hexdigest()[:16].upper()

    @staticmethod
    def generate_key(machine_id):
        """Generate a valid license key for a given machine ID."""
        raw = f"{machine_id}-{LicenseManager.SECRET_SALT}"
        # We take a specific slice of the hash
        full_hash = hashlib.sha256(raw.encode()).hexdigest().upper()
        # Format as XXXX-XXXX-XXXX-XXXX
        k = full_hash[:16]
        return f"{k[0:4]}-{k[4:8]}-{k[8:12]}-{k[12:16]}"

    @staticmethod
    def verify_key(machine_id, provided_key):
        """Check if the provided key is valid for this machine ID."""
        expected = LicenseManager.generate_key(machine_id)
        return expected == provided_key.strip().upper()
