import hashlib

class LicenseManager:
    SECRET_SALT = "EASYMATCH_SECURE_SALT_2024"

    @staticmethod
    def generate_key(machine_id):
        raw = f"{machine_id}-{LicenseManager.SECRET_SALT}"
        full_hash = hashlib.sha256(raw.encode()).hexdigest().upper()
        k = full_hash[:16]
        return f"{k[0:4]}-{k[4:8]}-{k[8:12]}-{k[12:16]}"

mid = "D77CC6FA44A050D8"
print(LicenseManager.generate_key(mid))
