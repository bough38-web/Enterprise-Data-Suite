import json
import os
import sys
import subprocess
import requests
import re
from pathlib import Path

class UpdateManager:
    @staticmethod
    def parse_version(version_str):
        """Extract numeric parts of version (e.g., 'v3.7 Pro' -> [3, 7])"""
        nums = re.findall(r'\d+', version_str)
        return [int(n) for n in nums]

    @staticmethod
    def is_newer(current, remote):
        """Compare two version strings."""
        c_parts = UpdateManager.parse_version(current)
        r_parts = UpdateManager.parse_version(remote)
        
        # Compare parts one by one
        for i in range(max(len(c_parts), len(r_parts))):
            c = c_parts[i] if i < len(c_parts) else 0
            r = r_parts[i] if i < len(r_parts) else 0
            if r > c: return True
            if c > r: return False
        return False

    @staticmethod
    def get_remote_manifest(url):
        """Fetch the update JSON manifest."""
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            print(f"Update check failed: {e}")
            return None

    @staticmethod
    def download_file(url, target_path, callback=None):
        """Download binary file with optional progress callback."""
        try:
            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            with open(target_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if callback and total_size > 0:
                            callback(downloaded / total_size)
            return True
        except Exception as e:
            print(f"Download failed: {e}")
            if os.path.exists(target_path):
                os.remove(target_path)
            return False

    @staticmethod
    def apply_update_windows(current_exe_path, new_exe_path):
        """
        Creates a temporary batch file to replace the running EXE.
        Must be called just before sys.exit()
        """
        bat_content = f"""@echo off
timeout /t 3 /nobreak > nul
del "{current_exe_path}"
move /y "{new_exe_path}" "{current_exe_path}"
start "" "{current_exe_path}"
del "%~f0"
"""
        bat_path = Path(current_exe_path).parent / "update_swap.bat"
        with open(bat_path, "w", encoding="cp949") as f:
            f.write(bat_content)
        
        # Run detached
        subprocess.Popen(["cmd.exe", "/c", str(bat_path)], 
                         creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0)
