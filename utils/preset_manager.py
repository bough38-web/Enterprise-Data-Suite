import json
import os
from pathlib import Path
from utils.excel_io import ExcelHandler

class PresetManager:
    def __init__(self, presets_path):
        self.presets_path = Path(presets_path)
        self._ensure_file_exists()

    def _ensure_file_exists(self):
        if not self.presets_path.exists():
            with open(self.presets_path, 'w', encoding='utf-8') as f:
                json.dump({}, f, indent=4, ensure_ascii=False)

    def load_all(self):
        """Load all presets from the local JSON file."""
        if not self.presets_path.exists():
            return {}
        try:
            with open(self.presets_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def save_all(self, presets):
        """Save presets dictionary to the local JSON file."""
        with open(self.presets_path, 'w', encoding='utf-8') as f:
            json.dump(presets, f, indent=4, ensure_ascii=False)

    def add_preset(self, name, data):
        """Add or update a single preset."""
        presets = self.load_all()
        presets[name] = data
        self.save_all(presets)

    def delete_preset(self, name):
        """Delete a preset by name."""
        presets = self.load_all()
        if name in presets:
            del presets[name]
            self.save_all(presets)
            return True
        return False

    def sync_from_remote(self, url, token=None):
        """
        Fetch presets from a remote URL and merge them into the local file.
        Remote presets overwrite local ones with the same name.
        Returns: (success, message, count)
        """
        if not url:
            return False, "동기화 URL이 설정되지 않았습니다.", 0

        try:
            remote_presets = ExcelHandler.fetch_json_from_url(url, token)
            if not isinstance(remote_presets, dict):
                return False, "유효하지 않은 프리셋 형식입니다 (JSON 객체여야 함).", 0

            local_presets = self.load_all()
            
            # Merge: Remote overwrites local
            count = len(remote_presets)
            local_presets.update(remote_presets)
            
            self.save_all(local_presets)
            return True, f"성공적으로 {count}개의 프리셋을 불러왔습니다.", count
        except Exception as e:
            return False, f"동기화 중 오류 발생: {str(e)}", 0
