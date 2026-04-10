import requests
import base64
import os
import re
from datetime import datetime

class GitHubSync:
    @staticmethod
    def extract_repo_info(url):
        """
        Extract owner and repo name from GitHub URL.
        Example: https://github.com/bough38-web/Enterprise-Data-Suite.git
        -> ('bough38-web', 'Enterprise-Data-Suite')
        """
        if not url: return None, None
        match = re.search(r"github\.com/([^/]+)/([^/.]+)", url)
        if match:
            return match.group(1), match.group(2)
        return None, None

    @staticmethod
    def upload_file(token, repo_url, local_path, commit_message=None):
        """
        Uploads a local file to the GitHub repository using REST API.
        Stored in: uploads/YYYY-MM-DD/filename
        """
        owner, repo = GitHubSync.extract_repo_info(repo_url)
        if not owner or not repo:
            return False, "유효한 GitHub URL을 찾을 수 없습니다."

        if not token:
            return False, "GitHub 토큰이 설정되어 있지 않습니다."

        filename = os.path.basename(local_path)
        date_str = datetime.now().strftime("%Y-%m-%d")
        remote_path = f"uploads/{date_str}/{filename}"
        
        try:
            with open(local_path, "rb") as f:
                content = base64.b64encode(f.read()).decode("utf-8")
            
            api_url = f"https://api.github.com/repos/{owner}/{repo}/contents/{remote_path}"
            headers = {
                "Authorization": f"token {token}",
                "Accept": "application/vnd.github.v3+json"
            }
            data = {
                "message": commit_message or f"Upload result: {filename}",
                "content": content,
                "branch": "main" # Can be made dynamic if needed
            }
            
            response = requests.put(api_url, headers=headers, json=data, timeout=30)
            
            if response.status_code in [201, 200]:
                return True, f"성공적으로 업로드되었습니다: {remote_path}"
            else:
                error_data = response.json()
                msg = error_data.get("message", "Unknown error")
                return False, f"업로드 실패 (HTTP {response.status_code}): {msg}"
                
        except Exception as e:
            return False, f"업로드 중 오류 발생: {str(e)}"
