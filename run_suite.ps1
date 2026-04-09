# 데이터 매니지먼트 스위트 실행 스크립트 (PowerShell)

# 1. 가상환경이 있다면 활성화 (선택 사항)
# if (Test-Path ".\venv\Scripts\Activate.ps1") {
#     .\venv\Scripts\Activate.ps1
# }

# 2. 필수 라이브러리 설치 확인
Write-Host "[SCAN] 라이브러리 설치 여부를 확인합니다..." -ForegroundColor Cyan
pip install -r requirements.txt --quiet

# 3. 프로그램 실행
Write-Host "[GO] 프로그램을 실행합니다..." -ForegroundColor Green
python .\app.py

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ 프로그램 실행 중 오류가 발생했습니다." -ForegroundColor Red
    Read-Host "종료하려면 엔터를 누르세요..."
}
