# 데이터 매니지먼트 스위트 EXE 빌드 스크립트 (PowerShell)

Write-Host "[GO] Windows EXE 빌드를 시작합니다 (PowerShell 버전)..." -ForegroundColor Cyan

# 1. PyInstaller 가 있는지 확인
python -m pip install -r requirements.txt pyinstaller --quiet

# 2. 빌드 캐시 및 이전 dist 폴더 삭제 (선택)
if (Test-Path ".\dist") { Remove-Item -Recurse -Force ".\dist" }
if (Test-Path ".\build") { Remove-Item -Recurse -Force ".\build" }

# 3. PyInstaller 실행
# --noconsole: GUI 전용 (콘솔창 비활성)
# --onefile: 단일 파일로 빌드
# --name: 결과 파일 이름
# --add-data: 필요한 리소스 폴더 포함
Write-Host "[PKG] 패키징 작업 중..." -ForegroundColor Yellow
pyinstaller --noconsole --onefile --name "데이터추출관리프로그램" --add-data "utils;utils" --add-data "ui;ui" --add-data "assets;assets" --hidden-import "pandas" --hidden-import "xlwings" --hidden-import "openpyxl" .\app.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n✅ 빌드가 성공적으로 완료되었습니다!" -ForegroundColor Green
    Write-Host "[DIR] 'dist' 폴더 안에 '데이터추출관리프로그램.exe'가 생성되었습니다." -ForegroundColor Green
} else {
    Write-Host "`n❌ 빌드 중 오류가 발생했습니다." -ForegroundColor Red
}

Read-Host "종료하려면 엔터를 누르세요..."
