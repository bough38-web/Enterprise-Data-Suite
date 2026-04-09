import os
import subprocess
import sys

def build_exe():
    print("[GO] 데이터 매니지먼트 스위트 - Windows EXE 빌드 도우미")
    
    # 1. 패키지 설치 확인
    try:
        import PyInstaller
    except ImportError:
        print("[TIP] PyInstaller가 설치되지 않았습니다. 설치를 진행합니다...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # 2. PyInstaller 명령어 구성
    # --noconsole: GUI 전용 (콘솔창 없음)
    # --onefile: 단일 파일로 압축
    # --collect-submodules: 필요한 서브 모듈 포함
    # --hidden-import: 혹시 누락될 수 있는 라이브러리 추가
    
    cmd = [
        "pyinstaller",
        "--noconsole",
        "--onefile",
        "--name", "데이터추출관리프로그램",
        "--add-data", f"utils{os.pathsep}utils",
        "--add-data", f"ui{os.pathsep}ui",
        "--hidden-import", "pandas",
        "--hidden-import", "xlwings",
        "--hidden-import", "openpyxl",
        "app.py"
    ]

    print(f"[PKG] 빌드 시작 (명령어: {' '.join(cmd)})")
    
    try:
        subprocess.check_call(cmd)
        print("\n✅ 빌드가 완료되었습니다!")
        print("[DIR] 'dist' 폴더 안에 '데이터추출관리프로그램.exe' 파일이 생성되었습니다.")
    except Exception as e:
        print(f"\n❌ 빌드 중 오류가 발생했습니다: {e}")
        print("[TIP] 수동 빌드 방법: pyinstaller --noconsole --onefile app.py")

if __name__ == "__main__":
    if sys.platform != "win32":
        print("⚠️  이 스크립트는 Windows OS에서 실행해야 .exe 파일이 만들어집니다.")
    build_exe()
