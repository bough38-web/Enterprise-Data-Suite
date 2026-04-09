#!/bin/bash

# Mac 전용 독립 실행 앱 (.app) 빌드 스크립트
echo "🍎 Mac용 앱 빌드를 시작합니다..."

# 1. 필수 라이브러리 설치
echo "📦 필요한 라이브러리를 설치 중입니다..."
python3 -m pip install -r requirements.txt

# 2. PyInstaller를 이용한 빌드
# --windowed: 콘솔 창 없이 실행
# --name: 앱 이름 설정
# --clean: 캐시 삭제 후 빌드
echo "🚀 빌드 프로세스 실행 중 (1~2분 정도 소요됩니다)..."
python3 -m PyInstaller --windowed \
    --name "데이터매니지먼트스위트" \
    --clean \
    --noconfirm \
    app.py

echo "✅ 빌드가 완료되었습니다!"
echo "📂 'dist' 폴더 안에 '데이터매니지먼트스위트.app' 파일이 생성되었습니다."
echo "💡 이 파일을 응용 프로그램(Applications) 폴더로 옮겨서 사용하세요."
