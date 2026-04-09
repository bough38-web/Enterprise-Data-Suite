# Git LFS (Large File Storage) 초기화 및 데이터 트래킹 설정 스크립트

echo "🚀 Git LFS 설정을 시작합니다..."

# 1. Git LFS 설치 확인 및 초기화
try {
    git lfs install
    echo "✅ Git LFS가 초기화되었습니다."
} catch {
    echo "❌ Git LFS가 설치되어 있지 않거나 실행할 수 없습니다."
    echo "지침: 윈도우용 Git이 설치되어 있어야 하며, 설치 시 LFS 옵션을 체크했는지 확인하세요."
    exit
}

# 2. 대용량 파일 타입 추적 시작
echo "📂 데이터 파일 추적을 시작합니다 (*.xlsx, *.csv)..."
git lfs track "*.xlsx"
git lfs track "*.csv"

# 3. 설정 파일 반영
echo "📝 .gitattributes 파일을 추가합니다..."
git add .gitattributes

echo ""
echo "--------------------------------------------------------"
echo "✅ 설정 완료!"
echo "이제 100MB가 넘는 대용량 데이터도 깃허브에 올릴 수 있습니다."
echo "평소처럼 아래 명령어로 커밋하고 푸시하세요:"
echo "git add 데이터파일.csv"
echo "git commit -m 'Add large dataset'"
echo "git push origin main"
echo "--------------------------------------------------------"
