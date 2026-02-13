# GSC Attainment Report Automator

Streamlit 기반 웹 UI로 FY26 Attainment Report를 매니저별로 자동 생성하고 Outlook 이메일 Draft를 작성/발송하는 도구입니다.

## 📋 주요 기능

- **Step 1-2**: Excel 파일 업로드 및 출력 폴더 선택
- **Step 3**: 매니저별 Attainment Report 자동 생성 (Region 필터 지원)
- **Step 4-5**: Outlook 이메일 Draft 일괄 생성
- **Step 6**: Draft 이메일 선택 및 일괄 발송

## 🚀 초기 설치 (1회만)

### 1. Git 설치
- Windows용 Git 다운로드: https://git-scm.com/download/win
- 설치 후 Git Bash 또는 Command Prompt 실행

### 2. 리포지토리 클론
```bash
cd C:\
git clone https://github.com/LI-JBLEE/AttnRptAutomation.git
cd AttnRptAutomation
```

### 3. Python 패키지 설치
```bash
pip install -r requirements.txt
```

## ▶️ 실행 방법

### 방법 1: 자동 실행 스크립트 (추천)
1. `run_app.bat` 파일을 더블클릭
2. 자동으로 GitHub에서 최신 코드를 다운로드하고 앱 실행

### 방법 2: 수동 실행
```bash
cd C:\AttnRptAutomation
git pull origin main
python -m streamlit run app.py
```

## 🔄 자동 업데이트

`run_app.bat` 스크립트를 사용하면 매번 실행 시 자동으로 GitHub에서 최신 코드를 받아옵니다.

개발자가 코드를 업데이트하면 → 다음 실행 시 자동 반영됩니다.

## 📁 폴더 구조

```
AttnRptAutomation/
├── app.py                      # Streamlit 웹 UI
├── generate_manager_reports.py # 리포트 생성 엔진
├── create_email_drafts.py      # Outlook 이메일 생성/발송
├── run_app.bat                 # 자동 실행 스크립트
├── requirements.txt            # Python 패키지 목록
└── README.md                   # 이 파일
```

## 💡 사용 방법

1. `run_app.bat` 실행
2. 브라우저에서 자동으로 앱 열림 (http://localhost:8501)
3. **Step 1**: Global Attainment Report + Sales Compensation Report 업로드
4. **Step 2**: 출력 폴더 선택 (예: `C:\Attainment Reports`)
5. **Step 3**: Region 선택 후 "Generate Reports" 클릭
6. **Step 4-5**: 이메일 받을 매니저 선택 후 Draft 생성
7. **Step 6**: Draft 확인 후 선택 발송

## ⚠️ 주의사항

- **Step 5-6** (Outlook 기능)은 Windows에서만 작동합니다
- Outlook이 설치되어 있어야 합니다
- 매니저 리포트는 자동으로 Region별 폴더에 저장됩니다

## 🔧 문제 해결

### "streamlit: command not found"
```bash
python -m streamlit run app.py
```

### Outlook 연결 오류
- Outlook이 실행 중인지 확인
- pywin32 재설치: `pip install --upgrade pywin32`

### Git pull 오류
```bash
git reset --hard origin/main
git pull origin main
```

## 📞 문의

이슈 발생 시: https://github.com/LI-JBLEE/AttnRptAutomation/issues
