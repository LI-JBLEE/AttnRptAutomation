# GSC Attainment Report Automator

Attainment Report를 매니저별로 자동 생성하고 Outlook 이메일 Draft를 작성/발송하는 도구입니다.

**두 가지 실행 방식 제공:**
- **🌐 웹 앱 (Streamlit)**: 클라우드에서 리포트 생성 후 다운로드
- **📧 이메일 관리자 (Windows .exe)**: 로컬에서 Outlook 이메일 발송

---

## 🌐 웹 앱: 리포트 생성

### 접속
- **로컬 실행**: `python -m streamlit run app.py`
- **클라우드 배포**: Streamlit Cloud에 배포 가능

### 사용 방법
1. **Step 1**: Global Attainment Report + Sales Compensation Report 업로드
2. **Step 2**: 출력 폴더 경로 입력 (UI 표시용, 실제로는 메모리에 생성됨)
3. **Step 3**: Region 선택 → **Generate Reports** 클릭 → **.zip 파일 다운로드**

### 출력물
- `Manager_Reports_FY26_YYYYMMDD.zip` 파일에 포함:
  - Region별 폴더로 정리된 Excel 리포트 파일들
  - `manager_metadata.json`: 매니저 정보 + 이메일 주소 매핑

---

## 📧 이메일 관리자: Outlook 이메일 발송

### 다운로드
- [EmailManager.exe](https://github.com/LI-JBLEE/AttnRptAutomation/releases/latest) (릴리스 페이지에서 다운로드)
- 설치 불필요 - 다운로드 후 바로 실행
- 요구사항: Windows 10/11 + Outlook 설치

### 사용 방법
1. **Step 1 — Load Reports**:
   - **📂 Load .zip File**: 웹 앱에서 다운로드한 .zip 파일 선택
   - 또는 **📁 Load Folder**: 로컬에서 생성한 리포트 폴더 선택

2. **Step 2 — Select Recipients**:
   - Region 체크박스로 필터링
   - 매니저 목록에서 선택 (✓ = 이메일 매칭 성공, ✗ = 이메일 없음)
   - **✅ Select All** / **❌ Deselect All** 버튼 사용

3. **Step 3 — Email Operations**:
   - **Tab 1: Create Drafts**
     - **📧 Create Outlook Drafts** 클릭
     - Outlook > Drafts > Manager Report 폴더에 Draft 생성됨
   - **Tab 2: Send Drafts**
     - **🔄 Load Drafts** 클릭하여 Outlook에서 Draft 목록 로드
     - 전송할 Draft 선택
     - **✉️ Send Selected** 클릭

---

## 🔄 전체 워크플로우

```
1. 🌐 웹 앱 접속
   ↓ 파일 업로드 → Region 선택 → Generate
   ↓ Manager_Reports_FY26_20260213.zip 다운로드

2. 💾 로컬 PC에 .zip 파일 저장

3. 📧 EmailManager.exe 실행
   ↓ .zip 파일 로드
   ↓ 매니저 선택
   ↓ Outlook Draft 생성
   ↓ Draft 확인 후 선택 발송

4. ✅ 이메일 전송 완료!
```

---

## 🚀 개발자용: 로컬 환경 설정

### 1. Git 설치
- Windows용 Git 다운로드: https://git-scm.com/download/win

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

### 4. 실행
```bash
# 웹 앱
python -m streamlit run app.py

# 이메일 관리자 (GUI)
python email_manager.py
```

### 5. EmailManager.exe 빌드
```bash
pip install pyinstaller
pyinstaller email_manager.spec
# Output: dist/EmailManager.exe
```

---

## 📁 프로젝트 구조

```
AttnRptAutomation/
├── app.py                      # Streamlit 웹 UI (Steps 1-3)
├── email_manager.py            # Tkinter GUI 이메일 관리자
├── email_manager.spec          # PyInstaller 설정 파일
├── generate_manager_reports.py # 리포트 생성 엔진
├── create_email_drafts.py      # Outlook 이메일 생성/발송
├── requirements.txt            # Python 패키지 목록
└── README.md                   # 이 파일
```

---

## ⚠️ 주의사항

- **이메일 관리자 (.exe)는 Windows 전용**입니다 (Outlook COM 사용)
- Outlook이 설치되어 있고 메일 계정이 설정되어 있어야 합니다
- 매니저 리포트는 자동으로 Region별 폴더에 저장됩니다
- Fiscal Year는 Attainment 파일의 "Fiscal Year" 컬럼에서 자동 감지됩니다 (FY26, FY27 등)

---

## 🔧 문제 해결

### 웹 앱 관련

**"streamlit: command not found"**
```bash
python -m streamlit run app.py
```

**Pandas 경고 메시지 (openpyxl)**
- 무시해도 됩니다. 파일은 정상적으로 생성됩니다.

### 이메일 관리자 관련

**Outlook 연결 오류**
- Outlook이 실행 중인지 확인
- pywin32 재설치: `pip install --upgrade pywin32`

**"Manager Report" 폴더를 찾을 수 없음**
- 한 번이라도 Draft를 생성하면 자동으로 폴더가 만들어집니다

**.zip 파일이 비어있음**
- 웹 앱에서 리포트를 먼저 생성했는지 확인
- 브라우저 다운로드 폴더 확인

---

## 📊 장점

✅ **클라우드 접근성**: 어디서나 웹 브라우저로 리포트 생성
✅ **Outlook 완전 통합**: Windows에서 로컬 COM 객체로 안전한 이메일 발송
✅ **간편한 배포**: .exe 파일 한 번 다운로드로 끝
✅ **오프라인 작업**: 이메일 관리자는 인터넷 없이도 동작
✅ **보안**: 이메일 주소는 클라우드에 업로드되지 않고 로컬 PC에만 존재
✅ **친숙한 워크플로우**: 기존 6단계 프로세스를 2개 도구로 분리

---

## 📞 문의

이슈 발생 시: https://github.com/LI-JBLEE/AttnRptAutomation/issues
