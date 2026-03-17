# AIWORKS.CODE

**버전:** 1.3 | **플랫폼:** Windows PowerShell 5.1

Windows 기업 내부망 환경에서 자연어 명령으로 파일시스템, 브라우저, MS Office / HWP 등을 제어하는 로컬 AI 에이전트 CLI 도구입니다.

---

## 특징

- **인터넷 없이 동작** — 기업 내부 AI Agent API와 연동
- **자연어 명령** — 한국어/영어 자연어로 PC 작업 자동화
- **순수 PowerShell** — 외부 모듈 없음 (PS 5.1 + COM)
- **모듈형 액션** — `actions/` 폴더에 개별 파일로 분리, 신규 액션 추가 가능
- **Vector DB** — 자주 쓰는 명령을 로컬 DB에 저장하여 AI 없이 즉시 실행

---

## 디렉토리 구조

```
AIWORKS.CODE/
├── AIWORKS.code.ps1        # 메인 스크립트 (REPL, 인증, VectorDB, 디스패처)
├── AIWORKS.code.config     # 설정 파일 (INI 형식)
├── AIWORKS.code.db         # Vector DB (JSON)
├── aiworks.log             # 위험 작업 로그 (자동 생성)
└── actions/
    ├── action.answer.ps1   # 텍스트 응답
    ├── action.office.ps1   # Excel / Word / PowerPoint COM 제어
    ├── action.hwp.ps1      # 한글(HWP) COM 제어
    ├── action.ie.ps1       # Internet Explorer COM 제어
    ├── action.chrome.ps1   # Chrome 브라우저 제어
    ├── action.pdf.ps1      # PDF 텍스트 추출
    └── action.shell.ps1    # PowerShell 명령 실행
```

---

## 빠른 시작

### 1. 설정 파일 편집

`AIWORKS.code.config`의 `[AI]` 섹션에 내부 AI Agent 정보를 입력합니다:

```ini
[AI]
Provider    = custom
LoginUrl    = https://internal-ai.company.com/login
Endpoint    = https://internal-ai.company.com/api/chat
CookieField = bearer_token
Model       = internal-llm
```

### 2. 실행

```powershell
powershell -ExecutionPolicy Bypass -File AIWORKS.code.ps1
```

### 3. 자연어로 명령

```
AIWORKS> 사내 포털 사이트 열어줘
AIWORKS> 바탕화면의 보고서.xlsx 파일 열어줘
AIWORKS> 이 엑셀 파일을 PDF로 변환해줘
AIWORKS> C:\logs 폴더의 파일 목록 보여줘
```

---

## 인증 방식

최초 실행 시 `LoginUrl`로 브라우저(IE COM)를 자동으로 열어 사용자 인증을 진행합니다.
인증 성공 후 쿠키에서 `CookieField` 값을 읽어 bearer token으로 사용합니다.
토큰은 세션 메모리에만 유지되며 파일에 저장되지 않습니다.

```
[*] 인증이 필요합니다. 브라우저에서 로그인 후 Enter를 누르세요...
```

정적 토큰을 사용하려면 `AuthToken` 항목에 직접 입력하세요.

---

## 슬래시 명령어

| 명령어 | 설명 |
|--------|------|
| `/help` | 전체 명령어 목록 |
| `/history` | 대화 히스토리 출력 |
| `/clear` | 화면 초기화 |
| `/reset` | 대화 히스토리 초기화 |
| `/status` | 현재 세션 정보 |
| `/exit` | 세션 종료 |
| `/config` | 설정 파일 편집 (AI 불필요) |
| `/db` | Vector DB 관리 |

### /config 상세

```powershell
/config list                        # 전체 설정 출력
/config list AI                     # 섹션별 출력
/config get AI Endpoint             # 특정 값 조회
/config set AI TimeoutSec 120       # 값 수정
/config add-program Notepad C:\Windows\System32\notepad.exe
/config add-warning 삭제 삭제 작업은 되돌릴 수 없습니다.
/config reload                      # 파일 변경사항 즉시 반영
```

### /db 상세

```powershell
/db list                # 저장된 항목 전체 출력
/db search 포털 열어줘  # 실행 없이 검색 테스트
/db delete 0            # 인덱스로 항목 삭제
/db clear               # 전체 초기화
```

---

## Vector DB

자주 사용하는 명령을 로컬에 저장하여 AI 호출 없이 즉시 실행합니다.

```
AIWORKS> 사내 포털 사이트 열어줘
  [>] Chrome 열기: https://portal.company.com

  [?] 이 작업을 DB에 저장하시겠습니까? [Y/N] Y
  [+] Vector DB에 저장되었습니다.

-- 다음 번 --
AIWORKS> 포털 열어줘
  [DB] Vector DB 매칭: 사내 포털 사이트 열어줘
  [>] Chrome 열기: https://portal.company.com
```

---

## 지원 액션

| 액션 | 설명 | 주요 파라미터 |
|------|------|--------------|
| `answer` | 텍스트 응답 | `message` |
| `office` | Excel/Word/PPT 제어 | `app`, `action`, `path` |
| `hwp` | 한글 HWP 제어 | `action`, `path`, `contents`(new 전용), `contentsPath`(new 전용) |
| `ie` | IE 브라우저 제어 | `action`, `url`, `selector` |
| `chrome` | Chrome 브라우저 | `action`, `url` |
| `pdf` | PDF 텍스트 추출 | `path` |
| `shell` | PowerShell 실행 | `command`, `workingDir` |

### hwp contents 파라미터 (메타 태그)

`action=new` 와 함께 `contents`를 전달하면 신규 문서에 텍스트를 바로 삽입합니다.
각 줄의 맨 앞에 `[font=FaceName,size=N]` 태그를 붙이면 줄별로 폰트/크기를 지정할 수 있으며, 태그가 없으면 기본값(휴먼명조체 16pt)이 적용됩니다.

```
/run hwp action=new contents="[font=휴먼명조체,size=18]보고서 제목
[font=바탕체,size=12]본문 첫째 줄입니다.
태그 없는 줄은 기본 폰트/사이즈로 입력됩니다."
```

내용이 길 경우 임시 파일에 저장 후 `contentsPath`로 전달합니다. `contentsPath`는 `contents`보다 우선 적용됩니다:

```powershell
# 임시 파일 생성
$tmp = [IO.Path]::GetTempFileName() -replace '\.tmp$', '.txt'
Set-Content $tmp -Value "[font=휴먼명조체,size=18]제목`n[font=바탕체,size=12]긴 본문..." -Encoding UTF8

/run hwp action=new contentsPath="$tmp"
```

```
AIWORKS> 새 HWP 문서에 '안녕하세요, 홍길동입니다.' 라고 굴림체 14pt로 써줘
```

---

### 신규 액션 추가

`actions/action.<이름>.ps1` 파일을 생성하고 `Invoke-Action-<이름>` 함수를 정의하면 자동으로 로드됩니다:

```powershell
function Invoke-Action-MyAction {
    param([object]$Params)
    # 구현
    return "결과"
}
```

---

## 안전 시스템

위험 키워드(`delete`, `drop`, `format`, `remove`, `rm`, `shutdown`, `삭제`) 감지 시 실행 전 Y/N 확인을 요구합니다.
위험 작업은 `aiworks.log`에 타임스탬프와 함께 기록됩니다.

```
  [!] 파일 또는 데이터를 영구 삭제합니다. 이 작업은 되돌릴 수 없습니다.
      계속 진행하시겠습니까? [Y/N]
```

---

## 설정 파일 섹션

| 섹션 | 내용 |
|------|------|
| `[AI]` | API 엔드포인트, 인증, 모델, 시스템 프롬프트 |
| `[Browser]` | Chrome/IE 경로, Playwright 설정 |
| `[Office]` | HWP 보안경로, PDF 도구, COM 초기화 지연 |
| `[Safety]` | 위험 키워드, 로그 설정 |
| `[UI]` | 프롬프트, 스피너, 컬러, 히스토리 수 |
| `[Programs]` | 이름 = 실행파일 경로 매핑 |
| `[Warnings]` | 키워드 = 경고 메시지 매핑 |

---

## 요구 사항

- Windows PowerShell 5.1 이상
- 제어 대상 프로그램 설치 (Excel, HWP, Chrome 등)
- 내부 AI Agent API 접근 권한

## 제약 사항

- 외부 인터넷 연결 불필요 (내부망 전용 설계)
- PS 5.1 호환: `??` 연산자 미지원 → 내부 `Coalesce()` 헬퍼 사용
- COM 자동화: 대상 프로그램이 설치되어 있어야 함

---

## 라이선스

사내 내부 사용 전용
