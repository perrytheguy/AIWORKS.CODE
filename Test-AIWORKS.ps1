#requires -Version 5.1
# ============================================================
#  Test-AIWORKS.ps1 - AIWORKS.CODE Test Suite
#  실행: powershell -ExecutionPolicy Bypass -File Test-AIWORKS.ps1
# ============================================================
Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ─────────────────────────────────────────────────────────────
# [0] 메인 스크립트 로드 (REPL 실행 방지)
# ─────────────────────────────────────────────────────────────
$env:AIWORKS_TEST_MODE = "1"
$MainScript = Join-Path $PSScriptRoot "AIWORKS.code.ps1"

if (-not (Test-Path $MainScript)) {
    Write-Host "  [x] AIWORKS.code.ps1 not found: $MainScript" -ForegroundColor Red
    exit 1
}

. $MainScript

# Write-AgentLog 출력 캡처 (노이즈 제거)
$script:AgentLogs = [System.Collections.Generic.List[string]]::new()
function Write-AgentLog {
    param([string]$Message, [string]$Type = "Info")
    $script:AgentLogs.Add("[$Type] $Message")
}

# ─────────────────────────────────────────────────────────────
# [1] 테스트 프레임워크
# ─────────────────────────────────────────────────────────────
$script:PassCount = 0
$script:FailCount = 0

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "  ---- $Title " -NoNewline -ForegroundColor DarkCyan
    $dashes = "-" * [Math]::Max(1, 44 - $Title.Length)
    Write-Host $dashes -ForegroundColor DarkGray
}

function Assert {
    param(
        [string]$Name,
        [bool]$Condition,
        [string]$FailDetail = ""
    )
    if ($Condition) {
        Write-Host "  [PASS] $Name" -ForegroundColor Green
        $script:PassCount++
    } else {
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
        if ($FailDetail) {
            Write-Host "         >> $FailDetail" -ForegroundColor DarkRed
        }
        if ($script:AgentLogs.Count -gt 0) {
            Write-Host "         Log: $($script:AgentLogs[-1])" -ForegroundColor DarkYellow
        }
        $script:FailCount++
    }
    $script:AgentLogs.Clear()
}

function Assert-Equal {
    param([string]$Name, $Expected, $Actual)
    Assert -Name $Name -Condition ("$Actual" -eq "$Expected") `
           -FailDetail "Expected='$Expected'  Got='$Actual'"
}

function Assert-FileContains {
    param([string]$Name, [string]$File, [string]$Text)
    $content = Get-Content $File -Raw -Encoding UTF8
    Assert -Name $Name -Condition ($content -match [regex]::Escape($Text)) `
           -FailDetail "File does not contain: '$Text'"
}

function Assert-FileNotContains {
    param([string]$Name, [string]$File, [string]$Text)
    $content = Get-Content $File -Raw -Encoding UTF8
    Assert -Name $Name -Condition ($content -notmatch [regex]::Escape($Text)) `
           -FailDetail "File should NOT contain: '$Text'"
}

# ─────────────────────────────────────────────────────────────
# [2] 테스트 환경 준비 (임시 config 파일)
# ─────────────────────────────────────────────────────────────
$OriginalConfigPath = $script:ConfigPath
$TempConfig = Join-Path $env:TEMP ("AIWORKS_test_" + [datetime]::Now.Ticks + ".config")
Copy-Item -Path $script:ConfigPath -Destination $TempConfig -Force
$script:ConfigPath = $TempConfig

# 임시 config 에서 로드
$script:Config.Clear()
$script:Config["UI.ColorOutput"] = "false"
$script:Config["ColorOutput"]    = "false"

$loadSection = ""
foreach ($line in Get-Content $TempConfig -Encoding UTF8) {
    $trimmed = $line.Trim()
    if ($trimmed -match "^\[(.+)\]$") {
        $loadSection = $matches[1]
    } elseif ($trimmed -match "^([^#=]+)=(.*)$") {
        $k = $matches[1].Trim()
        $v = $matches[2].Trim()
        $script:Config["$loadSection.$k"] = $v
        $script:Config[$k]                = $v
    }
}

# ─────────────────────────────────────────────────────────────
# 배너
# ─────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "  AIWORKS.CODE - Test Suite" -ForegroundColor Cyan
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "  Temp config: $TempConfig" -ForegroundColor DarkGray

# ─────────────────────────────────────────────────────────────
# TEST-1: Import-Config 파싱
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-1: Import-Config 파싱"

Assert-Equal "AI.Provider = custom"          "custom"   $script:Config["AI.Provider"]
Assert-Equal "AI.TimeoutSec = 60"            "60"       $script:Config["AI.TimeoutSec"]
Assert-Equal "UI.Prompt = AIWORKS"           "AIWORKS"  $script:Config["UI.Prompt"]
Assert       "Section.Key 형식 로드"         ($script:Config.ContainsKey("AI.Model"))
Assert       "단순키 형식 로드"              ($script:Config.ContainsKey("Model"))
Assert-Equal "Section.Key == 단순키 동일값"  $script:Config["AI.Model"] $script:Config["Model"]
Assert       "Safety.ConfirmKeywords 로드"   ($script:Config.ContainsKey("Safety.ConfirmKeywords"))

# ─────────────────────────────────────────────────────────────
# TEST-2: /config get
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-2: /config get"

Assert "AI.Model 키 존재"                ($script:Config.ContainsKey("AI.Model"))
Assert "Browser.EnableIE 키 존재"        ($script:Config.ContainsKey("Browser.EnableIE"))
Assert "없는 키는 false"                 (-not $script:Config.ContainsKey("NoSuch.Key"))

try {
    Invoke-ConfigCommand -Cmd "/config get AI Model"
    Assert "/config get 정상 실행" $true
} catch {
    Assert "/config get 정상 실행" $false -FailDetail $_.Exception.Message
}

try {
    Invoke-ConfigCommand -Cmd "/config get"
    Assert "/config get 인수 부족 - 크래시 없음" $true
} catch {
    Assert "/config get 인수 부족 - 크래시 없음" $false -FailDetail $_.Exception.Message
}

# ─────────────────────────────────────────────────────────────
# TEST-3: /config set - 기존 키 수정
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-3: /config set (기존 키 수정)"

Set-ConfigValue -Section "AI" -Key "TimeoutSec" -Value "90"

Assert-Equal "메모리 업데이트 AI.TimeoutSec=90" "90" $script:Config["AI.TimeoutSec"]
Assert-Equal "단순키 메모리도 동기화"            "90" $script:Config["TimeoutSec"]
Assert-FileContains "파일에 반영됨" $TempConfig "TimeoutSec = 90"

Invoke-ConfigCommand -Cmd "/config set UI Prompt TEST_PROMPT"
Assert-Equal "Invoke-ConfigCommand set 메모리 반영" "TEST_PROMPT" $script:Config["UI.Prompt"]
Assert-FileContains "Invoke-ConfigCommand set 파일 반영" $TempConfig "Prompt = TEST_PROMPT"

# 원복
Set-ConfigValue -Section "AI"  -Key "TimeoutSec" -Value "60"
Set-ConfigValue -Section "UI"  -Key "Prompt"     -Value "AIWORKS"

# ─────────────────────────────────────────────────────────────
# TEST-4: /config set - 새 키 추가 (기존 섹션)
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-4: /config set (새 키 추가 - 기존 섹션)"

Set-ConfigValue -Section "AI" -Key "NewTestKey" -Value "hello"

Assert-Equal "새 키 메모리 추가"  "hello" $script:Config["AI.NewTestKey"]
Assert-FileContains "새 키 파일 추가" $TempConfig "NewTestKey = hello"

Remove-ConfigValue -Section "AI" -Key "NewTestKey"
Assert       "정리 후 메모리에서 제거"   (-not $script:Config.ContainsKey("AI.NewTestKey"))
Assert-FileNotContains "정리 후 파일에서 제거" $TempConfig "NewTestKey = hello"

# ─────────────────────────────────────────────────────────────
# TEST-5: /config set - 새 섹션 생성
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-5: /config set (새 섹션 생성)"

Set-ConfigValue -Section "TestSection" -Key "TestKey" -Value "TestValue"

Assert-Equal "새 섹션+키 메모리 추가" "TestValue" $script:Config["TestSection.TestKey"]
Assert-FileContains "새 섹션 헤더 파일 반영"  $TempConfig "[TestSection]"
Assert-FileContains "새 섹션 키값 파일 반영"  $TempConfig "TestKey = TestValue"

Remove-ConfigValue -Section "TestSection" -Key "TestKey"

# ─────────────────────────────────────────────────────────────
# TEST-6: /config remove
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-6: /config remove"

Set-ConfigValue -Section "Safety" -Key "TempRemoveKey" -Value "will_be_deleted"
Assert       "remove 대상 키 사전 존재"       ($script:Config.ContainsKey("Safety.TempRemoveKey"))
Assert-FileContains "remove 대상 파일 존재"  $TempConfig "TempRemoveKey = will_be_deleted"

Remove-ConfigValue -Section "Safety" -Key "TempRemoveKey"
Assert       "remove 후 메모리에서 삭제"      (-not $script:Config.ContainsKey("Safety.TempRemoveKey"))
Assert-FileNotContains "remove 후 파일에서 삭제" $TempConfig "TempRemoveKey = will_be_deleted"

try {
    Remove-ConfigValue -Section "AI" -Key "KeyThatDoesNotExist"
    Assert "없는 키 remove - 크래시 없음" $true
} catch {
    Assert "없는 키 remove - 크래시 없음" $false -FailDetail $_.Exception.Message
}

# ─────────────────────────────────────────────────────────────
# TEST-7: /config add-program
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-7: /config add-program"

$TestExe = "C:\Windows\System32\notepad.exe"
if (Test-Path $TestExe) {
    Add-ConfigProgram -Name "TestNotepad" -ExePath $TestExe
    Assert-Equal "add-program 메모리 추가"  $TestExe $script:Config["Programs.TestNotepad"]
    Assert-FileContains "add-program 파일 반영" $TempConfig "TestNotepad = $TestExe"
    Remove-ConfigValue -Section "Programs" -Key "TestNotepad"
} else {
    Write-Host "  [SKIP] notepad.exe not found" -ForegroundColor DarkYellow
}

try {
    Invoke-ConfigCommand -Cmd "/config add-program"
    Assert "/config add-program 인수 부족 - 크래시 없음" $true
} catch {
    Assert "/config add-program 인수 부족 - 크래시 없음" $false -FailDetail $_.Exception.Message
}

# ─────────────────────────────────────────────────────────────
# TEST-8: /config add-warning
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-8: /config add-warning"

Add-ConfigWarning -Key "testword" -Message "Test warning message"

Assert-Equal "add-warning 메모리 추가" "Test warning message" $script:Config["Warnings.testword"]
Assert-FileContains "add-warning 파일 반영" $TempConfig "testword = Test warning message"

$keywords = $script:Config["Safety.ConfirmKeywords"]
$kwAdded  = ($keywords -like "*testword*")
Assert "add-warning: ConfirmKeywords 자동 추가" $kwAdded

# 정리
Remove-ConfigValue -Section "Warnings" -Key "testword"
$oldKw = $script:Config["Safety.ConfirmKeywords"]
$newKw = (($oldKw -split ",") | Where-Object { $_.Trim() -ne "testword" }) -join ","
Set-ConfigValue -Section "Safety" -Key "ConfirmKeywords" -Value $newKw

# ─────────────────────────────────────────────────────────────
# TEST-9: /config reload
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-9: /config reload"

# 파일 직접 수정 (TimeoutSec -> 999)
$rawLines = Get-Content $TempConfig -Encoding UTF8
$modified = $rawLines | ForEach-Object {
    if ($_ -match "^TimeoutSec\s*=") { "TimeoutSec = 999" } else { $_ }
}
Set-Content -Path $TempConfig -Value $modified -Encoding UTF8

# reload 전 메모리는 아직 이전 값
$beforeReload = $script:Config["AI.TimeoutSec"]

# reload
$script:Config.Clear()
Import-Config -Path $TempConfig

Assert-Equal "reload 후 메모리 갱신" "999" $script:Config["AI.TimeoutSec"]

# 원복
Set-ConfigValue -Section "AI" -Key "TimeoutSec" -Value "60"

# ─────────────────────────────────────────────────────────────
# TEST-10: Invoke-ConfigCommand 토큰 파싱
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-10: Invoke-ConfigCommand 토큰 파싱"

# 따옴표로 묶인 값 (공백 포함)
Invoke-ConfigCommand -Cmd '/config set UI Prompt "MY PROMPT"'
Assert-Equal "quoted value (with space)" "MY PROMPT" $script:Config["UI.Prompt"]
Set-ConfigValue -Section "UI" -Key "Prompt" -Value "AIWORKS"

# 다중 단어 값 (따옴표 없음)
Invoke-ConfigCommand -Cmd "/config set AI Model test model v2"
Assert-Equal "multi-word value (no quotes)" "test model v2" $script:Config["AI.Model"]
Set-ConfigValue -Section "AI" -Key "Model" -Value "internal-llm"

# help / list / list [Section] / unknown -> 크래시 없어야 함
foreach ($cmd in @("/config help", "/config list", "/config list Safety", "/config unknown_sub")) {
    try {
        Invoke-ConfigCommand -Cmd $cmd
        Assert "$cmd - no crash" $true
    } catch {
        Assert "$cmd - no crash" $false -FailDetail $_.Exception.Message
    }
}

# ─────────────────────────────────────────────────────────────
# TEST-11: Coalesce 헬퍼
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-11: Coalesce helper"

Assert-Equal "값 있을 때 첫 번째 반환"     "hello" (Coalesce "hello" "world")
Assert-Equal "null 이면 두 번째 반환"       "world" (Coalesce $null  "world")
Assert-Equal "empty string 이면 두 번째 반환" "world" (Coalesce ""   "world")
Assert-Equal "0 은 유효한 값"              "0"     (Coalesce "0"    "world")
Assert-Equal "중첩 Coalesce"               "b"     (Coalesce (Coalesce "" "") "b")

# ─────────────────────────────────────────────────────────────
# TEST-12: Safety - 위험 키워드 감지
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-12: Safety - 위험 키워드 감지"

# 이전 테스트에서 Config 상태가 변경될 수 있으므로 명시적으로 설정
$script:Config["Safety.ConfirmKeywords"] = "delete,drop,format,remove,rm,shutdown"
$script:Config["ConfirmKeywords"]        = "delete,drop,format,remove,rm,shutdown"

$safe = [PSCustomObject]@{ requires_confirmation = $false }

Assert "delete 감지"   (Test-DangerousAction -UserInput "delete file" -Parsed $safe)
Assert "rm 감지"       (Test-DangerousAction -UserInput "rm -rf temp"  -Parsed $safe)
Assert "shutdown 감지" (Test-DangerousAction -UserInput "shutdown /s"  -Parsed $safe)
Assert "drop 감지"     (Test-DangerousAction -UserInput "drop table x" -Parsed $safe)
Assert "일반 입력 안전" (-not (Test-DangerousAction -UserInput "open file" -Parsed $safe))
Assert "requires_confirmation=true 이면 위험" `
    (Test-DangerousAction -UserInput "normal" -Parsed ([PSCustomObject]@{ requires_confirmation = $true }))

# ─────────────────────────────────────────────────────────────
# TEST-13: Parse-AIResponse JSON 파싱
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-13: Parse-AIResponse JSON 파싱"

$clean = '{"action":"answer","params":{},"message":"hello","requires_confirmation":false}'
$r = Parse-AIResponse -Raw $clean
Assert-Equal "clean JSON: action"  "answer" $r.action
Assert-Equal "clean JSON: message" "hello"  $r.message

# markdown 코드 펜스 제거
$fenced = '```json' + "`n" + $clean + "`n" + '```'
$r2 = Parse-AIResponse -Raw $fenced
Assert-Equal "fenced JSON: action"  "answer" $r2.action
Assert-Equal "fenced JSON: message" "hello"  $r2.message

# 잘못된 JSON -> answer fallback
$broken = "Not JSON at all"
$r3 = Parse-AIResponse -Raw $broken
Assert-Equal "broken JSON -> answer fallback" "answer" $r3.action
Assert-Equal "broken JSON -> raw as message"  $broken   $r3.message

# 빈 응답
$r4 = Parse-AIResponse -Raw ""
Assert-Equal "empty response -> answer fallback" "answer" $r4.action

# ─────────────────────────────────────────────────────────────
# TEST-14: VectorDB 기본 동작
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-14: VectorDB 기본 동작"

$TempDb        = Join-Path $env:TEMP ("AIWORKS_db_" + [datetime]::Now.Ticks + ".db")
"[]"           | Set-Content -Path $TempDb -Encoding UTF8
$OrigDbPath    = $script:DbPath
$script:DbPath = $TempDb

$noMatch = Search-VectorDB -Query "portal open"
Assert "빈 DB 검색 -> null" ($null -eq $noMatch)

$actionObj = [PSCustomObject]@{
    action = "chrome"
    params = @{ action = "open"; url = "https://portal.test" }
}
Add-VectorDBEntry -Query "open company portal" -ActionObj $actionObj -Response "Chrome opened"

$found = Search-VectorDB -Query "portal"
Assert "키워드 검색 성공"           ($null -ne $found)
Assert-Equal "검색 결과 action"     "chrome" $found.action.action

$notFound = Search-VectorDB -Query "excel spreadsheet"
Assert "미매칭 검색 -> null"        ($null -eq $notFound)

$db = Load-VectorDB
Assert-Equal "DB 항목 수 = 1"       1 $db.Count

$script:DbPath = $OrigDbPath
Remove-Item $TempDb -ErrorAction SilentlyContinue

# ─────────────────────────────────────────────────────────────
# TEST-15: Import-ActionModules 모듈 로드
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-15: 액션 모듈 로드"

Import-ActionModules

$actionsDir = Join-Path $PSScriptRoot "actions"
Assert "actions/ 디렉토리 존재" (Test-Path $actionsDir)

$mods = @("action.answer","action.office","action.hwp",
          "action.ie","action.chrome","action.pdf","action.shell")
foreach ($mod in $mods) {
    Assert "$mod.ps1 존재" (Test-Path (Join-Path $actionsDir "$mod.ps1"))
}

$funcs = @("Invoke-Action-answer","Invoke-Action-office","Invoke-Action-hwp",
           "Invoke-Action-ie","Invoke-Action-chrome","Invoke-Action-pdf","Invoke-Action-shell")
foreach ($fn in $funcs) {
    Assert "$fn 로드됨" ($null -ne (Get-Command $fn -ErrorAction SilentlyContinue))
}

# ─────────────────────────────────────────────────────────────
# TEST-16: /run command
# ─────────────────────────────────────────────────────────────
Write-Section "TEST-16: /run command"

# Mock action functions to capture dispatch (no real COM/process side effects)
$script:RunCapture = $null
function global:Invoke-Action-answer {
    param([object]$Params)
    $script:RunCapture = @{ action="answer"; params=$Params }
    return "MOCK:answer"
}
function global:Invoke-Action-shell {
    param([object]$Params)
    $script:RunCapture = @{ action="shell"; params=$Params }
    return "MOCK:shell"
}
function global:Invoke-Action-chrome {
    param([object]$Params)
    $script:RunCapture = @{ action="chrome"; params=$Params }
    return "MOCK:chrome"
}
function global:Invoke-Action-office {
    param([object]$Params)
    $script:RunCapture = @{ action="office"; params=$Params }
    return "MOCK:office"
}
function global:Invoke-Action-ie {
    param([object]$Params)
    $script:RunCapture = @{ action="ie"; params=$Params }
    return "MOCK:ie"
}
function global:Invoke-Action-hwp {
    param([object]$Params)
    $script:RunCapture = @{ action="hwp"; params=$Params }
    return "MOCK:hwp"
}
function global:Invoke-Action-pdf {
    param([object]$Params)
    $script:RunCapture = @{ action="pdf"; params=$Params }
    return "MOCK:pdf"
}

# /run (no args) / /run help — should not crash
try { Invoke-RunCommand -Cmd "/run";      Assert "/run (no args) - no crash" $true } catch { Assert "/run (no args) - no crash" $false -FailDetail $_.Exception.Message }
try { Invoke-RunCommand -Cmd "/run help"; Assert "/run help - no crash"      $true } catch { Assert "/run help - no crash"      $false -FailDetail $_.Exception.Message }

# /run unknown_action — warning, no crash
try {
    Invoke-RunCommand -Cmd "/run nonexistent_action"
    Assert "/run unknown action - no crash" $true
} catch {
    Assert "/run unknown action - no crash" $false -FailDetail $_.Exception.Message
}

# /run answer — fully specified params, no prompts
$script:RunCapture = $null
Invoke-RunCommand -Cmd '/run answer message="Hello world"'
Assert       "/run answer - dispatched"       ($null -ne $script:RunCapture)
Assert-Equal "/run answer - action name"      "answer" ($script:RunCapture.action)
Assert-Equal "/run answer - message param"    "Hello world" ($script:RunCapture.params.message)

# /run answer — quoted value with spaces
$script:RunCapture = $null
Invoke-RunCommand -Cmd '/run answer message="multi word value"'
Assert-Equal "/run answer - quoted multi-word" "multi word value" ($script:RunCapture.params.message)

# /run shell — fully specified
$script:RunCapture = $null
Invoke-RunCommand -Cmd '/run shell command="Get-Date" timeout=30'
Assert       "/run shell - dispatched"        ($null -ne $script:RunCapture)
Assert-Equal "/run shell - action name"       "shell" ($script:RunCapture.action)
Assert-Equal "/run shell - command param"     "Get-Date" ($script:RunCapture.params.command)
Assert-Equal "/run shell - timeout param"     "30" ($script:RunCapture.params.timeout)

# /run chrome — action + url specified
$script:RunCapture = $null
Invoke-RunCommand -Cmd '/run chrome action=open url="https://naver.com"'
Assert       "/run chrome - dispatched"       ($null -ne $script:RunCapture)
Assert-Equal "/run chrome - action"           "chrome" ($script:RunCapture.action)
Assert-Equal "/run chrome - sub-action"       "open" ($script:RunCapture.params.action)
Assert-Equal "/run chrome - url"              "https://naver.com" ($script:RunCapture.params.url)

# /run office — app + action + path specified
$script:RunCapture = $null
Invoke-RunCommand -Cmd '/run office app=excel action=open path="C:\data.xlsx"'
Assert       "/run office - dispatched"       ($null -ne $script:RunCapture)
Assert-Equal "/run office - app"              "excel" ($script:RunCapture.params.app)
Assert-Equal "/run office - sub-action"       "open" ($script:RunCapture.params.action)
Assert-Equal "/run office - path"             "C:\data.xlsx" ($script:RunCapture.params.path)

# /run pdf — path + action specified
$script:RunCapture = $null
Invoke-RunCommand -Cmd '/run pdf path="C:\report.pdf" action=info'
Assert       "/run pdf - dispatched"          ($null -ne $script:RunCapture)
Assert-Equal "/run pdf - path"                "C:\report.pdf" ($script:RunCapture.params.path)
Assert-Equal "/run pdf - action"              "info" ($script:RunCapture.params.action)

# /run hwp — action + path
$script:RunCapture = $null
Invoke-RunCommand -Cmd '/run hwp action=read path="C:\doc.hwp"'
Assert       "/run hwp - dispatched"          ($null -ne $script:RunCapture)
Assert-Equal "/run hwp - sub-action"          "read" ($script:RunCapture.params.action)
Assert-Equal "/run hwp - path"                "C:\doc.hwp" ($script:RunCapture.params.path)

# /run ie — action + url
$script:RunCapture = $null
Invoke-RunCommand -Cmd '/run ie action=open url="http://intranet"'
Assert       "/run ie - dispatched"           ($null -ne $script:RunCapture)
Assert-Equal "/run ie - sub-action"           "open" ($script:RunCapture.params.action)
Assert-Equal "/run ie - url"                  "http://intranet" ($script:RunCapture.params.url)

# Get-RunParamMeta — @() forces array (PS pipeline may unroll single-element returns)
$meta = @(Get-RunParamMeta -Action "answer")
Assert       "meta answer - not empty"        ($meta.Count -gt 0)
Assert-Equal "meta answer - first key"        "message" $meta[0].k
Assert       "meta answer - message required" ($meta[0].req -eq $true)

$meta2 = @(Get-RunParamMeta -Action "shell")
Assert-Equal "meta shell - first key"         "command" $meta2[0].k
Assert       "meta shell - command required"  ($meta2[0].req -eq $true)
Assert       "meta shell - workdir optional"  ($meta2[1].req -eq $false)

$meta3 = @(Get-RunParamMeta -Action "chrome")
Assert-Equal "meta chrome - first key"        "action" $meta3[0].k
Assert       "meta chrome - action has choices" ($meta3[0].choices.Count -gt 0)
Assert       "meta chrome - 'open' in choices"  ($meta3[0].choices -contains "open")

# Show-RunHelp — no crash
try { Show-RunHelp; Assert "Show-RunHelp - no crash" $true } catch { Assert "Show-RunHelp - no crash" $false -FailDetail $_.Exception.Message }

# ─────────────────────────────────────────────────────────────
# CLEANUP
# ─────────────────────────────────────────────────────────────
$script:ConfigPath       = $OriginalConfigPath
$env:AIWORKS_TEST_MODE   = ""
Remove-Item $TempConfig -ErrorAction SilentlyContinue

# ─────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────
$total = $script:PassCount + $script:FailCount
Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "  Test Results" -ForegroundColor Cyan
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host ("  Total : " + $total)                -ForegroundColor White
Write-Host ("  PASS  : " + $script:PassCount)     -ForegroundColor Green
if ($script:FailCount -gt 0) {
    Write-Host ("  FAIL  : " + $script:FailCount) -ForegroundColor Red
} else {
    Write-Host ("  FAIL  : " + $script:FailCount) -ForegroundColor Green
}
Write-Host "  ---------------------------------------------" -ForegroundColor DarkGray
if ($script:FailCount -eq 0) {
    Write-Host "  [+] All tests passed!" -ForegroundColor Green
} else {
    $failMsg = "  [x] " + $script:FailCount + " test(s) failed."
    Write-Host $failMsg -ForegroundColor Red
}
Write-Host ""

if ($script:FailCount -gt 0) { exit 1 } else { exit 0 }
