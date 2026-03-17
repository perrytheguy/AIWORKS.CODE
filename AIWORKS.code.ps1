# ============================================================
#  AIWORKS.CODE - Local AI Agent for Windows PowerShell
#  Run: powershell -ExecutionPolicy Bypass -File AIWORKS.code.ps1
#  Requires: PowerShell 5.1
# ============================================================
#requires -Version 5.1

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

# ─────────────────────────────────────────────────────────────
# [0] Global variables
# ─────────────────────────────────────────────────────────────
$script:Config        = @{}
$script:ChatHistory   = [System.Collections.Generic.List[hashtable]]::new()
$script:SessionActive = $true
$script:AuthToken     = ""
$script:ConfigPath    = Join-Path $PSScriptRoot "AIWORKS.code.config"
$script:DbPath        = Join-Path $PSScriptRoot "AIWORKS.code.db"

# PS 5.1 null-coalescing helper (?? operator replacement)
function Coalesce {
    param($Value, $Default)
    if ($null -ne $Value -and $Value -ne "") { return $Value } else { return $Default }
}

# ─────────────────────────────────────────────────────────────
# [1] Utility functions
# ─────────────────────────────────────────────────────────────

function Write-AgentLog {
    param(
        [string]$Message,
        [ValidateSet("Info","Success","Warning","Error","Thinking","Action","System")]
        [string]$Type = "Info"
    )
    $color = switch ($Type) {
        "Info"     { "Cyan"     }
        "Success"  { "Green"    }
        "Warning"  { "Yellow"   }
        "Error"    { "Red"      }
        "Thinking" { "Magenta"  }
        "Action"   { "Blue"     }
        "System"   { "DarkGray" }
    }
    $prefix = switch ($Type) {
        "Info"     { "  [*]" }
        "Success"  { "  [+]" }
        "Warning"  { "  [!]" }
        "Error"    { "  [x]" }
        "Thinking" { "  [~]" }
        "Action"   { "  [>]" }
        "System"   { "  [-]" }
    }
    if ($script:Config["UI.ColorOutput"] -eq "true" -or $script:Config["ColorOutput"] -eq "true") {
        Write-Host "$prefix $Message" -ForegroundColor $color
    } else {
        Write-Host "$prefix $Message"
    }
}

function Show-Thinking {
    param([string]$Label = "Thinking")
    $showThinking = Coalesce $script:Config["UI.ShowThinking"] (Coalesce $script:Config["ShowThinking"] "true")
    if ($showThinking -ne "true") { return }
    $frames = @("|", "/", "-", "\")
    for ($i = 0; $i -lt 16; $i++) {
        $f = $frames[$i % $frames.Count]
        Write-Host "`r  $f $Label..." -NoNewline -ForegroundColor Magenta
        Start-Sleep -Milliseconds 80
    }
    Write-Host "`r                              `r" -NoNewline
}

function Request-Confirmation {
    param([string]$Message)
    Write-Host ""
    Write-Host "  [!] $Message" -ForegroundColor Yellow
    Write-Host "      Continue? [Y/N] " -NoNewline -ForegroundColor Yellow
    $answer = Read-Host
    return ($answer -match "^[Yy]$")
}

function Write-AppLog {
    param([string]$Message, [string]$Level = "INFO")
    $logPath = Coalesce $script:Config["Safety.LogPath"] (Coalesce $script:Config["LogPath"] ".\aiworks.log")
    $enabled = Coalesce $script:Config["Safety.LogDangerousActions"] (Coalesce $script:Config["LogDangerousActions"] "true")
    if ($logPath -and $enabled -eq "true") {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logPath -Value "[$timestamp][$Level] $Message" -Encoding UTF8
    }
}

# ─────────────────────────────────────────────────────────────
# [2] Config file parsing (INI format)
# ─────────────────────────────────────────────────────────────

function Import-Config {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        Write-Host "  [x] Config file not found: $Path" -ForegroundColor Red
        exit 1
    }
    $section = ""
    foreach ($line in Get-Content $Path -Encoding UTF8) {
        $line = $line.Trim()
        if ($line -match "^\[(.+)\]$") {
            $section = $matches[1]
        } elseif ($line -match "^([^#=]+)=(.*)$") {
            $key   = $matches[1].Trim()
            $value = $matches[2].Trim()
            $script:Config["$section.$key"] = $value
            $script:Config[$key]            = $value
        }
    }
    Write-AgentLog "Config loaded." -Type System
}

# ─────────────────────────────────────────────────────────────
# [3] Authentication (browser-based IE COM + fallback)
# ─────────────────────────────────────────────────────────────

function Get-AuthToken {
    # 1. Static token from config
    $staticToken = Coalesce $script:Config["AI.AuthToken"] (Coalesce $script:Config["AuthToken"] "")
    if ($staticToken -ne "") {
        $script:AuthToken = $staticToken
        Write-AgentLog "Auth token loaded from config." -Type System
        return
    }

    # 2. Browser-based auth via IE COM
    $loginUrl    = Coalesce $script:Config["AI.LoginUrl"] (Coalesce $script:Config["LoginUrl"] "")
    $cookieField = Coalesce $script:Config["AI.CookieField"] (Coalesce $script:Config["CookieField"] "bearer_token")
    $enableIE    = Coalesce $script:Config["Browser.EnableIE"] (Coalesce $script:Config["EnableIE"] "true")

    if ($loginUrl -ne "" -and $enableIE -eq "true") {
        Write-AgentLog "Opening IE for authentication: $loginUrl" -Type System
        try {
            $ie = New-Object -ComObject "InternetExplorer.Application"
            $ie.Visible = $true
            $ie.Navigate($loginUrl)

            Write-Host ""
            Write-Host "  [*] Browser opened. Please log in, then press Enter here." -ForegroundColor Yellow
            Read-Host | Out-Null

            # Wait for page to settle
            $waited = 0
            while ($ie.Busy -and $waited -lt 30000) {
                Start-Sleep -Milliseconds 300
                $waited += 300
            }

            # Read cookies from document
            $rawCookies = ""
            try { $rawCookies = $ie.Document.cookie } catch {}

            if ($rawCookies -ne "") {
                $pairs = $rawCookies -split ";"
                foreach ($pair in $pairs) {
                    $pair = $pair.Trim()
                    if ($pair -match "^$([regex]::Escape($cookieField))=(.+)$") {
                        $script:AuthToken = $matches[1].Trim()
                        Write-AgentLog "Auth token obtained from cookie '$cookieField'." -Type Success
                        break
                    }
                }
            }

            try { $ie.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ie) | Out-Null
        }
        catch {
            Write-AgentLog "IE auth flow error: $($_.Exception.Message)" -Type Warning
        }
    }

    # 3. Fallback: manual token input
    if ($script:AuthToken -eq "") {
        Write-Host ""
        Write-Host "  [!] Auth token not found. Please enter bearer token manually." -ForegroundColor Yellow
        Write-Host "      Bearer token: " -NoNewline -ForegroundColor Yellow
        $manual = Read-Host
        if ($manual -ne "") {
            $script:AuthToken = $manual.Trim()
            Write-AgentLog "Auth token set manually." -Type Success
        } else {
            Write-AgentLog "No auth token provided. AI calls may fail." -Type Warning
        }
    }
}

# ─────────────────────────────────────────────────────────────
# [4] AI communication
# ─────────────────────────────────────────────────────────────

function Send-AIRequest {
    param([string]$UserInput)

    $endpoint  = Coalesce $script:Config["AI.Endpoint"]  (Coalesce $script:Config["Endpoint"] "")
    $model     = Coalesce $script:Config["AI.Model"]     (Coalesce $script:Config["Model"] "internal-llm")
    $maxTokens = [int](Coalesce $script:Config["AI.MaxTokens"]  (Coalesce $script:Config["MaxTokens"] "4096"))
    $timeout   = [int](Coalesce $script:Config["AI.TimeoutSec"] (Coalesce $script:Config["TimeoutSec"] "60"))
    $sysPrompt = Coalesce $script:Config["AI.SystemPrompt"] (Coalesce $script:Config["SystemPrompt"] "")
    $maxHist   = [int](Coalesce $script:Config["AI.MaxHistory"] (Coalesce $script:Config["MaxHistory"] "20"))

    # Prepend action constraint instructions to the user input
    $actionConstraint = "IMPORTANT: Respond ONLY in valid JSON with fields: action, params, message, requires_confirmation. " +
                        "Valid actions: answer, office, hwp, ie, chrome, pdf, shell. " +
                        "User query: "
    $fullInput = $actionConstraint + $UserInput

    $script:ChatHistory.Add(@{ role = "user"; content = $fullInput })

    while ($script:ChatHistory.Count -gt $maxHist) {
        $script:ChatHistory.RemoveAt(0)
    }

    # Build messages list, ensuring alternating user/assistant roles
    $messages = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($h in $script:ChatHistory) { $messages.Add($h) }

    $cleaned = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($m in $messages) {
        if ($cleaned.Count -gt 0 -and $cleaned[$cleaned.Count - 1].role -eq $m.role) {
            $cleaned[$cleaned.Count - 1] = $m
        } else {
            $cleaned.Add($m)
        }
    }
    $messages = $cleaned

    if ($endpoint -eq "") {
        Write-AgentLog "AI endpoint not configured. Set AI.Endpoint in config." -Type Error
        return $null
    }

    try {
        $authHeader = if ($script:AuthToken -ne "") { "Bearer $script:AuthToken" } else { "" }

        $allMessages = New-Object System.Collections.Generic.List[object]
        $allMessages.Add([ordered]@{ role = "system"; content = $sysPrompt })
        foreach ($m in $messages) { $allMessages.Add($m) }

        $bodyObj = [ordered]@{
            model      = $model
            messages   = $allMessages.ToArray()
            max_tokens = $maxTokens
        }
        $body = $bodyObj | ConvertTo-Json -Depth 10 -Compress

        $headers = @{ "Content-Type" = "application/json" }
        if ($authHeader -ne "") { $headers["Authorization"] = $authHeader }

        $response = Invoke-RestMethod `
            -Uri        $endpoint `
            -Method     POST `
            -Headers    $headers `
            -Body       $body `
            -TimeoutSec $timeout

        # Handle both OpenAI-style and generic response formats
        $content = ""
        if ($response.choices -and $response.choices.Count -gt 0) {
            $content = $response.choices[0].message.content
        } elseif ($response.content) {
            $content = $response.content
        } elseif ($response.message) {
            $content = $response.message
        } else {
            $content = $response | ConvertTo-Json -Depth 5
        }

        $script:ChatHistory.Add(@{ role = "assistant"; content = $content })
        return $content
    }
    catch {
        Write-AgentLog "AI communication error: $($_.Exception.Message)" -Type Error
        return $null
    }
}

function Parse-AIResponse {
    param([string]$Raw)
    if (-not $Raw) {
        return [PSCustomObject]@{
            action               = "answer"
            message              = "(no response)"
            params               = @{}
            requires_confirmation = $false
        }
    }
    try {
        # Strip markdown code fences if present
        $json = $Raw.Trim()
        if ($json -match '(?s)```json\s*(.*?)\s*```') {
            $json = $matches[1].Trim()
        } elseif ($json -match '(?s)```\s*(.*?)\s*```') {
            $json = $matches[1].Trim()
        }
        return $json | ConvertFrom-Json
    }
    catch {
        return [PSCustomObject]@{
            action               = "answer"
            message              = $Raw
            params               = @{}
            requires_confirmation = $false
        }
    }
}

# ─────────────────────────────────────────────────────────────
# [5] Safety system
# ─────────────────────────────────────────────────────────────

function Test-DangerousAction {
    param([string]$UserInput, [object]$Parsed)
    if ($Parsed.requires_confirmation -eq $true) { return $true }
    $raw      = Coalesce $script:Config["Safety.ConfirmKeywords"] (Coalesce $script:Config["ConfirmKeywords"] "")
    $keywords = $raw -split ","
    foreach ($kw in $keywords) {
        $kw = $kw.Trim()
        if ($kw -ne "" -and $UserInput -match [regex]::Escape($kw)) { return $true }
    }
    return $false
}

function Get-DangerWarningMessage {
    param([string]$UserInput)
    $raw      = Coalesce $script:Config["Safety.ConfirmKeywords"] (Coalesce $script:Config["ConfirmKeywords"] "")
    $keywords = $raw -split ","
    foreach ($kw in $keywords) {
        $kw = $kw.Trim()
        if ($kw -ne "" -and $UserInput -match [regex]::Escape($kw)) {
            $warnMsg = Coalesce $script:Config["Warnings.$kw"] (Coalesce $script:Config[$kw] "")
            if ($warnMsg -ne "") { return $warnMsg }
        }
    }
    return "This operation may be dangerous. Proceed with caution."
}

# ─────────────────────────────────────────────────────────────
# [6] Action dispatcher (delegates to action modules)
# ─────────────────────────────────────────────────────────────

function Invoke-AgentAction {
    param([string]$UserInput, [object]$Parsed)

    $action = $Parsed.action
    $params = $Parsed.params
    $msg    = $Parsed.message

    if (Test-DangerousAction -UserInput $UserInput -Parsed $Parsed) {
        $warnMsg   = Get-DangerWarningMessage -UserInput $UserInput
        $confirmed = Request-Confirmation -Message $warnMsg
        if (-not $confirmed) {
            Write-AgentLog "Action cancelled by user." -Type Warning
            Write-AppLog "User rejected: $action / $UserInput" -Level "WARN"
            return $null
        }
        Write-AppLog "User approved: $action / $UserInput" -Level "INFO"
    }

    if ($msg) { Write-AgentLog $msg -Type Thinking }

    $result = $null
    switch ($action) {
        "answer" { Invoke-Action-answer -Params $Parsed;  return $null }
        "office" { $result = Invoke-Action-office -Params $params }
        "hwp"    { $result = Invoke-Action-hwp    -Params $params }
        "ie"     { $result = Invoke-Action-ie     -Params $params }
        "chrome" { $result = Invoke-Action-chrome -Params $params }
        "pdf"    { $result = Invoke-Action-pdf    -Params $params }
        "shell"  { $result = Invoke-Action-shell  -Params $params }
        default  {
            Write-AgentLog "Unknown action: $action" -Type Warning
            return $null
        }
    }

    if ($result) {
        Write-Host ""
        Write-Host "  [Result]" -ForegroundColor DarkGray
        Write-Host "  $result" -ForegroundColor White
        Write-Host ""
        $script:ChatHistory.Add(@{ role = "assistant"; content = "[ACTION RESULT] $result" })
    }

    return $result
}

# ─────────────────────────────────────────────────────────────
# [7] Vector DB management
# ─────────────────────────────────────────────────────────────

function Load-VectorDB {
    # Returns [System.Collections.Generic.List[object]].
    # Wrapped with unary comma (,$list) to prevent PS pipeline enumeration.
    $list = [System.Collections.Generic.List[object]]::new()
    if (Test-Path $script:DbPath) {
        try {
            $raw = [System.IO.File]::ReadAllText($script:DbPath, [System.Text.Encoding]::UTF8)
            if ($raw.Trim() -ne "" -and $raw.Trim() -ne "[]") {
                foreach ($item in @($raw | ConvertFrom-Json)) { $list.Add($item) }
            }
        }
        catch {
            Write-AgentLog "DB load error: $($_.Exception.Message)" -Type Warning
        }
    }
    return ,$list
}

function Save-VectorDB {
    param([System.Collections.Generic.List[object]]$Entries)
    try {
        # Force array output even for 0 or 1 entries (PS pipeline would otherwise unwrap single items)
        $json = if ($null -eq $Entries -or $Entries.Count -eq 0) { "[]" } else { ConvertTo-Json -InputObject ([object[]]$Entries) -Depth 10 }
        [System.IO.File]::WriteAllText($script:DbPath, $json, [System.Text.Encoding]::UTF8)
    }
    catch {
        Write-AgentLog "DB save error: $($_.Exception.Message)" -Type Error
    }
}

function Search-VectorDB {
    param([string]$Query)
    $db = Load-VectorDB
    if ($null -eq $db -or $db.Count -eq 0) { return $null }
    $queryLower = $Query.ToLower()
    foreach ($entry in $db) {
        foreach ($kw in $entry.keywords) {
            if ($queryLower.Contains($kw.ToLower())) {
                return $entry
            }
        }
    }
    return $null
}

function Add-VectorDBEntry {
    param([string]$Query, [object]$ActionObj, [string]$Response)
    # Extract keywords: split on spaces, filter short tokens
    $keywords = $Query.ToLower() -split "\s+" | Where-Object { $_.Length -ge 2 } | Select-Object -Unique
    $entry = [PSCustomObject]@{
        keywords = @($keywords)
        action   = $ActionObj
        response = $Response
        created  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }
    $db = Load-VectorDB
    if ($null -eq $db) { $db = [System.Collections.Generic.List[object]]::new() }
    $db.Add($entry)
    Save-VectorDB -Entries $db
    Write-AgentLog "Entry saved to DB." -Type Success
}

function Invoke-DbCommand {
    param([string]$Cmd)
    $tokens = $Cmd.Trim() -split "\s+", 3
    $sub    = if ($tokens.Count -gt 1) { $tokens[1].ToLower() } else { "list" }

    switch ($sub) {
        "list" {
            $db = Load-VectorDB
            if ($null -eq $db) { $db = [System.Collections.Generic.List[object]]::new() }
            Write-Host ""
            if ($db.Count -eq 0) {
                Write-AgentLog "Vector DB is empty." -Type System
            } else {
                Write-Host "  -- Vector DB Entries ($($db.Count)) --" -ForegroundColor DarkGray
                for ($i = 0; $i -lt $db.Count; $i++) {
                    $e = $db[$i]
                    $kwStr = $e.keywords -join ", "
                    Write-Host "  [$i] Keywords: $kwStr" -ForegroundColor Cyan
                    Write-Host "       Action : $($e.action.action)" -ForegroundColor White
                    Write-Host "       Created: $($e.created)" -ForegroundColor DarkGray
                }
                Write-Host ""
            }
        }
        "delete" {
            if ($tokens.Count -lt 3) {
                Write-AgentLog "Usage: /db delete <index>" -Type Warning
                return
            }
            $idx = [int]$tokens[2]
            $db  = Load-VectorDB
            if ($null -eq $db) { $db = [System.Collections.Generic.List[object]]::new() }
            if ($idx -ge 0 -and $idx -lt $db.Count) {
                $db.RemoveAt($idx)
                Save-VectorDB -Entries $db
                Write-AgentLog "Entry $idx deleted." -Type Success
            } else {
                Write-AgentLog "Invalid index: $idx" -Type Warning
            }
        }
        "clear" {
            $empty = [System.Collections.Generic.List[object]]::new()
            Save-VectorDB -Entries $empty
            Write-AgentLog "Vector DB cleared." -Type Success
        }
        "search" {
            if ($tokens.Count -lt 3) {
                Write-AgentLog "Usage: /db search <query>" -Type Warning
                return
            }
            $q      = $tokens[2]
            $result = Search-VectorDB -Query $q
            Write-Host ""
            if ($result) {
                Write-Host "  [DB Match Found]" -ForegroundColor Green
                Write-Host "  Keywords: $($result.keywords -join ', ')" -ForegroundColor Cyan
                Write-Host "  Action  : $($result.action.action)" -ForegroundColor White
                Write-Host "  Response: $($result.response)" -ForegroundColor DarkGray
            } else {
                Write-AgentLog "No DB match for: $q" -Type System
            }
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Host "  -- /db commands --" -ForegroundColor DarkGray
            Write-Host "  /db list              Show all DB entries" -ForegroundColor Yellow
            Write-Host "  /db delete <index>    Delete entry by index" -ForegroundColor Yellow
            Write-Host "  /db clear             Clear all entries" -ForegroundColor Yellow
            Write-Host "  /db search <query>    Test search without executing" -ForegroundColor Yellow
            Write-Host ""
        }
    }
}

# ─────────────────────────────────────────────────────────────
# [8] Config editor (/config commands)
# ─────────────────────────────────────────────────────────────

function Get-ConfigLineIndex {
    param([string[]]$Lines, [string]$Section, [string]$Key)
    $currentSection = ""
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i].Trim()
        if ($line -match "^\[(.+)\]$") { $currentSection = $matches[1] }
        elseif ($line -match "^([^#=]+)=(.*)$") {
            if ($currentSection -eq $Section -and $matches[1].Trim() -eq $Key) { return $i }
        }
    }
    return -1
}

function Get-SectionEndIndex {
    param([string[]]$Lines, [string]$Section)
    $inSection = $false
    $lastIdx   = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i].Trim()
        if ($line -match "^\[(.+)\]$") {
            if ($inSection) { return $lastIdx }
            if ($matches[1] -eq $Section) { $inSection = $true }
        }
        if ($inSection -and $line -ne "" -and -not $line.StartsWith("#")) { $lastIdx = $i }
    }
    return $lastIdx
}

function Test-ConfigSection {
    param([string[]]$Lines, [string]$Section)
    return ($Lines | Where-Object { $_ -match "^\[$([regex]::Escape($Section))\]" }).Count -gt 0
}

function Show-ConfigList {
    param([string]$FilterSection = "")
    if (-not (Test-Path $script:ConfigPath)) {
        Write-AgentLog "Config file not found: $script:ConfigPath" -Type Error
        return
    }
    $rawLines = [System.IO.File]::ReadAllLines($script:ConfigPath, [System.Text.Encoding]::UTF8)
    Write-Host ""
    $currentSection = ""
    $show = $true
    foreach ($line in $rawLines) {
        $l = $line.Trim()
        if ($l -match "^\[(.+)\]$") {
            $currentSection = $matches[1]
            $show = (-not $FilterSection) -or ($currentSection -eq $FilterSection)
            if ($show) { Write-Host "  [$currentSection]" -ForegroundColor Yellow }
        } elseif ($show) {
            if ($l -eq "" -or $l.StartsWith("#")) {
                Write-Host "  $line" -ForegroundColor DarkGray
            } elseif ($l -match "^([^=]+)=(.*)$") {
                $k = $matches[1].TrimEnd()
                $v = $matches[2].TrimStart()
                Write-Host "  " -NoNewline
                Write-Host $k -NoNewline -ForegroundColor Cyan
                Write-Host " = " -NoNewline -ForegroundColor DarkGray
                Write-Host $v -ForegroundColor White
            }
        }
    }
    Write-Host ""
}

function Get-ConfigValue {
    param([string]$Section, [string]$Key)
    $fullKey = "$Section.$Key"
    if ($script:Config.ContainsKey($fullKey)) {
        Write-Host ""
        Write-Host "  [$Section] $Key" -ForegroundColor Cyan
        Write-Host "  => $($script:Config[$fullKey])" -ForegroundColor White
        Write-Host ""
    } else {
        Write-AgentLog "Key not found: [$Section] $Key" -Type Warning
    }
}

function Set-ConfigValue {
    param([string]$Section, [string]$Key, [string]$Value)
    $lines   = [System.Collections.Generic.List[string]](Get-Content $script:ConfigPath -Encoding UTF8)
    $lineIdx = Get-ConfigLineIndex -Lines $lines -Section $Section -Key $Key
    if ($lineIdx -ge 0) {
        $lines[$lineIdx] = "$Key = $Value"
        Write-AgentLog "[$Section] $Key updated: $Value" -Type Success
    } else {
        if (-not (Test-ConfigSection -Lines $lines -Section $Section)) {
            $lines.Add("")
            $lines.Add("[$Section]")
            $lines.Add("$Key = $Value")
            Write-AgentLog "New section [$Section] and key added: $Key = $Value" -Type Success
        } else {
            $endIdx = Get-SectionEndIndex -Lines $lines -Section $Section
            if ($endIdx -ge 0) { $lines.Insert($endIdx + 1, "$Key = $Value") }
            else                { $lines.Add("$Key = $Value") }
            Write-AgentLog "[$Section] New key added: $Key = $Value" -Type Success
        }
    }
    Set-Content -Path $script:ConfigPath -Value $lines -Encoding UTF8
    $script:Config["$Section.$Key"] = $Value
    $script:Config[$Key]            = $Value
}

function Remove-ConfigValue {
    param([string]$Section, [string]$Key)
    $lines   = [System.Collections.Generic.List[string]](Get-Content $script:ConfigPath -Encoding UTF8)
    $lineIdx = Get-ConfigLineIndex -Lines $lines -Section $Section -Key $Key
    if ($lineIdx -ge 0) {
        $lines.RemoveAt($lineIdx)
        Set-Content -Path $script:ConfigPath -Value $lines -Encoding UTF8
        $script:Config.Remove("$Section.$Key") | Out-Null
        $script:Config.Remove($Key)            | Out-Null
        Write-AgentLog "[$Section] $Key removed." -Type Success
    } else {
        Write-AgentLog "Key not found: [$Section] $Key" -Type Warning
    }
}

function Add-ConfigProgram {
    param([string]$Name, [string]$ExePath)
    if (-not (Test-Path $ExePath)) {
        Write-AgentLog "Path not found: $ExePath" -Type Warning
        Write-Host "     Path does not exist. Save anyway? [Y/N] " -NoNewline -ForegroundColor Yellow
        $ans = Read-Host
        if ($ans -notmatch "^[Yy]$") { return }
    }
    Set-ConfigValue -Section "Programs" -Key $Name -Value $ExePath
}

function Add-ConfigWarning {
    param([string]$Key, [string]$Message)
    Set-ConfigValue -Section "Warnings" -Key $Key -Value $Message
    $raw      = Coalesce $script:Config["Safety.ConfirmKeywords"] (Coalesce $script:Config["ConfirmKeywords"] "")
    $existing = $raw -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    if ($existing -notcontains $Key) {
        $newKeywords = ($existing + $Key) -join ","
        Set-ConfigValue -Section "Safety" -Key "ConfirmKeywords" -Value $newKeywords
        Write-AgentLog "Safety.ConfirmKeywords: '$Key' auto-added." -Type Info
    }
}

function Show-ConfigHelp {
    Write-Host ""
    Write-Host "  -- /config commands ----------------------------------" -ForegroundColor DarkGray
    Write-Host "  /config list                    Show all settings" -ForegroundColor Yellow
    Write-Host "  /config list <Section>          Show section settings" -ForegroundColor Yellow
    Write-Host "  /config get <Section> <Key>     Get a value" -ForegroundColor Yellow
    Write-Host "  /config set <Section> <Key> <Value>  Set a value" -ForegroundColor Yellow
    Write-Host "  /config remove <Section> <Key>  Remove a key" -ForegroundColor Yellow
    Write-Host "  /config add-program <Name> <Path>    Add program" -ForegroundColor Yellow
    Write-Host "  /config add-warning <Key> <Msg>      Add warning" -ForegroundColor Yellow
    Write-Host "  /config reload                  Reload config from file" -ForegroundColor Yellow
    Write-Host "  /config help                    Show this help" -ForegroundColor Yellow
    Write-Host "  ------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
}

function Invoke-ConfigCommand {
    param([string]$Cmd)
    $tokens  = [System.Collections.Generic.List[string]]::new()
    $pattern = '"([^"]*)"|(\S+)'
    [regex]::Matches($Cmd.Trim(), $pattern) | ForEach-Object {
        if ($_.Groups[1].Success) { $tokens.Add($_.Groups[1].Value) }
        else                      { $tokens.Add($_.Groups[2].Value) }
    }
    $sub = if ($tokens.Count -gt 1) { $tokens[1].ToLower() } else { "help" }

    switch ($sub) {
        "list" {
            $sec = if ($tokens.Count -gt 2) { $tokens[2] } else { "" }
            Show-ConfigList -FilterSection $sec
        }
        "get" {
            if ($tokens.Count -lt 4) { Write-AgentLog "Usage: /config get <Section> <Key>" -Type Warning; return }
            Get-ConfigValue -Section $tokens[2] -Key $tokens[3]
        }
        "set" {
            if ($tokens.Count -lt 5) { Write-AgentLog "Usage: /config set <Section> <Key> <Value>" -Type Warning; return }
            $val = $tokens[4..($tokens.Count-1)] -join " "
            Set-ConfigValue -Section $tokens[2] -Key $tokens[3] -Value $val
        }
        "add-program" {
            if ($tokens.Count -lt 4) { Write-AgentLog "Usage: /config add-program <Name> <Path>" -Type Warning; return }
            Add-ConfigProgram -Name $tokens[2] -ExePath ($tokens[3..($tokens.Count-1)] -join " ")
        }
        "add-warning" {
            if ($tokens.Count -lt 4) { Write-AgentLog "Usage: /config add-warning <Key> <Message>" -Type Warning; return }
            Add-ConfigWarning -Key $tokens[2] -Message ($tokens[3..($tokens.Count-1)] -join " ")
        }
        "remove" {
            if ($tokens.Count -lt 4) { Write-AgentLog "Usage: /config remove <Section> <Key>" -Type Warning; return }
            Remove-ConfigValue -Section $tokens[2] -Key $tokens[3]
        }
        "reload" {
            $script:Config.Clear()
            Import-Config -Path $script:ConfigPath
            Write-AgentLog "Config reloaded." -Type Success
        }
        default { Show-ConfigHelp }
    }
}

# ─────────────────────────────────────────────────────────────
# [9] /run command - Direct action executor
# ─────────────────────────────────────────────────────────────

function Show-RunHelp {
    Write-Host ""
    Write-Host "  -- /run <action> [key=value ...]  ---------------" -ForegroundColor DarkGray
    @(
        @("answer", "Display a text message",            "message=TEXT"),
        @("shell",  "Execute a PowerShell command",      "command=CMD [workdir=PATH] [timeout=60]"),
        @("chrome", "Control Chrome browser",            "action=open|navigate|screenshot|script url=URL"),
        @("office", "Control Office (Excel/Word/PPT)",   "app=excel|word|ppt action=open|read|new|save|close|pdf [path=FILE]"),
        @("ie",     "Control Internet Explorer",         "action=open|read|input|click|wait|close [url=URL]"),
        @("hwp",    "Control HWP word processor",        "action=open|read|new|save|pdf|close [path=FILE]"),
        @("pdf",    "Read/inspect a PDF file",           "path=FILE [action=read|info] [maxchars=3000]")
    ) | ForEach-Object {
        Write-Host ("  {0,-8}  {1}" -f $_[0], $_[2]) -ForegroundColor Yellow
        Write-Host ("           {0}" -f $_[1])        -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  Omit params for interactive prompts." -ForegroundColor DarkGray
    Write-Host "  -----------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
}

function Get-RunParamMeta {
    param([string]$Action)
    switch ($Action) {
        "answer" { return @(
            @{ k="message"; label="Message text";                                   req=$true;  choices=@() }
        )}
        "shell"  { return @(
            @{ k="command"; label="PowerShell command";                             req=$true;  choices=@() }
            @{ k="workdir"; label="Working directory (Enter to skip)";              req=$false; choices=@() }
            @{ k="timeout"; label="Timeout in seconds (default 60, Enter to skip)"; req=$false; choices=@() }
        )}
        "chrome" { return @(
            @{ k="action"; label="Sub-action";                                      req=$true;  choices=@("open","navigate","screenshot","script") }
            @{ k="url";    label="Target URL";                                      req=$true;  choices=@() }
            @{ k="script"; label="JavaScript code (script action, Enter to skip)";  req=$false; choices=@() }
            @{ k="output"; label="Output path (screenshot, Enter to skip)";         req=$false; choices=@() }
            @{ k="profile";label="Chrome profile directory (Enter to skip)";        req=$false; choices=@() }
        )}
        "office" { return @(
            @{ k="app";    label="Application";                                     req=$true;  choices=@("excel","word","ppt") }
            @{ k="action"; label="Sub-action";                                      req=$true;  choices=@("open","read","new","save","close","pdf") }
            @{ k="path";   label="File path (Enter to skip for new/close)";         req=$false; choices=@() }
            @{ k="sheet";  label="Sheet name or index (Excel read, Enter to skip)"; req=$false; choices=@() }
        )}
        "ie"     { return @(
            @{ k="action";   label="Sub-action";                                    req=$true;  choices=@("open","read","input","click","wait","close") }
            @{ k="url";      label="Target URL (Enter to skip for close/wait)";     req=$false; choices=@() }
            @{ k="selector"; label="Element ID (input/click, Enter to skip)";       req=$false; choices=@() }
            @{ k="value";    label="Input value (input action, Enter to skip)";     req=$false; choices=@() }
            @{ k="timeout";  label="Timeout in seconds (default 30, Enter to skip)";req=$false; choices=@() }
        )}
        "hwp"    { return @(
            @{ k="action"; label="Sub-action";                                      req=$true;  choices=@("open","read","new","save","pdf","close") }
            @{ k="path";   label="File path (Enter to skip for new/close)";         req=$false; choices=@() }
        )}
        "pdf"    { return @(
            @{ k="path";     label="PDF file path";                                 req=$true;  choices=@() }
            @{ k="action";   label="Sub-action (Enter for read)";                   req=$false; choices=@("read","info") }
            @{ k="maxchars"; label="Max characters (default 3000, Enter to skip)";  req=$false; choices=@() }
        )}
        default  { return @() }
    }
}

function Invoke-RunPrompt {
    param([string]$Action, [hashtable]$Params)
    # @() forces array even when Get-RunParamMeta returns a single hashtable (PS pipeline unrolling)
    $meta = @(Get-RunParamMeta -Action $Action)
    Write-Host ""
    $dash = "-" * [Math]::Max(1, 38 - $Action.Length)
    Write-Host ("  -- /run {0} {1}" -f $Action, $dash) -ForegroundColor DarkGray

    foreach ($p in $meta) {
        $k = $p.k
        # Already provided on command line — just display
        if ($Params.ContainsKey($k) -and "$($Params[$k])" -ne "") {
            Write-Host ("  {0,-10} = {1}" -f $k, $Params[$k]) -ForegroundColor Cyan
            continue
        }
        # Choice-based param: show numbered sub-menu
        if ($p.choices.Count -gt 0) {
            Write-Host ""
            Write-Host ("  {0}:" -f $p.label) -ForegroundColor Yellow
            for ($i = 0; $i -lt $p.choices.Count; $i++) {
                Write-Host ("    [{0}] {1}" -f ($i + 1), $p.choices[$i]) -ForegroundColor White
            }
            $sel = Read-Host ("  Select [1-{0}] or type value" -f $p.choices.Count)
            if ($sel -match '^\d+$') {
                $idx = [int]$sel - 1
                if ($idx -ge 0 -and $idx -lt $p.choices.Count) {
                    $Params[$k] = $p.choices[$idx]
                } else {
                    Write-AgentLog "Invalid choice '$sel' for '$k'. Cancelled." -Type Warning
                    return $null
                }
            } elseif ($sel -ne "") {
                $Params[$k] = $sel
            } elseif ($p.req) {
                Write-AgentLog "Required param '$k' not provided. Cancelled." -Type Warning
                return $null
            }
        } else {
            # Free-text param
            $val = Read-Host ("  {0}" -f $p.label)
            if ($val -ne "") {
                $Params[$k] = $val
            } elseif ($p.req) {
                Write-AgentLog "Required param '$k' not provided. Cancelled." -Type Warning
                return $null
            }
        }
    }
    Write-Host ""
    return $Params
}

function Invoke-RunCommand {
    param([string]$Cmd)
    # Split on first two whitespace tokens to get /run + action, rest is key=value string
    $words = $Cmd.Trim() -split '\s+', 3

    $validActions = @("answer","shell","chrome","office","ie","hwp","pdf")
    $actionName   = if ($words.Count -gt 1) { $words[1].ToLower() } else { "" }

    if ($actionName -eq "" -or $actionName -eq "help") { Show-RunHelp; return }

    if ($validActions -notcontains $actionName) {
        Write-AgentLog "Unknown action: '$actionName'. Available: $($validActions -join ', ')" -Type Warning
        return
    }

    # Parse key=value pairs supporting both key=plain and key="quoted value"
    $params = @{}
    $rest   = if ($words.Count -gt 2) { $words[2] } else { "" }
    if ($rest -ne "") {
        [regex]::Matches($rest, '(\w+)=(?:"([^"]*)"|([\S]+))') | ForEach-Object {
            $k = $_.Groups[1].Value.ToLower()
            $v = if ($_.Groups[2].Success) { $_.Groups[2].Value } else { $_.Groups[3].Value }
            $params[$k] = $v
        }
    }

    # Interactive completion for missing params
    $params = Invoke-RunPrompt -Action $actionName -Params $params
    if ($null -eq $params) { return }

    # Build PSCustomObject for action modules
    $paramObj = New-Object PSObject -Property $params

    $fnName = "Invoke-Action-$actionName"
    $fn     = Get-Command $fnName -ErrorAction SilentlyContinue
    if ($null -eq $fn) {
        Write-AgentLog "Action '$actionName' not loaded. Check actions/ directory." -Type Error
        return
    }

    Write-AgentLog "Running action: $actionName" -Type Action
    $result = & $fn -Params $paramObj
    if ($result) {
        Write-Host ""
        Write-Host "  -- Result ----------------------------------" -ForegroundColor DarkGray
        Write-Host ($result -join "`n") -ForegroundColor White
        Write-Host "  -------------------------------------------" -ForegroundColor DarkGray
        Write-Host ""
    }
}

# ─────────────────────────────────────────────────────────────
# [10] Slash commands
# ─────────────────────────────────────────────────────────────

function Invoke-SlashCommand {
    param([string]$Cmd)

    switch -Regex ($Cmd.Trim()) {
        "^/exit$" {
            Write-AgentLog "Ending AIWORKS session." -Type System
            $script:SessionActive = $false
            return $true
        }
        "^/clear$" {
            Clear-Host
            Show-Banner
            return $true
        }
        "^/status$" {
            $tokenStatus = if ($script:AuthToken -ne "") { "Set" } else { "Not set" }
            Write-Host ""
            Write-Host "  -- Session Status ------------------------------" -ForegroundColor DarkGray
            Write-Host "  History  : $($script:ChatHistory.Count) entries" -ForegroundColor Cyan
            Write-Host "  Endpoint : $($script:Config["AI.Endpoint"])"      -ForegroundColor Cyan
            Write-Host "  Model    : $(Coalesce $script:Config["AI.Model"] $script:Config["Model"])" -ForegroundColor Cyan
            Write-Host "  Auth     : $tokenStatus"                           -ForegroundColor Cyan
            Write-Host "  DB Path  : $script:DbPath"                         -ForegroundColor Cyan
            $db = Load-VectorDB
            $dbCount = if ($null -eq $db) { 0 } else { $db.Count }
            Write-Host "  DB Items : $dbCount"                                -ForegroundColor Cyan
            Write-Host "  -----------------------------------------------" -ForegroundColor DarkGray
            Write-Host ""
            return $true
        }
        "^/history$" {
            Write-Host ""
            if ($script:ChatHistory.Count -eq 0) {
                Write-AgentLog "No chat history." -Type System
            } else {
                foreach ($h in $script:ChatHistory) {
                    $color   = if ($h.role -eq "user") { "Cyan" } else { "White" }
                    $preview = $h.content.Substring(0, [Math]::Min($h.content.Length, 100))
                    Write-Host "  [$($h.role.ToUpper())] $preview" -ForegroundColor $color
                }
            }
            Write-Host ""
            return $true
        }
        "^/reset$" {
            $script:ChatHistory.Clear()
            Write-AgentLog "Chat history cleared." -Type Success
            return $true
        }
        "^/config" {
            Invoke-ConfigCommand -Cmd $Cmd
            return $true
        }
        "^/db" {
            Invoke-DbCommand -Cmd $Cmd
            return $true
        }
        "^/run" {
            Invoke-RunCommand -Cmd $Cmd
            return $true
        }
        "^/help$" {
            Write-Host ""
            Write-Host "  -- Available Commands --------------------------" -ForegroundColor DarkGray
            Write-Host "  /exit              End session"                    -ForegroundColor Yellow
            Write-Host "  /clear             Clear screen + show banner"     -ForegroundColor Yellow
            Write-Host "  /status            Show session status"            -ForegroundColor Yellow
            Write-Host "  /history           Show chat history"              -ForegroundColor Yellow
            Write-Host "  /reset             Clear chat history"             -ForegroundColor Yellow
            Write-Host "  /config [sub]      Edit config file (see /config help)" -ForegroundColor Yellow
            Write-Host "  /db [sub]          Vector DB management (see /db help)" -ForegroundColor Yellow
            Write-Host "  /run [action] [k=v]  Execute action directly (see /run help)" -ForegroundColor Yellow
            Write-Host "  /help              Show this help"                 -ForegroundColor Yellow
            Write-Host "  -----------------------------------------------" -ForegroundColor DarkGray
            Write-Host ""
            return $true
        }
    }
    return $false
}

# ─────────────────────────────────────────────────────────────
# [10] Banner
# ─────────────────────────────────────────────────────────────

function Show-Banner {
    Write-Host ""
    Write-Host "  +==========================================+" -ForegroundColor Cyan
    Write-Host "  |     A I W O R K S . C O D E             |" -ForegroundColor Cyan
    Write-Host "  |     Local AI Agent for Windows PS 5.1   |" -ForegroundColor DarkCyan
    Write-Host "  +==========================================+" -ForegroundColor Cyan
    Write-Host "  /help for commands  |  /exit to quit"          -ForegroundColor DarkGray
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────
# [11] Load action modules
# ─────────────────────────────────────────────────────────────

function Import-ActionModules {
    $actionsDir = Join-Path $PSScriptRoot "actions"
    if (-not (Test-Path $actionsDir)) {
        Write-AgentLog "actions/ directory not found: $actionsDir" -Type Warning
        return
    }
    $modules = Get-ChildItem -Path $actionsDir -Filter "action.*.ps1" -File
    foreach ($mod in $modules) {
        try {
            . $mod.FullName
            Write-AgentLog "Loaded module: $($mod.Name)" -Type System
        }
        catch {
            Write-AgentLog "Failed to load module $($mod.Name): $($_.Exception.Message)" -Type Error
        }
    }
}

# ─────────────────────────────────────────────────────────────
# [12] Main REPL loop
# ─────────────────────────────────────────────────────────────

function Start-AgentREPL {
    Import-Config      -Path $script:ConfigPath
    Import-ActionModules

    Clear-Host
    Show-Banner

    # Authenticate
    Get-AuthToken

    $promptLabel = Coalesce $script:Config["UI.Prompt"] (Coalesce $script:Config["Prompt"] "AIWORKS")
    Write-AgentLog "Session started. Enter natural language commands." -Type System
    Write-Host ""

    while ($script:SessionActive) {
        Write-Host "  $promptLabel> " -NoNewline -ForegroundColor Green
        $userInput = Read-Host

        if ([string]::IsNullOrWhiteSpace($userInput)) { continue }

        # Slash commands
        if ($userInput.StartsWith("/")) {
            $handled = Invoke-SlashCommand -Cmd $userInput
            if (-not $handled) {
                Write-AgentLog "Unknown command. Type /help for help." -Type Warning
            }
            continue
        }

        # Vector DB lookup
        $dbMatch = Search-VectorDB -Query $userInput
        if ($dbMatch) {
            Write-Host ""
            Write-Host "  [DB] Match found - executing cached action." -ForegroundColor Green
            $actionObj = $dbMatch.action
            $parsed = [PSCustomObject]@{
                action               = $actionObj.action
                params               = $actionObj.params
                message              = $dbMatch.response
                requires_confirmation = $false
            }
            Invoke-AgentAction -UserInput $userInput -Parsed $parsed
            continue
        }

        # AI call
        Write-Host ""
        Show-Thinking -Label "Processing"

        $raw = Send-AIRequest -UserInput $userInput
        if (-not $raw) { continue }

        $parsed = Parse-AIResponse -Raw $raw
        $result = Invoke-AgentAction -UserInput $userInput -Parsed $parsed

        # Offer to save to DB
        if ($result -ne $null -and $parsed.action -ne "answer") {
            Write-Host "  Save this action to DB? [Y/N] " -NoNewline -ForegroundColor DarkGray
            $saveAns = Read-Host
            if ($saveAns -match "^[Yy]$") {
                $actionToSave = [PSCustomObject]@{
                    action = $parsed.action
                    params = $parsed.params
                }
                Add-VectorDBEntry -Query $userInput -ActionObj $actionToSave -Response (Coalesce $parsed.message "")
            }
        }
    }

    Write-Host ""
    Write-Host "  AIWORKS session ended." -ForegroundColor DarkGray
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────
if ($env:AIWORKS_TEST_MODE -ne "1") {
    Start-AgentREPL
}
