$env:AIWORKS_TEST_MODE = "1"
$MainScript = Join-Path $PSScriptRoot "AIWORKS.code.ps1"
. $MainScript

function Write-AgentLog {
    param([string]$Message, [string]$Type = "Info")
}

$script:Config["Safety.ConfirmKeywords"] = "delete,drop,format,remove,rm,shutdown"
$script:Config["ConfirmKeywords"]        = "delete,drop,format,remove,rm,shutdown"

$safe = [PSCustomObject]@{ requires_confirmation = $false }

# Debug: check what happens inside Test-DangerousAction
function Test-DangerousAction-Debug {
    param([string]$InputStr, [object]$Parsed)
    Write-Host "  Parsed.requires_confirmation: $($Parsed.requires_confirmation)"
    $raw = $script:Config["Safety.ConfirmKeywords"]
    Write-Host "  raw: '$raw'"
    $keywords = $raw -split ","
    Write-Host "  keywords count: $($keywords.Count)"
    foreach ($kw in $keywords) {
        $kw = $kw.Trim()
        $match = $InputStr -match [regex]::Escape($kw)
        Write-Host "  '$kw' match '$InputStr' => $match"
        if ($kw -ne "" -and $match) { return $true }
    }
    return $false
}

Write-Host "=== Using -Input (original param name) ==="
$r1 = Test-DangerousAction -Input "delete file" -Parsed $safe
Write-Host "Result: $r1"

Write-Host ""
Write-Host "=== Using -InputStr (renamed) ==="
$r2 = Test-DangerousAction-Debug -InputStr "delete file" -Parsed $safe
Write-Host "Result: $r2"
