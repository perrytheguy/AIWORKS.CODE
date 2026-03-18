$env:AIWORKS_TEST_MODE = "1"
. (Join-Path $PSScriptRoot "AIWORKS.code.ps1")
function Write-AgentLog { param([string]$Message, [string]$Type="Info") }

$TempDb = Join-Path $env:TEMP "test_db_debug.db"
"[]" | Set-Content -Path $TempDb -Encoding UTF8
$script:DbPath = $TempDb

$actionObj = [PSCustomObject]@{ action = "chrome"; params = @{} }
Add-VectorDBEntry -Query "open portal" -ActionObj $actionObj -Response "opened"

Write-Host "File content: $(Get-Content $TempDb -Raw -Encoding UTF8)"

$db = Load-VectorDB
Write-Host "db type: $($db.GetType().FullName)"
Write-Host "db count: $($db.Count)"
Write-Host "db is null: $($null -eq $db)"

Remove-Item $TempDb -ErrorAction SilentlyContinue
