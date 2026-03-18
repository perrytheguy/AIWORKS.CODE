# Convert Test-AIWORKS.ps1 to UTF-8 with BOM
$path = 'C:\Users\student\claude_test\AIWORKS.CODE\Test-AIWORKS.ps1'
$content = [System.IO.File]::ReadAllText($path, (New-Object System.Text.UTF8Encoding($false)))
$utf8bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($path, $content, $utf8bom)
$bytes = [System.IO.File]::ReadAllBytes($path)
$bom = ($bytes[0].ToString('X2') + ' ' + $bytes[1].ToString('X2') + ' ' + $bytes[2].ToString('X2'))
Write-Host ('BOM added. First 3 bytes: ' + $bom)
