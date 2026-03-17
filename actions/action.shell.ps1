# ============================================================
#  AIWORKS.CODE - Action Module: shell
#  Executes PowerShell commands in a controlled manner.
#
#  Expected $Params fields:
#    command    : PowerShell command or script block (string)
#    workdir    : working directory for the command (optional)
#    timeout    : max execution time in seconds (optional, default 60)
#    capture    : capture output as string (bool, default true)
# ============================================================

function global:Invoke-Action-shell {
    param([object]$Params)

    $command = if ($Params.command)  { $Params.command }   else { "" }
    $workDir = if ($Params.workdir)  { $Params.workdir }   else { "" }
    $timeout = if ($Params.timeout)  { [int]$Params.timeout } else { 60 }
    $capture = if ($null -ne $Params.capture) { $Params.capture } else { $true }

    if ($command -eq "") {
        Write-AgentLog "shell: 'command' parameter is required." -Type Error
        return "Error: 'command' parameter missing."
    }

    Write-AgentLog "Shell exec: $command" -Type Action
    Write-AppLog "Shell command executed: $command" -Level "INFO"

    # Change working directory if specified
    $prevLocation = $null
    if ($workDir -ne "" -and (Test-Path $workDir)) {
        $prevLocation = Get-Location
        Set-Location -Path $workDir
    }

    try {
        # Execute in a script block with timeout via a job
        $job = Start-Job -ScriptBlock {
            param($cmd)
            try {
                $result = Invoke-Expression $cmd 2>&1
                return $result
            }
            catch {
                return "Error: $($_.Exception.Message)"
            }
        } -ArgumentList $command

        $completed = Wait-Job -Job $job -Timeout $timeout
        if ($completed) {
            $output = Receive-Job -Job $job
            Remove-Job  -Job $job -Force

            if ($output -eq $null) {
                return "(command completed with no output)"
            }

            $outputStr = ($output | ForEach-Object { "$_" }) -join "`n"

            # Truncate very long output
            $maxLen = 5000
            if ($outputStr.Length -gt $maxLen) {
                $outputStr = $outputStr.Substring(0, $maxLen) + "`n...(output truncated)"
            }

            Write-AgentLog "Shell command completed." -Type Success
            return $outputStr
        } else {
            Remove-Job -Job $job -Force
            Write-AgentLog "Shell command timed out after ${timeout}s." -Type Warning
            return "Error: Command timed out after ${timeout}s."
        }
    }
    catch {
        Write-AgentLog "Shell error: $($_.Exception.Message)" -Type Error
        return "Error: $($_.Exception.Message)"
    }
    finally {
        if ($prevLocation) {
            Set-Location -Path $prevLocation
        }
    }
}
