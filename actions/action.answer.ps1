# ============================================================
#  AIWORKS.CODE - Action Module: answer
#  Displays a text reply from the AI to the user.
# ============================================================

function Invoke-Action-answer {
    param([object]$Params)

    # $Params here is the full parsed AI response object (action, params, message, ...)
    $message = ""
    if ($Params -and $Params.message) {
        $message = $Params.message
    } elseif ($Params -and $Params.params -and $Params.params.text) {
        $message = $Params.params.text
    } elseif ($Params -and $Params.params -and $Params.params.message) {
        $message = $Params.params.message
    }

    Write-Host ""

    if ($message -eq "") {
        Write-Host "  (no message)" -ForegroundColor DarkGray
    } else {
        # Word-wrap at ~100 chars for readability
        $maxWidth = 100
        $words    = $message -split " "
        $line     = "  "
        foreach ($word in $words) {
            if (($line + $word).Length -gt $maxWidth) {
                Write-Host $line -ForegroundColor White
                $line = "  $word "
            } else {
                $line += "$word "
            }
        }
        if ($line.Trim() -ne "") {
            Write-Host $line -ForegroundColor White
        }
    }

    Write-Host ""
}
