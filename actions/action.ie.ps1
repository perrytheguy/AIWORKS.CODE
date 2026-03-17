# ============================================================
#  AIWORKS.CODE - Action Module: ie
#  Controls Internet Explorer via COM automation.
#
#  Expected $Params fields:
#    action   : "open" | "read" | "input" | "click" | "close" | "wait"
#    url      : target URL (string)
#    selector : element ID for input/click actions (string)
#    value    : value to type into a field (string, input action)
#    timeout  : wait timeout in seconds (optional, default 30)
# ============================================================

function global:Invoke-Action-ie {
    param([object]$Params)

    $action   = if ($Params.action)   { $Params.action }   else { "open" }
    $url      = if ($Params.url)      { $Params.url }      else { "" }
    $selector = if ($Params.selector) { $Params.selector } else { "" }
    $value    = if ($Params.value)    { $Params.value }    else { "" }
    $timeout  = if ($Params.timeout)  { [int]$Params.timeout } else { 30 }

    Write-AgentLog "IE control: $action => $url" -Type Action

    $ie = $null
    try {
        $ie = New-Object -ComObject "InternetExplorer.Application"
        $ie.Visible = $true

        # Helper: wait for page to finish loading
        $waitForLoad = {
            param([int]$MaxSec = 30)
            $elapsed = 0
            while ($ie.Busy -and $elapsed -lt ($MaxSec * 1000)) {
                Start-Sleep -Milliseconds 200
                $elapsed += 200
            }
        }

        switch ($action.ToLower()) {

            "open" {
                if ($url -eq "") { return "Error: 'url' required for open action." }
                $ie.Navigate($url)
                & $waitForLoad $timeout
                Write-AgentLog "IE page loaded: $url" -Type Success
                return "IE opened: $url"
            }

            "read" {
                if ($url -eq "") { return "Error: 'url' required for read action." }
                $ie.Navigate($url)
                & $waitForLoad $timeout
                $body = ""
                try { $body = $ie.Document.Body.InnerText } catch {}
                return $body.Substring(0, [Math]::Min($body.Length, 3000))
            }

            "input" {
                if ($url -ne "") {
                    $ie.Navigate($url)
                    & $waitForLoad $timeout
                }
                if ($selector -eq "") { return "Error: 'selector' (element ID) required for input action." }
                $el = $ie.Document.getElementById($selector)
                if ($el) {
                    $el.value = $value
                    Write-AgentLog "IE input set: #$selector = $value" -Type Success
                    return "Input complete: #$selector"
                } else {
                    Write-AgentLog "IE element not found: #$selector" -Type Warning
                    return "Element not found: #$selector"
                }
            }

            "click" {
                if ($url -ne "") {
                    $ie.Navigate($url)
                    & $waitForLoad $timeout
                }
                if ($selector -eq "") { return "Error: 'selector' (element ID) required for click action." }
                $el = $ie.Document.getElementById($selector)
                if ($el) {
                    $el.click()
                    & $waitForLoad $timeout
                    Write-AgentLog "IE clicked: #$selector" -Type Success
                    return "Clicked: #$selector"
                } else {
                    Write-AgentLog "IE element not found: #$selector" -Type Warning
                    return "Element not found: #$selector"
                }
            }

            "wait" {
                & $waitForLoad $timeout
                Write-AgentLog "IE wait complete." -Type Success
                return "Wait complete."
            }

            "close" {
                $ie.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ie) | Out-Null
                $ie = $null
                Write-AgentLog "IE closed." -Type Success
                return "IE closed."
            }

            default {
                Write-AgentLog "Unknown IE action: $action" -Type Warning
                return "Unknown action: $action"
            }
        }
    }
    catch {
        Write-AgentLog "IE error: $($_.Exception.Message)" -Type Error
        return "Error: $($_.Exception.Message)"
    }
    finally {
        if ($ie -and $action.ToLower() -ne "close") {
            try { $ie.Quit() } catch {}
        }
        if ($ie) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ie) | Out-Null } catch {}
        }
    }
}
