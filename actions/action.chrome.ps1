# ============================================================
#  AIWORKS.CODE - Action Module: chrome
#  Controls Google Chrome browser.
#  Strategy: Start-Process for simple URL opening.
#  For scripted automation, uses Playwright CLI (if available)
#  or ChromeDriver path from config.
#
#  Expected $Params fields:
#    action  : "open" | "script" | "navigate" | "screenshot"
#    url     : target URL (string)
#    script  : JavaScript code snippet (string, script action)
#    output  : output file path (string, screenshot action)
#    profile : Chrome profile directory (optional)
# ============================================================

function global:Invoke-Action-chrome {
    param([object]$Params)

    $action  = if ($Params.action)  { $Params.action }  else { "open" }
    $url     = if ($Params.url)     { $Params.url }     else { "" }
    $jsCode  = if ($Params.script)  { $Params.script }  else { "" }
    $output  = if ($Params.output)  { $Params.output }  else { "" }
    $profile = if ($Params.profile) { $Params.profile } else { "" }

    Write-AgentLog "Chrome control: $action => $url" -Type Action

    $chromePath = Coalesce $script:Config["Browser.ChromeExePath"] (Coalesce $script:Config["ChromeExePath"] "")
    # Fallback: look for chrome in PATH
    if ($chromePath -eq "" -or -not (Test-Path $chromePath)) {
        $chromePath = "chrome"
    }

    switch ($action.ToLower()) {

        "open" {
            if ($url -eq "") { return "Error: 'url' required for open action." }
            $args = @($url)
            if ($profile -ne "") { $args += "--profile-directory=`"$profile`"" }
            try {
                if ($chromePath -ne "chrome" -and (Test-Path $chromePath)) {
                    Start-Process -FilePath $chromePath -ArgumentList $args
                } else {
                    Start-Process "chrome" -ArgumentList $args
                }
                Write-AgentLog "Chrome opened: $url" -Type Success
                return "Chrome opened: $url"
            }
            catch {
                Write-AgentLog "Chrome launch error: $($_.Exception.Message)" -Type Error
                return "Error: $($_.Exception.Message)"
            }
        }

        "navigate" {
            # Alias for open
            if ($url -eq "") { return "Error: 'url' required for navigate action." }
            $args = @($url)
            try {
                if ($chromePath -ne "chrome" -and (Test-Path $chromePath)) {
                    Start-Process -FilePath $chromePath -ArgumentList $args
                } else {
                    Start-Process "chrome" -ArgumentList $args
                }
                Write-AgentLog "Chrome navigated: $url" -Type Success
                return "Chrome navigated: $url"
            }
            catch {
                Write-AgentLog "Chrome navigate error: $($_.Exception.Message)" -Type Error
                return "Error: $($_.Exception.Message)"
            }
        }

        "script" {
            # Requires Node.js + Playwright to be installed
            $playwrightCLI = Coalesce $script:Config["Browser.PlaywrightCLI"] (Coalesce $script:Config["PlaywrightCLI"] "npx playwright")
            if ($url -eq "" -and $jsCode -eq "") {
                return "Error: 'url' and/or 'script' required for script action."
            }

            $safeUrl     = $url -replace "'", "\'"
            $tempScript  = [IO.Path]::Combine([IO.Path]::GetTempPath(), "aiworks_chrome_$(Get-Date -Format 'yyyyMMddHHmmss').js")

            $nodeScript = @"
const { chromium } = require('playwright');
(async () => {
  const browser = await chromium.launch({ headless: false });
  const page = await browser.newPage();
  try {
    await page.goto('$safeUrl');
    $jsCode
  } catch(e) {
    console.error('Script error:', e.message);
  } finally {
    await browser.close();
  }
})();
"@
            Set-Content -Path $tempScript -Value $nodeScript -Encoding UTF8

            try {
                $nodeExe = "node"
                $result  = & $nodeExe $tempScript 2>&1
                Remove-Item $tempScript -ErrorAction SilentlyContinue
                $output  = $result -join "`n"
                Write-AgentLog "Chrome script executed." -Type Success
                return $output
            }
            catch {
                Remove-Item $tempScript -ErrorAction SilentlyContinue
                Write-AgentLog "Chrome script error: $($_.Exception.Message)" -Type Error
                return "Error: $($_.Exception.Message)"
            }
        }

        "screenshot" {
            if ($url -eq "") { return "Error: 'url' required for screenshot action." }
            $outPath = if ($output -ne "") { $output } else {
                [IO.Path]::Combine([IO.Path]::GetTempPath(), "aiworks_screenshot_$(Get-Date -Format 'yyyyMMddHHmmss').png")
            }

            $safeUrl  = $url -replace "'", "\'"
            $safePath = $outPath -replace "\\", "\\\\"

            $tempScript = [IO.Path]::Combine([IO.Path]::GetTempPath(), "aiworks_shot_$(Get-Date -Format 'yyyyMMddHHmmss').js")
            $nodeScript = @"
const { chromium } = require('playwright');
(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  await page.goto('$safeUrl');
  await page.screenshot({ path: '$safePath', fullPage: true });
  await browser.close();
  console.log('Screenshot saved: $safePath');
})();
"@
            Set-Content -Path $tempScript -Value $nodeScript -Encoding UTF8

            try {
                $result = & node $tempScript 2>&1
                Remove-Item $tempScript -ErrorAction SilentlyContinue
                Write-AgentLog "Screenshot saved: $outPath" -Type Success
                return "Screenshot saved: $outPath"
            }
            catch {
                Remove-Item $tempScript -ErrorAction SilentlyContinue
                Write-AgentLog "Screenshot error: $($_.Exception.Message)" -Type Error
                return "Error: $($_.Exception.Message)"
            }
        }

        default {
            Write-AgentLog "Unknown Chrome action: $action" -Type Warning
            return "Unknown action: $action"
        }
    }
}
