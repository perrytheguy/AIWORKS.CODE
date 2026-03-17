# ============================================================
#  AIWORKS.CODE - Action Module: hwp
#  Controls Hangul HWP word processor via COM.
#
#  Expected $Params fields:
#    action : "open" | "read" | "close" | "new" | "save" | "pdf"
#    path   : file path (string)
# ============================================================

function global:Invoke-Action-hwp {
    param([object]$Params)

    $action = if ($Params.action) { $Params.action } else { "open" }
    $path   = if ($Params.path)   { $Params.path }   else { "" }

    Write-AgentLog "HWP control: $action => $path" -Type Action

    $hwp      = $null
    $keepOpen = $action.ToLower() -in @("open", "new")
    try {
        $hwp = New-Object -ComObject "HWPFrame.HwpObject"

        $secPath = Coalesce $script:Config["Office.HwpSecurityPath"] (Coalesce $script:Config["HwpSecurityPath"] "")
        if ($secPath -ne "" -and (Test-Path $secPath)) {
            $hwp.XHwpDocuments.RegisterModule("FilePathCheckDLL", $secPath)
        }

        # Make the HWP window visible
        try { $hwp.XHwpWindows.Active_XHwpWindow.Visible = $true } catch {}

        switch ($action.ToLower()) {

            "open" {
                if ($path -eq "") { return "Error: 'path' required for open action." }
                $hwp.Open($path, "HWP", "forceopen:true")
                Write-AgentLog "HWP file opened: $path" -Type Success
                return "HWP opened: $path"
            }

            "new" {
                $hwp.XHwpDocuments.Add($true)
                Write-AgentLog "New HWP document created." -Type Success
                return "New HWP document created."
            }

            "read" {
                if ($path -eq "") { return "Error: 'path' required for read action." }
                $hwp.Open($path, "HWP", "forceopen:true")
                $text = $hwp.GetTextFile("TEXT", "")
                return $text.Substring(0, [Math]::Min($text.Length, 3000))
            }

            "save" {
                if ($path -eq "") {
                    $hwp.HAction.Run("FileSave")
                } else {
                    $hwp.SaveAs($path, "HWP", "")
                }
                Write-AgentLog "HWP document saved." -Type Success
                return "HWP saved."
            }

            "pdf" {
                if ($path -eq "") { return "Error: 'path' required for pdf action." }
                $pdfPath = [IO.Path]::ChangeExtension($path, ".pdf")
                # Open source file first
                $hwp.Open($path, "HWP", "forceopen:true")
                # Export as PDF
                $hwp.SaveAs($pdfPath, "PDF", "")
                Write-AgentLog "HWP PDF exported: $pdfPath" -Type Success
                return "HWP PDF saved: $pdfPath"
            }

            "close" {
                $hwp.Quit()
                Write-AgentLog "HWP closed." -Type Success
                return "HWP closed."
            }

            default {
                Write-AgentLog "Unknown HWP action: $action" -Type Warning
                return "Unknown action: $action"
            }
        }
    }
    catch {
        Write-AgentLog "HWP error: $($_.Exception.Message)" -Type Error
        return "Error: $($_.Exception.Message)"
    }
    finally {
        # Quit only for non-interactive actions (read, save, pdf, close)
        if ($hwp -and -not $keepOpen -and $action.ToLower() -ne "close") {
            try { $hwp.Quit() } catch {}
        }
        if (-not $keepOpen -and $hwp) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null } catch {}
        }
    }
}
