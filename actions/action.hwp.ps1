# ============================================================
#  AIWORKS.CODE - Action Module: hwp
#  Controls Hangul HWP word processor via COM.
#
#  Expected $Params fields:
#    action   : "open" | "read" | "close" | "new" | "save" | "pdf"
#    path     : file path (string)
#    contents : text to write on new (optional, new action only)
#               Each line may begin with a meta tag to set font/size:
#                 [font=FaceName,size=N]line text
#               Default: font=휴먼명조체, size=16 (pt)
# ============================================================

# HWP 기본 폰트 상수 (Unicode 코드포인트로 정의 — 인코딩 독립적)
# 휴(D734) 먼(BA39) 명(BA85) 조(C870) 체(CCB4)
$script:HwpDefaultFont = [char[]](0xD734,0xBA3C,0xBA85,0xC870,0xCCB4) -join ''

# ---------------------------------------------------------
#  Helper: parse per-line font/size meta tag
#    Input  : raw line string
#    Output : PSCustomObject { Font, SizeHwp(1/100pt), SizePt, Text }
# ---------------------------------------------------------
function global:Parse-HwpLineMeta {
    param(
        [string]$Line,
        [string]$DefaultFont   = $null,
        [int]   $DefaultSizePt = 0
    )

    if ([string]::IsNullOrEmpty($DefaultFont)) { $DefaultFont   = $script:HwpDefaultFont }
    if ($DefaultSizePt -eq 0)                  { $DefaultSizePt = 16 }

    $font   = $DefaultFont
    $sizePt = $DefaultSizePt
    $text   = $Line

    # Tag format: [font=FaceName,size=N]  (spaces around = allowed)
    if ($Line -match '^\[font\s*=\s*([^\],]+),\s*size\s*=\s*(\d+)\](.*)$') {
        $font   = $matches[1].Trim()
        $sizePt = [int]$matches[2]
        $text   = $matches[3]
    }

    return [PSCustomObject]@{
        Font    = $font
        SizeHwp = $sizePt * 100   # HWP unit: 1/100 pt
        SizePt  = $sizePt
        Text    = $text
    }
}

# ---------------------------------------------------------
#  Helper: write contents into an open HWP COM object
# ---------------------------------------------------------
function global:Write-HwpContents {
    param(
        [object]$Hwp,
        [string]$Contents,
        [string]$DefaultFont   = $null,
        [int]   $DefaultSizePt = 0
    )

    if ([string]::IsNullOrEmpty($DefaultFont)) { $DefaultFont   = $script:HwpDefaultFont }
    if ($DefaultSizePt -eq 0)                  { $DefaultSizePt = 16 }

    # Normalise line endings then split
    $lines = $Contents -replace "`r`n", "`n" -replace "`r", "`n" -split "`n"

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $meta = Parse-HwpLineMeta -Line $lines[$i] `
                                  -DefaultFont $DefaultFont `
                                  -DefaultSizePt $DefaultSizePt

        # Apply character shape (font + size)
        $Hwp.HAction.GetDefault("CharShape", $Hwp.HParameterSet.HCharShape.HSet)
        $Hwp.HParameterSet.HCharShape.FaceNameHangul = $meta.Font
        $Hwp.HParameterSet.HCharShape.FaceNameLatin  = $meta.Font
        $Hwp.HParameterSet.HCharShape.FaceNameOther  = $meta.Font
        $Hwp.HParameterSet.HCharShape.Height         = $meta.SizeHwp
        $Hwp.HAction.Execute("CharShape", $Hwp.HParameterSet.HCharShape.HSet)

        # Insert text
        $Hwp.HAction.GetDefault("InsertText", $Hwp.HParameterSet.HInsertText.HSet)
        $Hwp.HParameterSet.HInsertText.Text = $meta.Text
        $Hwp.HAction.Execute("InsertText", $Hwp.HParameterSet.HInsertText.HSet)

        # Paragraph break between lines (not after the last line)
        if ($i -lt $lines.Count - 1) {
            $Hwp.HAction.Run("BreakPara")
        }
    }
}

function global:Invoke-Action-hwp {
    param([object]$Params)

    $action       = if ($Params.action)       { $Params.action }       else { "open" }
    $path         = if ($Params.path)         { $Params.path }         else { "" }
    $contents     = if ($Params.contents)     { $Params.contents }     else { "" }
    $contentsPath = if ($Params.contentsPath) { $Params.contentsPath } else { "" }

    # contentsPath가 지정된 경우 파일에서 contents 읽기 (contents보다 우선)
    if ($contentsPath -ne "") {
        if (-not (Test-Path $contentsPath)) {
            Write-AgentLog "hwp: contentsPath not found: $contentsPath" -Type Error
            return "Error: contentsPath not found: $contentsPath"
        }
        try {
            $contents = [System.IO.File]::ReadAllText($contentsPath, [System.Text.Encoding]::UTF8)
            Write-AgentLog "hwp: contents loaded from file: $contentsPath ($($contents.Length) chars)" -Type System
        }
        catch {
            Write-AgentLog "hwp: failed to read contentsPath: $($_.Exception.Message)" -Type Error
            return "Error: failed to read contentsPath: $($_.Exception.Message)"
        }
    }

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

                if ($contents -ne "") {
                    Write-HwpContents -Hwp $hwp -Contents $contents
                    Write-AgentLog "New HWP document created with contents." -Type Success
                    return "New HWP document created with contents."
                }

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
