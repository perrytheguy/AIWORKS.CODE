# ============================================================
#  AIWORKS.CODE - Action Module: office
#  Controls Excel, Word, and PowerPoint via COM (OLE).
#
#  Expected $Params fields:
#    app    : "excel" | "word" | "ppt"
#    action : "open" | "read" | "pdf" | "new" | "save" | "close"
#    path   : file path (string)
#    sheet  : sheet name or index (Excel only, optional)
#    range  : cell range e.g. "A1:C5" (Excel only, optional)
#    text   : text to insert (Word only, optional)
# ============================================================

function Invoke-Action-office {
    param([object]$Params)

    $app    = if ($Params.app)    { $Params.app }    else { "" }
    $action = if ($Params.action) { $Params.action } else { "open" }
    $path   = if ($Params.path)   { $Params.path }   else { "" }

    if ($app -eq "") {
        Write-AgentLog "office: 'app' parameter is required (excel/word/ppt)." -Type Error
        return "Error: 'app' parameter missing."
    }

    Write-AgentLog "Office control: [$app] $action => $path" -Type Action

    $comName = switch ($app.ToLower()) {
        "excel"      { "Excel.Application"      }
        "word"       { "Word.Application"        }
        "ppt"        { "PowerPoint.Application"  }
        "powerpoint" { "PowerPoint.Application"  }
        default      {
            Write-AgentLog "Unsupported Office app: $app" -Type Error
            return "Error: unsupported app '$app'."
        }
    }

    $comApp = $null
    try {
        $comApp = New-Object -ComObject $comName
        $comApp.Visible = $true
        $delayMs = [int](Coalesce $script:Config["Office.ComInitDelayMs"] (Coalesce $script:Config["ComInitDelayMs"] "1500"))
        Start-Sleep -Milliseconds $delayMs

        switch ($action.ToLower()) {

            "open" {
                if ($path -eq "") { return "Error: 'path' required for open action." }
                switch ($app.ToLower()) {
                    "excel"      { $null = $comApp.Workbooks.Open($path) }
                    "word"       { $null = $comApp.Documents.Open($path) }
                    { $_ -in @("ppt","powerpoint") } { $null = $comApp.Presentations.Open($path) }
                }
                Write-AgentLog "File opened: $path" -Type Success
                return "File opened: $path"
            }

            "new" {
                switch ($app.ToLower()) {
                    "excel" { $null = $comApp.Workbooks.Add() }
                    "word"  { $null = $comApp.Documents.Add() }
                    { $_ -in @("ppt","powerpoint") } { $null = $comApp.Presentations.Add() }
                }
                Write-AgentLog "New document created in $app" -Type Success
                return "New document created."
            }

            "read" {
                if ($path -eq "") { return "Error: 'path' required for read action." }
                switch ($app.ToLower()) {
                    "excel" {
                        $wb     = $comApp.Workbooks.Open($path)
                        $sheetParam = if ($Params.sheet) { $Params.sheet } else { 1 }
                        $sheet  = $wb.Sheets.Item($sheetParam)
                        $used   = $sheet.UsedRange
                        $rows   = [Math]::Min($used.Rows.Count, 50)
                        $data   = @()
                        for ($r = 1; $r -le $rows; $r++) {
                            $row = @()
                            for ($c = 1; $c -le $used.Columns.Count; $c++) {
                                $row += $sheet.Cells.Item($r, $c).Text
                            }
                            $data += ($row -join "`t")
                        }
                        $wb.Close($false)
                        return $data -join "`n"
                    }
                    "word" {
                        $doc  = $comApp.Documents.Open($path)
                        $text = $doc.Content.Text
                        $doc.Close($false)
                        return $text.Substring(0, [Math]::Min($text.Length, 3000))
                    }
                    { $_ -in @("ppt","powerpoint") } {
                        $prs  = $comApp.Presentations.Open($path)
                        $text = ""
                        for ($s = 1; $s -le $prs.Slides.Count; $s++) {
                            $slide = $prs.Slides.Item($s)
                            for ($sh = 1; $sh -le $slide.Shapes.Count; $sh++) {
                                $shape = $slide.Shapes.Item($sh)
                                if ($shape.HasTextFrame) {
                                    $text += "[Slide $s] " + $shape.TextFrame.TextRange.Text + "`n"
                                }
                            }
                        }
                        $prs.Close()
                        return $text.Substring(0, [Math]::Min($text.Length, 3000))
                    }
                }
            }

            "save" {
                if ($path -eq "") {
                    Write-AgentLog "Save: no path, saving active document." -Type Warning
                    switch ($app.ToLower()) {
                        "excel" { $comApp.ActiveWorkbook.Save() }
                        "word"  { $comApp.ActiveDocument.Save() }
                        { $_ -in @("ppt","powerpoint") } { $comApp.ActivePresentation.Save() }
                    }
                } else {
                    switch ($app.ToLower()) {
                        "excel" {
                            $wb = $comApp.Workbooks.Open($path)
                            $wb.Save()
                            $wb.Close($false)
                        }
                        "word" {
                            $doc = $comApp.Documents.Open($path)
                            $doc.Save()
                            $doc.Close($false)
                        }
                        { $_ -in @("ppt","powerpoint") } {
                            $prs = $comApp.Presentations.Open($path)
                            $prs.Save()
                            $prs.Close()
                        }
                    }
                }
                Write-AgentLog "Document saved." -Type Success
                return "Document saved."
            }

            "pdf" {
                if ($path -eq "") { return "Error: 'path' required for pdf action." }
                $pdfPath = [IO.Path]::ChangeExtension($path, ".pdf")
                switch ($app.ToLower()) {
                    "excel" {
                        $wb = $comApp.Workbooks.Open($path)
                        $wb.ExportAsFixedFormat(0, $pdfPath)
                        $wb.Close($false)
                    }
                    "word" {
                        $doc = $comApp.Documents.Open($path)
                        # 17 = wdFormatPDF
                        $doc.SaveAs([ref]$pdfPath, [ref]17)
                        $doc.Close($false)
                    }
                    { $_ -in @("ppt","powerpoint") } {
                        $prs = $comApp.Presentations.Open($path)
                        # 2 = ppSaveAsPDF
                        $prs.SaveAs($pdfPath, 2)
                        $prs.Close()
                    }
                }
                Write-AgentLog "PDF exported: $pdfPath" -Type Success
                return "PDF saved: $pdfPath"
            }

            "close" {
                switch ($app.ToLower()) {
                    "excel" { try { $comApp.Workbooks | ForEach-Object { $_.Close($false) } } catch {} }
                    "word"  { try { $comApp.Documents | ForEach-Object { $_.Close($false) } } catch {} }
                    { $_ -in @("ppt","powerpoint") } { try { $comApp.Presentations | ForEach-Object { $_.Close() } } catch {} }
                }
                Write-AgentLog "$app documents closed." -Type Success
                return "$app closed."
            }

            default {
                Write-AgentLog "Unknown office action: $action" -Type Warning
                return "Unknown action: $action"
            }
        }
    }
    catch {
        Write-AgentLog "Office error: $($_.Exception.Message)" -Type Error
        return "Error: $($_.Exception.Message)"
    }
    finally {
        try { if ($comApp) { $comApp.Quit() } } catch {}
        if ($comApp) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comApp) | Out-Null } catch {}
        }
    }
}
