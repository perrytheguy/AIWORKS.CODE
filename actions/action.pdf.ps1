# ============================================================
#  AIWORKS.CODE - Action Module: pdf
#  Extracts text from PDF files using pdftotext or fallback.
#
#  Expected $Params fields:
#    path     : path to the PDF file (string)
#    action   : "read" | "info" (default: "read")
#    maxchars : max characters to return (optional, default 3000)
# ============================================================

function global:Invoke-Action-pdf {
    param([object]$Params)

    $path     = if ($Params.path)     { $Params.path }             else { "" }
    $action   = if ($Params.action)   { $Params.action.ToLower() } else { "read" }
    $maxChars = if ($Params.maxchars) { [int]$Params.maxchars }    else { 3000 }

    if ($path -eq "") {
        Write-AgentLog "pdf: 'path' parameter is required." -Type Error
        return "Error: 'path' parameter missing."
    }

    if (-not (Test-Path $path)) {
        Write-AgentLog "pdf: File not found: $path" -Type Error
        return "Error: File not found: $path"
    }

    Write-AgentLog "PDF action: $action => $path" -Type Action

    switch ($action) {

        "read" {
            return Invoke-PdfRead -Path $path -MaxChars $maxChars
        }

        "info" {
            $fileInfo = Get-Item $path
            $info = @(
                "File    : $($fileInfo.FullName)",
                "Size    : $([Math]::Round($fileInfo.Length / 1KB, 2)) KB",
                "Created : $($fileInfo.CreationTime.ToString('yyyy-MM-dd HH:mm:ss'))",
                "Modified: $($fileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
            )
            return $info -join "`n"
        }

        default {
            # Default to read
            return Invoke-PdfRead -Path $path -MaxChars $maxChars
        }
    }
}

function Invoke-PdfRead {
    param([string]$Path, [int]$MaxChars = 3000)

    $tool     = Coalesce $script:Config["Office.PdfExtractTool"] (Coalesce $script:Config["PdfExtractTool"] "pdftotext")
    $toolPath = Coalesce $script:Config["Office.PdfToolPath"]    (Coalesce $script:Config["PdfToolPath"] "")
    $exe      = if ($toolPath -ne "") { $toolPath } else { $tool }
    $outFile  = [IO.Path]::Combine([IO.Path]::GetTempPath(), "aiworks_pdf_$(Get-Date -Format 'yyyyMMddHHmmss').txt")

    try {
        # Try pdftotext (poppler)
        & $exe $Path $outFile 2>&1 | Out-Null

        if (Test-Path $outFile) {
            $text = Get-Content $outFile -Raw -Encoding UTF8
            Remove-Item $outFile -ErrorAction SilentlyContinue
            if ($text -and $text.Trim() -ne "") {
                Write-AgentLog "PDF text extracted via $tool." -Type Success
                return $text.Substring(0, [Math]::Min($text.Length, $MaxChars))
            }
        }
    }
    catch {
        Write-AgentLog "pdftotext not available ($($_.Exception.Message)); trying fallback." -Type Warning
    }
    finally {
        Remove-Item $outFile -ErrorAction SilentlyContinue
    }

    # Fallback: try iTextSharp or .NET PDF reading via Word COM
    try {
        return Invoke-PdfReadViaWord -Path $Path -MaxChars $MaxChars
    }
    catch {
        Write-AgentLog "PDF Word fallback also failed: $($_.Exception.Message)" -Type Warning
    }

    # Last resort: inform user
    Write-AgentLog "PDF text extraction failed. Install pdftotext (poppler) for best results." -Type Error
    return "PDF text extraction failed. Please install pdftotext (poppler utilities) or configure PdfToolPath in config."
}

function Invoke-PdfReadViaWord {
    param([string]$Path, [int]$MaxChars = 3000)

    Write-AgentLog "PDF: attempting extraction via Word COM..." -Type System
    $word = $null
    try {
        $word    = New-Object -ComObject "Word.Application"
        $word.Visible = $false
        # Word can open PDF and expose text content
        $doc     = $word.Documents.Open($Path, $false, $true)
        $text    = $doc.Content.Text
        $doc.Close($false)
        if ($text -and $text.Trim() -ne "") {
            Write-AgentLog "PDF text extracted via Word COM." -Type Success
            return $text.Substring(0, [Math]::Min($text.Length, $MaxChars))
        }
        throw "No text extracted from PDF via Word."
    }
    finally {
        if ($word) {
            try { $word.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
        }
    }
}
