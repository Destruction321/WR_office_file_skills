param(
    [string]$DocPath,
    [string]$OutFile
)

try {
    $word = New-Object -ComObject Word.Application
} catch {
    Write-Error "Word is not installed or COM registration failed."
    exit 1
}

try { $word.Visible = $false } catch { }

try {
    $doc = $word.Documents.Open($DocPath, $false, $true, $false)
    $lines = @()

    foreach ($para in $doc.Paragraphs) {
        $text = $para.Range.Text
        # Remove trailing CR (Word uses \r as paragraph separator)
        $text = $text -replace '\r$', ''
        if (-not $text.Trim()) { continue }
        $lines += $text
    }

    # Tables
    foreach ($table in $doc.Tables) {
        $lines += ""
        for ($r = 1; $r -le $table.Rows.Count; $r++) {
            $rowText = @()
            for ($c = 1; $c -le $table.Columns.Count; $c++) {
                $cellText = $table.Cell($r, $c).Range.Text
                $cellText = $cellText -replace '\r\s*$', ''
                $rowText += $cellText.Trim() -replace '\r?\n', ' '
            }
            $lines += ($rowText -join " | ")
        }
    }

    $doc.Close($false)
    $lines -join "`n" | Out-File -FilePath $OutFile -Encoding UTF8
    Write-Host "Output written to: $OutFile"
} catch {
    $lines = @()
    $lines += "[Error: Failed to open Word file]"
    $lines += $_.Exception.Message
    $lines -join "`n" | Out-File -FilePath $OutFile -Encoding UTF8
    Write-Error $_.Exception.Message
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
