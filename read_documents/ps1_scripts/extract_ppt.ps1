param(
    [string]$PptPath,
    [string]$OutFile
)

try {
    $ppt = New-Object -ComObject PowerPoint.Application
} catch {
    Write-Error "PowerPoint is not installed or COM registration failed."
    exit 1
}

# Suppress Visible error — some Office versions block this
try { $ppt.Visible = 0 } catch { }

try {
    $pres = $ppt.Presentations.Open($PptPath, 0, 0, 0)
    $lines = @()

    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $slide = $pres.Slides.Item($i)
        $lines += "--- Slide $i ---"

        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1) {
                $text = $shape.TextFrame.TextRange.Text
                if ($text.Trim()) {
                    $lines += $text
                }
            }
            # Check for tables
            if ($shape.HasTable -eq -1) {
                $table = $shape.Table
                for ($r = 1; $r -le $table.Rows.Count; $r++) {
                    $rowText = @()
                    for ($c = 1; $c -le $table.Columns.Count; $c++) {
                        $cellText = $table.Cell($r, $c).Shape.TextFrame.TextRange.Text
                        $rowText += $cellText.Trim() -replace '\r?\n', ' '
                    }
                    $lines += ($rowText -join " | ")
                }
            }
        }
        $lines += ""
    }

    $pres.Close()
    $lines -join "`n" | Out-File -FilePath $OutFile -Encoding UTF8
    Write-Host "Output written to: $OutFile"
}
catch {
    $lines = @()
    $lines += "[Error: Failed to open PowerPoint file]"
    $lines += $_.Exception.Message
    $lines -join "`n" | Out-File -FilePath $OutFile -Encoding UTF8
    Write-Error $_.Exception.Message
}
finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
