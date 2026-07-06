param(
    [string]$XlsPath,
    [string]$OutFile
)

try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Write-Error "Excel is not installed or COM registration failed."
    exit 1
}

try { $excel.Visible = $false } catch { }
try { $excel.DisplayAlerts = $false } catch { }

try {
    $wb = $excel.Workbooks.Open($XlsPath, 0, $true)
    $lines = @()

    for ($s = 1; $s -le $wb.Sheets.Count; $s++) {
        $ws = $wb.Sheets.Item($s)
        $lines += "--- Sheet: $($ws.Name) ---"

        $usedRange = $ws.UsedRange
        if ($usedRange) {
            $rows = $usedRange.Rows.Count
            $cols = $usedRange.Columns.Count
            $lines += "  ($rows rows x $cols cols)"

            $maxRows = [Math]::Min($rows, 100000)
            for ($r = 1; $r -le $maxRows; $r++) {
                $rowText = @()
                for ($c = 1; $c -le $cols; $c++) {
                    $cell = $ws.Cells.Item($r, $c).Text
                    $rowText += $cell.Trim() -replace '\r?\n', ' '
                }
                $lines += ($rowText -join " | ")
            }

            if ($rows -gt 100000) {
                $lines += "  [Warning: Only first 100000 rows shown (total: $rows)]"
            }
        }
        $lines += ""
    }

    $wb.Close($false)
    $lines -join "`n" | Out-File -FilePath $OutFile -Encoding UTF8
    Write-Host "Output written to: $OutFile"
} catch {
    $lines = @()
    $lines += "[Error: Failed to open Excel file]"
    $lines += $_.Exception.Message
    $lines -join "`n" | Out-File -FilePath $OutFile -Encoding UTF8
    Write-Error $_.Exception.Message
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
