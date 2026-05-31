# Export a worksheet range to PNG via Excel COM (visual QA of the dashboard design).
# Sheets are referenced by INDEX (not name) so no Hebrew is typed on the command line
# (Windows PowerShell 5.1 mis-decodes UTF-8 args). Range/paths are ASCII.
#   powershell -File screenshot.ps1 -Xlsx <path> -SheetIndex 1 -Range A1:X66 -Png <out.png>
param(
    [Parameter(Mandatory=$true)][string]$Xlsx,
    [Parameter(Mandatory=$true)][int]$SheetIndex,
    [Parameter(Mandatory=$true)][string]$Range,
    [Parameter(Mandatory=$true)][string]$Png
)
$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($Xlsx)
    $ws = $wb.Worksheets.Item($SheetIndex)
    $ws.Activate() | Out-Null
    $ws.DisplayRightToLeft = $true
    $rng = $ws.Range($Range)
    # copy the range as an on-screen bitmap (includes overlapping charts)
    $rng.CopyPicture(1, 2) | Out-Null    # xlScreen=1, xlBitmap=2
    Start-Sleep -Milliseconds 400
    $co = $ws.ChartObjects().Add(0, 0, $rng.Width, $rng.Height)
    $co.Chart.ChartArea.Format.Line.Visible = $false
    Start-Sleep -Milliseconds 200
    $co.Chart.Paste()
    Start-Sleep -Milliseconds 400
    $co.Chart.Export($Png, "PNG") | Out-Null
    $co.Delete() | Out-Null
    $wb.Close($false)
    Write-Output ("SHOT ok: " + [IO.Path]::GetFileName($Png))
} catch {
    Write-Output ("SHOT ERR: " + $_.Exception.Message)
} finally {
    if ($xl) {
        $xl.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    }
}
exit 0
