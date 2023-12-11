Function Split-Xlsx ($excelFileName, $csvLoc)
{
    $excelFile = $excelFileName
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($excelFile)
    foreach ($ws in $wb.Worksheets)
    {
        $n = $ws.Name
        $ws.SaveAs($csvLoc + $n + ".csv", 6)
    }
    $E.Quit()
    stop-process -processname EXCEL
}
Split-Xlsx("C:\Temp\What.xlsx", 'C:\Temp\')