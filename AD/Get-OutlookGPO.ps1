$reportFolder = "C:\Users\user\Desktop\OutlookGPO"
if (-not (Test-Path $reportFolder)) {
    New-Item -Path $reportFolder -ItemType Directory
}
$allGPOs = Get-GPO -All
foreach ($gpo in $allGPOs) {
    $gpoName = $gpo.DisplayName
    $gpoID = $gpo.Id
    $gpoReport = Get-GPOReport -Guid $gpoID -ReportType xml
    if ($gpoReport -like "*Outlook*") {
        $htmlReportPath = Join-Path -Path $reportFolder -ChildPath "$($gpoName)_report.html"
        Get-GPOReport -Guid $gpoID -ReportType html -Path $htmlReportPath
    }
}