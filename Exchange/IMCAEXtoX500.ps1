[cmdletbinding()]
param (
    [Parameter(Position=0,Mandatory)]
    [string]
    $IMCEAEX
)
$IMCEAEX = ($IMCEAEX).Replace("IMCEAEX-", "")
$IMCEAEX = ($IMCEAEX).Replace("_", "/")
$IMCEAEX = ($IMCEAEX).Replace("+20", " ")
$IMCEAEX = ($IMCEAEX).Replace("+28", "(")
$IMCEAEX = ($IMCEAEX).Replace("+29", ")")
$IMCEAEX = ($IMCEAEX).Replace("+2E", ".")

Write-Host ""
Write-Host "- Converted to X500"
"X500:$($IMCEAEX)" | clip
Write-Host "- Copied to Clipboard"
Write-Host ""
Return "X500:$($IMCEAEX)"