Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Set-ADServerSettings -ViewEntireForest $true
$ReadHost = Read-Host "Введите PrimarySMTPAddress или sAMAccountName ящика, к которому нужно вывести разрешения"
$choice = Read-Host "Выберите какие права нужно показать (1 - FullAccess, 2 - SendAs)"
switch ($choice) {
    '1' {
        $PermissionsFullAccess = Get-MailboxPermission $ReadHost | Where-Object { ($_.IsInherited -eq $false) -AND ($_.User -notlike "NT AUTHORITY\SELF") -AND ($_.User -notlike "S-1-5*") -AND ($_.AccessRights -like "FullAccess") } | Select-Object User
        $arrayFA = foreach ($PermissionFullAccess in $PermissionsFullAccess) {
            Get-Mailbox "$($PermissionFullAccess.User)" | Select-Object DisplayName, sAMAccountName, PrimarySMTPAddress
        }
        $arrayFA | Out-GridView
    }
    '2' {
        $PermissionsSendAs = Get-MailboxPermission $ReadHost | Get-ADPermission | Where-Object { ($_.IsInherited -eq $false) -AND ($_.User -notlike "NT AUTHORITY\SELF") -AND ($_.User -notlike "S-1-5*") -AND ($_.AccessRights -like "Send-As") } | Select-Object User
        $arraySA = foreach ($PermissionSendAs in $PermissionsSendAs) {
            Get-Mailbox "$($PermissionSendAs.User)" | Select-Object DisplayName, sAMAccountName, PrimarySMTPAddress
        }
        $arraySA | Out-GridView
    }
}