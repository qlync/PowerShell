function Send-Notification {
    param (
        [string]$to,
        [string]$subject,
        [string]$body
    )
    $smtpServer = "smtp.example.com"
    $smtpFrom = "noreply@example.com"
    $message = New-Object System.Net.Mail.MailMessage $smtpFrom, $to, $subject, $body
    $smtp = New-Object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($message)
}

$mailbox = "user1@example.com"
$startDate = Get-Date
$endDate = (Get-Date).AddDays(60)
$organizerEmail = "user@example.com"
$alertedMeetingsFile = "C:\Temp\alertedMeetings.txt"

Add-Type -Path "C:\Program Files\Microsoft\Exchahge Server\V15\Bin\Microsoft.Exchange.WebServices.dll"
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$service.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$service.Url = New-Object Uri("https://exchange-server/EWS/Exchange.asmx")
$service.ImpersonatedUserId
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailbox)

if (Test-Path $alertedMeetingsFile) {
    $alertedMeetings = Get-Content $alertedMeetingsFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
}
else {
    $alertedMeetings = @()
}

$alerts = @()

try {
    $mailboxId = new-object Microsoft.Exchange.WebServices.Data.MailboxId($mailbox)
    $calendarId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $mailboxId)
    $calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $calendarId)
    $view = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate)
    $findResults = $calendar.FindAppointments($view)

    foreach ($appointment in $findResults.Items) {
        $appointment.Load([Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties)
        if ($appointment.Organizer.Address -eq $organizerEmail) {
            $body = $appointment.Body.Text
            $subject = $appointment.Subject
            $uniqueId = $appointment.Id.ToString()
            $meetingDate = $appointment.Start.ToString("yyy-MM-ddTHH:mm:ss")
            $meetingsRecord = "$uniqueId;$meetingDate"
            if ($body -notmatch "https://s4b.example.com" -and $alertedMeetings -notcontains $meetingsRecord -and $appointment.JoinOnlineMeetingUrl -ne "$null" -and $appointment.Location -like "*Skype Meeting*") {
                $htmlBody = @"
                <html>
                <body>
                <p class=MsoNormal style='text-autospace:none'><span style='font-size:8.0pt;
                color: #404040* >.........................................................................................................................................</span><b><span
                style='font-size: 14.0pt*><0:p></o:p></span></b></p>

                <p class-MsoNormal style-'text-autospace:none*><a name=OutJoinLink></a><a 
                href='$(Sappointment.JoinOnlineMeetingUrl)'><span style=*mso-bookmark:
                OutJoinLink*><span style='font-size:16.0pt;color:#0866CC*>Join Skype Meeting</span></span><span 
                style=*mso-bookmark:OutJoinLink*></span></a><span style=*mso-bookmark:OutJointink'><span
                style='font-size:14.0pt*>&nbsp; <a name=OutsharedNoteBorder>&nbsp;</a>&nbsp;&bsp;<a
                name-OutSharedNoteLink>&nbsp;</a></span></span><span style='font-size:14.0pt*><o:p></o:p></span></p>

                <p class-MsoNormal style='margin-top:3.0pt;margin-right:0cm;margin-bottom:12.0pt; 
                margin-left:16.0pt;line-height:125%; text-autospace: none* ><span 
                style= 'font-size:10.0pt;line-height:125%*>Trouble Joining? <u><a 
                href= '$($appointment.JoinOnlineMeetingurl)'><span
                style= 'color:#0066CC'>Try Skype Web App</span></a> </u><о:р></о:р></span></p>

                <p class=MsoNormal style='text-autospace:none*><span lang=EN style='font-size:
                8.0pt;mso-ansi-language: EN'><o:p>&nbsp;</o:p></span></p>
                <p class-MsoNormal style= margin-bottom:10.0pt;line-height:115%; text-autospace: 
                none'><span lang=EN style='font-size:8.0pt;line-height:115%;color:#404040;
                mso-ansi-language: EN'>.........................................................................................................................................</span><b><span
                lang-EN style-'font-size: 10. 5pt; line-height :115%;mso-ansi-language: EN**<0:p></0:p></span></p>

                </body>
                </html>
                "@

                $appointment.Body = New-object Microsoft.Exchange.WebServices.Data.MessageBody([Microsoft.Exchange.WebServices.Data.BodyType]::HTML, $htmlBody)
                $appointment.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverWrite)
                $alerts += "Meeting from $organizerEmail with missing Skype link: $subject on $($appointment.Start). Missed Link: $($appointment.JoinOnlineMeetingUrl)"
                $alertedMeetings += $meetingsRecord
                Add-Content -Path $alertedMeetingsFile -Value $meetingRecord
"@
            }
        }
    }

    if ($alerts.Count -gt 0) {
        $messageBody = $alerts -join ("`n" + "`n")
        Send-Notification -to "admin@example.com" -subject "Missing Skype Link Alerts" -body $messageBody
        Write-Host "$($alerts.Count)"
    }
}
catch {
    Write-Host "Error: $_"
}