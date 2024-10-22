$currentUser = ((Get-WmiObject -Class Win32_ComputerSystem).UserName).Split('\')[1]

$logFilePath = "C:\Users\$currentUser\AppData\Local\Microsoft\Office\16.0\Lync\Tracing"
$logFileMask = "Lync-UccApi-*.UccApilog"
$eventLogSource = "LyncConferenceDeletion"
$existSource = Get-WinEvent -ListLog Application | Select-Object ProviderNames | Where-Object { $_.ProviderNames -like "*$eventLogSource*" }

if ($existSource -eq "$null") {
    New-EventLog -LogName Application -Source $eventLogSource
}

$logFiles = Get-ChildItem -Path $logFilePath -Filter $logFileMask

foreach ($logFile in $logFiles) {
    $date = $null
    $requestId = $null
    $id = $null
    $to = $null
    $from = $null

    Get-Content -Path $logFile.FullName -ReadCount 1 | ForEach-Object {
        $line = $_
        if ($line -match "^(\d{2}\/\d{2}\/\d{4}\|\d{2}:\d{2}:\d{2}\.\d{3})") {
            $date = $matches[1]
        }

        if ($line -match "<deleteConference><conferenceKeys>") {
            $requestId = [regex]::Match($line, 'requestId="([^"]+)"').Groups[1].Value
            $id = [regex]::Match($line, 'id:([^"]+)"').Groups[1].Value
            $to = [regex]::Match($line, 'to=:([^"]+)"').Groups[1].Value
            $from = [regex]::Match($line, 'from=:([^"]+)"').Groups[1].Value

            #$eventId = $id + $requestId
            $eventMessage = "User: $currentUser. Conference deletion request: date = $date id=$id, requestId = $requestId, to=$to, from=$from"

            $existingEvents = Get-EventLog -LogName Application -Source $eventLogSource -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            if (!$existingEvents) {
                Write-EventLog -LogName Application -Source $eventLogSource -EntryType Information -EventId 1 -Message $eventMessage
            }
        }
    }
}