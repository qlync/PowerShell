<#
    .SYNOPSIS
    Поиск сообщений в логах протоколов на серверах Exchange

    .DESCRIPTION
    Скрипт парсит логи транспорта на серверах Exchange на основании указанных параметров

    .PARAMETER Server
    Указывает имя сервера Exchange для фильтрации логов. По умолчанию проверяются все доступные сервера Exchange в организации.

    .PARAMETER IP
    Фильтрует логи по IP-адресу.

    .PARAMETER Sender
    Указывает адрес почты отправителя.

    .PARAMETER Recipient
    Указывает адрес почты получателя.

    .PARAMETER ID
    Идентификатор сообещения в протокол логе

    .PARAMETER Optional
    Поиск по прочим критериям, например InternetMessageId. Совместимо с использованием ключей -Date и -Server.

    .PARAMETER Date
    Фильтрует логи по дате создрания файла в формате yyyy-MM-dd (пр. 2024-10-25)

    .EXAMPLE
    .\Find-ReceiveLogs.ps1 -Sender user@domain.com -Date "2024-10-25"
    Запускает поиск сообщений, отправленных пользователем user@domain.com, в логах созданных 25 октября 2024 года.

    .NOTES
    Требуется установленный модуль управления Exchange PowerShell (EMS).

    .LINK
    Get-Help
#>

param (
    [switch]$Help,
    [string]$Server,
    [string]$IP,
    [string]$Sender,
    [string]$Recipient,
    [string]$ID,
    [string]$Date,
    [string]$Optional
)

if ($Help) {
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host "Find-ReceiveLogs    -   PowerShell script to search for specific log entries in Exchange Server protocol logs."
    Write-Host ""
    Write-Host "SYNTAX:" -ForegroundColor Green
    Write-Host "    .\Find-ReceiveLogs.ps1 [-Server <string>] [-Sender <string>] [-Recipient <string>] [-Date <string>] [-Optional <string>]"
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Green
    Write-Host "    This script searches Exchange protocol logs for specified criteria."
    Write-Host "    Parameters:"
    Write-Host "        -Server         : Specify the Exchange server to target."
    Write-Host "        -Sender         : Filter results by the sender email address."
    Write-Host "        -Recipient      : Filter results by the recipient email address."
    Write-Host "        -Optional       : Search string on logs by specified criteria (example: InternetMessageId). Can be combined with keys -Date and -Server."
    Write-Host "        -Date           : Filter results by the date of the log file (format: yyyy-MM-dd. Example: 2024-10-25)."
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Green
    Write-Host ".\Find-ReceiveLogs.ps1 -Sender user@domain.com -Date '2024-10-25' -Server EXCH1"
    Write-Host "Запускает поиск сообщения, отправленных пользователем user@domain.com, в логах созданных 25 октября 2024 года. Поиск логов осуществляется только на сервере EXCH1."
    Write-Host ""
    Write-Host ".\Find-ReceiveLogs.ps1 -Sender user@domain.com"
    Write-Host "Запускает поиск всех сообщений, отправленных пользователем user@domain.com на всех серверам по всем датам."
    Write-Host ""
    Write-Host ".\Find-ReceiveLogs.ps1 -Date '2024-10-25' -Server EXCH1 -Optional '123123123.123123.12312312@EXCH1'"
    Write-Host "Запускает поиск сообщений, по совпадению заданного слова (InternetMessageId), в логах созданных 25 октября 2024 года, на сервере EXCH1."
    Write-Host "=====================================" -ForegroundColor Green
    return
}

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

if ($Optional) {
    function Find-ReceiveLogs {
        param (
            [string]$Server,
            [string]$Date,
            [string]$Optional
        )
        $exchangeServers = if ($Server) {
            Get-TransportService -Identity $Server | Select-Object -Property Identity, ReceiveProtocolLogPath
        }
        else {
            Get-TransportService | Select-Object -Property Identity, ReceiveProtocolLogPath
        }

        foreach ($exchangeServer in $exchangeServers) {
            $serverName = $exchangeServer.Identity
            $logPath = $exchangeServer.ReceiveProtocolLogPath
            $logPath = $logPath -replace '^C:', 'c$'
            $fullPath = "\\$servername\$logPath"
            $fileList = Get-ChildItem -Path $fullPath -Recurse
    
            foreach ($line in $fileList) {
                if ($Date -and $file.CreationTime.ToString("yyyy-MM-dd") -ne $Date) {
                    continue
                }
    
                $content = Get-Content -Path $file.FullName -Raw
                if ($content -match $Optional) {
                    Write-Host "Exchange Server :           $serverName"
                    Write-Host "Searched File   :           $($file.FullName)" -ForegroundColor Green
                    Write-Host "Date creation   :           $($file.CreationTime)"
                    Write-Host "Finded word     :           $Optional"
                    Write-Host "====================================="
                    return
                }
            }
        }
    }
    Find-ReceiveLogs -Server $Server -Date $Date -Optional $Optional
    return
}

function Find-ReceiveLogs {
    param (
        [string]$Server,
        [string]$IP,
        [string]$Sender,
        [string]$Recipient,
        [string]$ID,
        [string]$Date
    )

    $exchangeServers = if ($Server) {
        Get-TransportService -Identity $Server | Select-Object -Property Identity, ReceiveProtocolLogPath
    }
    else {
        Get-TransportService | Select-Object -Property Identity, ReceiveProtocolLogPath
    }

    foreach ($exchangeServer in $exchangeServers) {
        $serverName = $exchangeServer.Identity
        $logPath = $exchangeServer.ReceiveProtocolLogPath
        $logPath = $logPath -replace '^C:', 'c$'
        $fullPath = "\\$servername\$logPath"
        $fileList = Get-ChildItem -Path $fullPath -Recurse

        foreach ($line in $fileList) {
            if ($Date -and $file.CreationTime.ToString("yyyy-MM-dd") -ne $Date) {
                continue
            }

            $content = Get-Content -Path $file.FullName
            $messageInfo = @{}

            foreach ($line in $content) {
                $fields = $line -split ','

                if ($fields.Length -lt 7) { continue }
                $logDate = $fields[0].Trim()
                $logServerRore = $fields[1].Trim()
                $logMessageID = $fields[2].Trim()
                $logIP = $fields[4].Trim()

                $mailFrom = $null
                $rcptTo = @()
                $ehlo = $null

                if ($line -match 'MAIL FROM:<([^>]+)') {
                    $mailFrom = $matches[1]
                    if (-not $messageInfo.ContainsKey($logMessageID)) {
                        $messageInfo[$logMessageID] = @{
                            EHLO      = $null;
                            MAIL_FROM = $mailFrom;
                            RCPT_TO   = @();
                            LogDate   = $logDate;
                            Server    = $serverName;
                            File      = $file.FullName
                        }
                    }
                    else {
                        $messageInfo[$logMessageID].MAIL_FROM = $mailFrom
                    }
                }

                if ($line -match 'RCPT TO:<([^>]+)') {
                    $rcptTo = $matches[1]
                    if ($messageInfo.ContainsKey($logMessageID)) {
                        $messageInfo[$logMessageID].RCPT_TO += $rcptTo
                    }
                }

                if ($line -match 'EHLO \[(.+?)\]') {
                    $ehlo = $matches[1].Trim()
                    if (-not $messageInfo.ContainsKey($logMessageID)) {
                        $messageInfo[$logMessageID] = @{
                            EHLO      = $ehlo;
                            MAIL_FROM = $null;
                            RCPT_TO   = @();
                            LogDate   = $logDate;
                            Server    = $serverName;
                            File      = $file.FullName
                        }
                    }
                    else {
                        $messageInfo[$logMessageID].EHLO = $ehlo
                    }
                }
            }

            foreach ($msgId in $messageInfo.Keys) {
                $msgDetails = $messageInfo[$msgId]
                $matchFound = $true
                if ($Sender -and $msgDetails.MAIL_FROM -notlike "*$Sender*") {
                    $matchFound = $false
                }

                if ($Recipient) {
                    $recipientMatch = $false
                    foreach ($rcpt in $msgDetails.RCPT_TO) {
                        if ($rcpt -like "*$Recipient*") {
                            $recipientMatch = $true
                            break
                        }
                    }

                    if (-not $recipientMatch) {
                        $matchFound = $false
                    }
                }

                if ($Recipient -and -not $msgDetails.RCPT_TO -contains $Recipient) {
                    $matchFound = $fasle
                }

                if (-not $matchFound) {
                    continue
                }

                Write-Host "Exchange Server $($msgDetails.Server)"
                Write-Host "File: $($msgDetails.File)"
                Write-Host "Log Date : $($msgDetails.LogDate)"
                Write-Host "Matching Entry : MessageID: $msgId"

                if ($msgDetails.MAIL_FROM) {
                    Write-Host "Sender      : $($msgDetails.MAIL_FROM)"
                }
                if ($msgDetails.EHLO) {
                    Write-Host "EHLO        : $($msgDetails.EHLO)"
                }
                else {
                    Write-Host "EHLO: Not found"
                }
                if ($msgDetails.RCPT_TO.Count -gt 0) {
                    Write-Host "Recipient   : $($msgDetails.RCPT_TO -join ', ')"
                }

                Write-Host "====================================="
            }
        }
    }
}

Find-ReceiveLogs -Server $Server -IP $IP -Sender $Sender -Recipient $Recipient -ID $ID -Date $Date