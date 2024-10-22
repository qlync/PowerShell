<#
        .SYNOPSIS
        This script created and sets an email signature for an Outlook profile based on user attributes from Active Directory

        .DESCRIPTION
        The script retreives user attributes from Active Directory (LDAP) to generate a personalized email signature in HTM format.
        It checks existing signatures and updated or sets a new signature if necessary. The script also logs all actions to a log file and removes logs before running script.

        .INPUTS
        None. You cannot pipe input ti this script.

        .OUTPUTS
        None. The script writes output to the console and logs actions in a log file.

        .EXAMPLE
        PS> .\Set-OutlookSignature.ps1
        This will create or update the user's email signature and log the actions.

        .NOTES
        The script should be run with appropriate permissions to modify Outlook profiles and registry settings.
        User permissions is enough.
#>
<#
        .SYNOPSIS
        InitializeParameters.

        .DESCRIPTION
        Initialize parameters for email signature generation and logging.
        This function sets up the default parameters such as signature name, display name, paths for signature storage, image storage, and log file location.
        It helps in organizing and managing the configuration values used in the script.

        .PARAMETER signatureName
        The name of the signature file to be created in the Microsoft Signatures Folder. Default is "Signature.htm".
        Can be changed.

        .PARAMETER signatureDisplayName
        The display name for the signature that will be set in Outlook. Default is "Signature".
        Can be changed.

        .PARAMETER signaturePath
        The file path where the signature HTM file will be created. Default path is "$env:APPDATA\Microsoft\Signatures\$signatureName".
        Can't be changed.

        .PARAMETER imagePath
        The path to the image (company logo) to be included in the email signature. Default path is "C:\Windows\Picture.png".
        Can be changed.

        .PARAMETER signatureFolder
        The folder where Outlook signatures are stored. Default path is "$env:APPDATA\Microsoft\Signatures".
        Can't be changed.

        .PARAMETER profilesPath
        The path to the Outlook profiles registry where signature settings will be checked and applied. Default is "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\".
        Can't be changed.

        .PARAMETER logFilePath
        The path where the log file will be stored. Default is "$env:APPDATA\Microsoft\Signatures\script_log.log".
        Can be changed.

        .OUTPUTS
        Returns a hashtable containing the initialized parameters.

        .NOTES
        This function should be called at the beggining of the script to ensure all necessary parameters are available for further processing.
#>
function InitializeParameters {
    $params = @{
        $signatureDisplayName = "Signature"
        $signatureName        = "$signatureDisplayName.htm"
        $signaturePath        = "$env:APPDATA\Microsoft\Signatures\$signatureName"
        $imagePath            = "C:\Windows\Picture.png"
        $signatureFolder      = "$env:APPDATA\Microsoft\Signatures"
        $profilesPath         = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\"
        $logFilePath          = "$env:APPDATA\Microsoft\Signatures\script_log.log"
    }
    return $params
}

<#
        .SYNOPSIS
        Write-LogToFile.

        .DESCRIPTION
        Writes a log entry to the log file.

        .PARAMETER message
        The message to log.

        .OUTPUTS
        None.
#>
function Write-LogToFile {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logEntry
}

<#
        .SYNOPSIS
        Get-LdapUserAttributes.

        .DESCRIPTION
        Retrieves specified attributes from LDAP for a given user.

        .PARAMETER userName
        The user's login name (sAMAccountName)

        .PARAMETER domain
        The domain to connect for LDAP queries.

        .OUTPUTS
        None.
#>
function Get-LdapUserAttributes {
    param (
        [string]$userName,
        [string]$domain
    )

    $ldapConnection = "LDAP://$domain"
    $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]$ldapConnection)
    $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$userName))"
    $searcher.PropertiesToLoad.AddRange(@("displayName", "Name", "Title", "Department", "Phone", "Address", "Company", "Mail", "sAMAccountName"))
    $user = $searcher.FindOne()

    if ($user) {
        $department = $user.Properties["Department"] -split "`n"
        $phone = $user.Properties["Phone"] -split "`n"
        return @{
            DisplayName    = $user.Properties["displayName"][0]
            Name           = $user.Properties["Name"][0]
            Title          = $user.Properties["Title"][0]
            Department     = $department
            Phone          = $phone
            Address        = $user.Properties["Address"][0]
            Company        = $user.Properties["Company"][0]
            Mail           = $user.Properties["Mail"][0]
            sAMAccountName = $user.Properties["sAMAccountName"][0]
        }
    }
    else {
        Write-LogToFile -message "User not found in domain $domain"
        return $null
    }
}

<#
        .SYNOPSIS
        CheckAndSetNewMailSignature.

        .DESCRIPTION
        Checks if the signature need to be set for the user's Outlook account and sets it is required.

        .PARAMETER signatureDisplayName
        The display name of the signature to set.

        .PARAMETER userMail
        The email address of the user for whom to set the signature.

        .OUTPUTS
        None.
#>
function CheckAndSetNewMailSignature {
    param (
        [string]$signatureDisplayName,
        [string]$userMail
    )

    $htmFiles = Get-ChildItem -Path $signatureFolder -Filter "*.htm"
    $standartSignatureExists = Test-Path -Path "$signatureFolder\$signatureName"
    $profileFolders = Get-ChildItem -Path $profilesPath
    if (($htmFiles.Count -eq 0) -or ($htmFiles.Count -eq 1 -and $standartSignatureExists)) {
        foreach ($profileFolder in $profileFolders) {
            $accountFolderPath = "$($profileFolder.PSPath)\9375CFF0413111d3B88A00104B2A6676"
            if (Test-Path -Path $accountFolderPath) {
                $accountKeyPath = "$accountFolderPath\00000002"
                if (Test-Path -Path $accountKeyPath) {
                    $mailSettings = Get-ItemProperty -Path $accountKeyPath
                    $accountName = $mailSettings.'Account Name'
                    if ($accountName -like $userMail) {
                        if (-not $mailSettings.'New Signature') {
                            Set-ItemProperty -Path $accountKeyPath -Name 'New Signature' -Value $signatureDisplayName
                            Write-LogToFile "Signature set for account $accountName"
                        }
                        else {
                            Write-LogToFile "Signature already set for account $accountName"
                        }
                    }
                    else {
                        Write-LogToFile "Email mismatch for account $accountName"
                    }
                }
                else {
                    Write-LogToFile "Key 00000002 not found for profile $($profileFolder.PSChildName)"
                }
            }
            else {
                Write-LogToFile "Registry path 9375CFF0413111d3B88A00104B2A6676 not found for profile $($profileFolder.PSChildName)"
            }
        }
    }
    else {
        Write-LogToFile "No signature or only defauly signature present"
    }
}

<#
        .SYNOPSIS
        GenerateSignature.

        .DESCRIPTION
        Generates an email signature based on user attributes and the specified template.

        .PARAMETER userAttributes
        A hashtable of user attributes retrieved from LDAP.

        .PARAMETER imagePath
        The path to the company logo image to be included in the signature.

        .OUTPUTS
        A string representing the generated HTM signature.
#>
function GenerateSignature {
    param (
        [hashtable]$userAttributes,
        [string]$imagePath
    )
    if ($userAttributes) {
        if ($currentUser -like "external*") {
            $signature = @"
            <html>
            <body style='font-family:Calibri; font-size:12pt'>
            <p style='margin:0; padding:0; '>_________________________________</p>
            <p style='margin:0; padding:0; '>С уважением,</p>
"@
            if ($userAttributes.Name) {
                $signature += "<p style='margin:0; padding:0;'><b><i>$($userAttributes.Name)</i></b></p>"
            }
            $signature += "<p style='margin:0; padding:0;'><i>Сотрудник внешней компании</i></p>"
        }
        else {
            $signature = @"
            <html>
            <body style='font-family:Calibri; font-size:12pt'>
            <p style='margin:0; padding:0; '>_________________________________</p>
            <p style='margin:0; padding:0; '>С уважением,</p>
"@
            if ($userAttributes.Name) {
                $signature += "<p style='margin:0; padding:0;'><b>$($userAttributes.Name)</b></p>"
            }
            else {
                Write-LogToFile "User $($userAttributes.sAMAccountName) have not attribute Name (Фамилия Имя). Attribute Name will not be included in signature"
            }
            if ($userAttributes.Title) {
                $signature += "<p style='margin:0; padding:0;'>$($userAttributes.Title)</p>"
            }
            else {
                Write-LogToFile "User $($userAttributes.sAMAccountName) have not attribute Title (Должность). Attribute Title will not be included in signature"
            }
            if ($userAttributes.Department.Length -ge 1) {
                $signature += "<p style='margin:0; padding:0;'>$($userAttributes.Department[0])</p>"
            }
            else {
                Write-LogToFile "User $($userAttributes.sAMAccountName) have not attribute Department (Департамент1). Attribute Department1 will not be included in signature"
            }
            if ($userAttributes.Department.Length -ge 2) {
                $signature += "<p style='margin:0; padding:0;'>$($userAttributes.Department[1])</p>"
            }
            else {
                Write-LogToFile "User $($userAttributes.sAMAccountName) have not attribute Department (Департамент2). Attribute Department2 will not be included in signature"
            }
            if ($userAttributes.Company) {
                $signature += "<p style='margin:0; padding:0;'>$($userAttributes.Company)</p>"
            }
            else {
                Write-LogToFile "User $($userAttributes.sAMAccountName) have not attribute Company (Компания). Attribute Company will not be included in signature"
            }
            if ($userAttributes.Address) {
                $signature += "<p style='margin:0; padding:0;'>$($userAttributes.Address)</p>"
            }
            else {
                Write-LogToFile "User $($userAttributes.sAMAccountName) have not attribute Address (Адрес). Attribute Address will not be included in signature"
            }
            if ($userAttributes.Department.Length -ge 3) {
                $signature += "<p style='margin:0; padding:0;'>$($userAttributes.Department[2])</p>"
            }
            else {
                Write-LogToFile "User $($userAttributes.sAMAccountName) have not attribute Department (Департамент3). Attribute Department3 will not be included in signature"
            }
            if ($userAttributes.Phone) {
                $signature += "<p style='margin:0; padding:0;'>$($userAttributes.Phone[0])</p>"
            }
            else {
                Write-LogToFile "User $($userAttributes.sAMAccountName) have not attribute Phone (Телефон). Attribute Phone will not be included in signature"
            }
            $signature += "<p style='margin:0; padding:0;'><a href='https://site-your-company.com'>https://site-your-company.com</a></p>"
            $signature += "<p style='margin:0; padding:0; '>_________________________________</p>"
            if (Test-Path -Path $imagePath) {
                $signature += "<p style='margin:0; padding:0; line-height:2;'><img srv='file:///$imagePath' alt='Company Logo' width='370' height='30'></p>"
            }
            else {
                Write-LogToFile "Image not found at the path $imagePath. Logo will not be included in signature"
            }
            $signature += "<p style='margin:0; padding:0; color:green; line-height:2;'>Пожалуйста, подумайте об окружающей среде, прежде чем распечатывать данное письмо</p>"
        }

        $signature += @"
        </body>
        </html>
"@
        return $signature
    }
    else {
        Write-LogToFile "User attributes not available for signature generation"
        return $null
    }
}

<#
        .SYNOPSIS
        ProcessUserSignature

        .DESCRIPTION
        Processes the user's email signature on their attributes.
        This function retrieves the current user's attributes from Active Directory, generates a personalized email signature in HTM format,
        and saves it to the specified signature path. It also sets the new signature in the user's Outlook Profile if it is successfully created.

        .PARAMETER userAttributes
        A hashtable of user attributes retrieved from LDAP, including display name, title, and other relevant information.

        .PARAMETER imagePath
        The path to the company logo image to be included in the signature.

        .OUTPUTS
        None.

        .EXAMPLE
        ProcessUserSignature -userAttributes $userAttributes -imagePath $imagePath

        .NOTES
        This function should be called after user attributes have been successfully retrieved from Active Directory.
#>
function ProcessUserSignature {
    param (
        [string]$imagePath,
        [hashtable]$params
    )
    $currentUser = $env:USERNAME
    $currentDomain = $env:USERDNSDOMAIN

    if (Test-Path $logFilePath) {
        Clear-Content -Path $logFilePath
    }

    Write-LogToFile "Script started for user $currentUser in domain $currentDomain"

    $userAttributes = Get-LdapUserAttributes -userName $currentUser -domain $currentDomain

    $signature = GenerateSignature -userAttributes $userAttributes -imagePath $imagePath

    if ($signature) {
        $signature | Out-File -FilePath $signaturePath -Encoding UTF8
        Write-LogToFile "Signature created at path $signaturePath"
        CheckAndSetNewMailSignature -signatureDisplayName $signatureDisplayName -userMail $userAttributes.Mail
    }
    else {
        Write-LogToFile "Failed to generate signature"
    }
    Write-LogToFile "Script finished"
}

$params = InitializeParameters
ProcessUserSignature -imagePath $imagePath -params $params

