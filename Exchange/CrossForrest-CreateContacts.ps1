$hostname = 'IP-Address-Remote-Domain-Controller'
$passwd = 'PASSWORD'
$username = 'DOMAIN\LDAPuser'

$ldap = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$hostname", $username, $passwd)

try {
    $userSearcher = New-Object System.DirectoryServices.DirectorySearcher($ldap)
    $userSearcher.Filter = "(&(objectClass=user)(userAccountControl=512)(mail=*))"
    $userSearcher.Filter = "(&(objectClass=user)(|(userAccountControl=512)(userAccountControl=66048)(userAccountControl=544))(mail=*)(!msExchRecipientTypeDetails=549755813888))"
    $userSearcher.PropertiesToLoad.AddRange(@("displayName", "mail", "mailNickname", "department", "Title", "company", "telephoneNumber", "Manager", "Office", "DistinguishedName"))
    $userResults = $userSearcher.FindAll()

    $groupSearcher = New-Object System.DirectoryServices.DirectorySearcher($ldap)
    $groupSearcher.Filter = "(&(objectClass=group)(mail=*))"
    $groupSearcher.PropertiesToLoad.AddRange(@("displayName", "mail", "mailNickname", "department", "Title", "company", "telephoneNumber", "Manager", "Office", "DistinguishedName", "GroupType"))
    $groupResults = $groupSearcher.FindAll()

    $sharedSearcher = New-Object System.DirectoryServices.DirectorySearcher($ldap)
    $sharedSearcher.Filter = "(&(objectClass=user)(msExchRecipientTypeDetails=4))"
    $sharedSearcher.PropertiesToLoad.AddRange(@("displayName", "mail", "mailNickname", "department", "Title", "company", "telephoneNumber", "Manager", "Office", "DistinguishedName"))
    $sharedResults = $sharedSearcher.FindAll()

    $results = $userResults + $groupResults + $sharedResults
    if ($results.Count -gt 0) {

        $allContacts = @{}
        $existingContacts = Get-MailContact -ResultSize Unlimited | Where-Object { $_.OrganizationalUnit -eq "DOMAIN.COM/DOMAIN/Contacts/CrossForrestContacts" }
        foreach ($existingContact in $existingContacts) {
            $existingPrimarySMTPAddress = $existingContact.PrimarySMTPAddress
            $allContacts[$existingPrimarySMTPAddress] = $existingContact
        }

        $contactsToDelete = @()

        foreach ($result in $results) {
            $displayName = $result.Properties["displayName"][0]
            $primarySMTPAddress = $result.Properties["mail"][0]
            $alias = $result.Properties["mailNickname"][0]
            $department = $result.Properties["department"][0]
            $title = $result.Properties["Title"][0]
            $company = $result.Properties["company"][0]
            $telephoneNumber = $result.Properties["telephoneNumber"][0]
            $GroupType = $result.Properties["GroupType"][0]
            $managerDN = $result.Properties["Manager"][0]
            $manager = $managerDN -replace '^CN=([^,]+),.*', '$1'



            if ($allContacts.ContainsKey($primarySMTPAddress)) {
                $existingContact = $allContacts[$primarySMTPAddress]

                $allContacts.Remove($primarySMTPAddress)

                Set-Contact -Identity $existingContact.Identity -DisplayName $displayName -Department $department -Title $title -Company $company -Phone $telephoneNumber
                $getuscontact = Get-Contact $manager | Select-Object Identity

                if (($getuscontact -ne "$null") -AND ($getuscontact -ne "")) {
                    Set-Contact -Identity $existingContact.Identity -Manager $getuscontact.Identity
                }

                if (($GroupType -ne "$null") -AND ($GroupType -ne "")) {
                    $displayName = $displayName + " (Company)"
                    Set-Contact -Identity $existingContact.Identity -DisplayName $displayName 
                }

            }
            else {
                if (($primarySMTPAddress -ne "$null") -or ($primarySMTPAddress -ne "")) {
                    New-MailContact -Name $displayName -DisplayName $displayName -ExternalEmailAddress $primarySMTPAddress -Alias $alias -OrganizationalUnit "DOMAIN.COM/DOMAIN/Contacts/CrossForrestContacts"
                    Set-Contact -Identity $primarySMTPAddress -DisplayName $displayName -Department $department -Title $title -Company $company -Phone $telephoneNumber

                    $getuscontactn = Get-Contact $manager | Select-Object Identity

                    if (($getuscontactn -ne "$null") -AND ($getuscontactn -ne "")) {
                        Set-Contact -Identity $existingContact.Identity -Manager $getuscontactn.Identity
                    }

                    if (($GroupType -ne "$null") -AND ($GroupType -ne "")) {
                        $displayName = $displayName + " (Company)"
                        Set-Contact -Identity $primarySMTPAddress -DisplayName $displayName 
                    }
                }
            }
        }

        $contactsToDelete += $allContacts.Values
        #$contactsToDelete | Remove-Contact
        foreach ($contactsToHide in $contactsToDelete) {
            Set-MailContact $contactsToHide.primarySMTPAddress -HiddenFromAddressListsEnabled $True
        }
        if ($results.Count -gt 0) {
            $outputData = @()

            foreach ($result in $results) {
                $displayName = $result.Properties["displayName"][0]
                $primarySMTPAddress = $result.Properties["mail"][0]
                $alias = $result.Properties["mailNickname"][0]
                $department = $result.Properties["department"][0]
                $title = $result.Properties["Title"][0]
                $company = $result.Properties["company"][0]
                $telephoneNumber = $result.Properties["telephoneNumber"][0]
                $managerDN = $result.Properties["Manager"][0]
                $manager = $managerDN -replace '^CN=([^,]+),.*', '$1' 


                $properties = @{
                    "DisplayName"        = $displayName
                    "PrimarySMTPAddress" = $primarySMTPAddress
                    "Alias"              = $alias
                    "Department"         = $department
                    "Title"              = $title
                    "Company"            = $company
                    "TelephoneNumber"    = $telephoneNumber
                }

                $outputData += New-Object PSObject -Property $properties
            }

        }
    }
    else {
        Write-Host "No results found in LDAP."
    }
}
catch {
    throw "Problem looking up model account - $($_.Exception.Message)"
}
finally {
    $ldap.Dispose()
}
