Import-Module ActiveDirectory

Get-ADUser -Filter * | foreach {Set-ADUser $_ -Replace @{c="BR";co="Brasil";countrycode=76}}

Get-ADUser -Filter {UserPrincipalName -like "*@ADPMS.local"} -SearchBase "OU=PMS,DC=ADPMS,DC=local" |
ForEach-Object {
    $UPN = $_.UserPrincipalName.Replace("ADPMS.local","salvador.ba.gov.br")
    Set-ADUser $_ -UserPrincipalName $UPN #-Verbose #-WhatIf
}