Import-Module ActiveDirectory
Get-ADUser -Filter * | foreach {Set-ADUser $_ -Replace @{c="BR";co="Brasil";countrycode=76}}
Get-ADUser -Filter {UserPrincipalName -like "*@OldSuffix"} -SearchBase "OU=,DC=,DC=" |
ForEach-Object {
    $UPN = $_.UserPrincipalName.Replace("OldSuffix","NewSuffix")
    Set-ADUser $_ -UserPrincipalName $UPN #-Verbose -WhatIf
}
