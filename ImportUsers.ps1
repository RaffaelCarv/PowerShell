$secpass = Read-Host "Password" -AsSecureString
Import-Csv ListaInseridasNoADCSV.csv |
foreach {
         $name = "$($_.FirstName) $($_.LastName)"
         New-ADUser -GivenName $($_.FirstName) -Surname $($_.LastName) `
         -Name $name -SamAccountName $($_.SamAccountName) `
         -UserPrincipalName "$($_.SamAccountName)@Suffix" `
         -AccountPassword $secpass -Path "DC=,DC=" `
         -Enabled:$true
}
