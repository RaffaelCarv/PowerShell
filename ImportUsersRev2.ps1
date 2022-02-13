Import-Module ActiveDirectory
Import-Csv ListaUsuariosFinalUPNConforme.csv |
foreach {
         $name = "$($_.FirstName) $($_.LastName)"
	 $secpass = ConvertTo-SecureString -AsPlainText $($_.Password) -Force
         New-ADUser -GivenName $($_.FirstName) -Surname $($_.LastName) `
         -Name $name -SamAccountName "$($_.SamAccountName)" `
         -UserPrincipalName "$($_.SamAccountName)@Suffix" `
         -AccountPassword $secpass -Path "DC=,DC=" `
         -Enabled:$true -Verbose -WhatIf

}
