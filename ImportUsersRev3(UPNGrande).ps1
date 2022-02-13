Import-Module ActiveDirectory
Import-Csv ListaUsuariosFinalUPNGrande.csv |
foreach {
         $name = "$($_.FirstName) $($_.LastName)"
	 $secpass = ConvertTo-SecureString -AsPlainText $($_.Password) -Force
         New-ADUser -GivenName $($_.FirstName) -Surname $($_.LastName) `
         -Name $name -SamAccountName "$($_.FirstName).$($_.UPN2000) " `
         -UserPrincipalName "$($_.SamAccountName)@Suffix" `
         -AccountPassword $secpass -Path "DC=,DC=" `
         -Enabled:$true -Verbose -WhatIf

}
