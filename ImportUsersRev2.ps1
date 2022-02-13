cls
Import-Module ActiveDirectory
Import-Csv C:\Users\rafael.carvalho\Desktop\Brasoftware\Users\ListaUsuariosFinalUPNConforme.csv |
foreach {
         $name = "$($_.FirstName) $($_.LastName)"
		 $secpass = ConvertTo-SecureString -AsPlainText $($_.Password) -Force
         New-ADUser -GivenName $($_.FirstName) -Surname $($_.LastName) `
         -Name $name -SamAccountName "$($_.SamAccountName)" `
         -UserPrincipalName "$($_.SamAccountName)@ADPMS.local" `
         -AccountPassword $secpass -Path "OU=Zimbra,OU=PMS,DC=ADPMS,DC=local" `
         -Enabled:$true -Verbose #-WhatIf

}