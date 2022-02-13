Import-Module ActiveDirectory
$Arquivo = Read-Host "Nome do Arquivo"
Import-Csv $Arquivo.csv |
foreach {
         $name = "$($_.SamAccountName) Service"
		 $secpass = ConvertTo-SecureString -AsPlainText $($_.Password) -Force
         New-ADUser -GivenName $($_.FirstName) -Surname $($_.LastName) `
         -Name $name -SamAccountName "$($_.SamAccountName)" `
         -UserPrincipalName "$($_.SamAccountName)@ADPMS.local" `
         -AccountPassword $secpass -Path "DC=,DC=" `
         -Enabled:$true -Verbose -WhatIf

}
