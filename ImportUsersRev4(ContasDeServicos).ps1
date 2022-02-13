cls
Import-Module ActiveDirectory
$Arquivo = Read-Host "Nome do Arquivo"
Import-Csv C:\Users\rafael.carvalho\Desktop\Brasoftware\Users\ContasDeServico\$Arquivo.csv |
foreach {
         $name = "$($_.SamAccountName) Service"
		 $secpass = ConvertTo-SecureString -AsPlainText $($_.Password) -Force
         New-ADUser -GivenName $($_.FirstName) -Surname $($_.LastName) `
         -Name $name -SamAccountName "$($_.SamAccountName)" `
         -UserPrincipalName "$($_.SamAccountName)@ADPMS.local" `
         -AccountPassword $secpass -Path "OU=Contas de Servico,OU=Usuarios,OU=$Arquivo,OU=PMS,DC=ADPMS,DC=local" `
         -Enabled:$true -Verbose #-WhatIf

}