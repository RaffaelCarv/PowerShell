cls
Import-Module ActiveDirectory
$Arquivo = Read-Host "Nome do grupo"
Import-Csv C:\Users\rafael.carvalho\Desktop\Brasoftware\Groups\Lista\ListaDeDistribuicao\NovaPasta\$Arquivo.csv |
foreach {
        $User = $($_.Membros)
        $Group = Get-ADGroup -Identity "CN=$Arquivo,OU=ListaDeDistribuicao,OU=Zimbra,OU=PMS,DC=ADPMS,DC=local"
        Add-ADGroupMember -Identity $Group -Members $User -Verbose #-WhatIf

}