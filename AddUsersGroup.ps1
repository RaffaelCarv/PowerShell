Import-Module ActiveDirectory
$Arquivo = Read-Host "Nome do grupo"
Import-Csv "file path"\$Arquivo.csv |
foreach {
        $User = $($_.Membros)
        $Group = Get-ADGroup -Identity "CN=$Arquivo,OU=,DC=,DC="
        Add-ADGroupMember -Identity $Group -Members $User -Verbose -WhatIf

}
