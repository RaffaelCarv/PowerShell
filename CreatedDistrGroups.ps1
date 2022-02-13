Import-Module ActiveDirectory
Import-Csv "FilePath"\ListagemGrupo.csv |
foreach {
         New-ADGroup -Name $($_.Nome) -GroupCategory Distribution `
         -Path "DC=,DC=" `
         -GroupScope Global `
         -Verbose -WhatIf

}
