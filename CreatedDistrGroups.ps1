cls
Import-Module ActiveDirectory
Import-Csv C:\Users\rafael.carvalho\Desktop\Brasoftware\Groups\Lista\ListagemGrupo.csv |
foreach {
         New-ADGroup -Name $($_.Nome) -GroupCategory Distribution `
         -Path "OU=ListaDeDistribuicao,OU=Zimbra,OU=PMS,DC=ADPMS,DC=local" `
         -GroupScope Global `
         -Verbose #-WhatIf

}