cls
Import-Csv C:\Users\rafael.carvalho\Desktop\Brasoftware\Users\ContasServico.csv |
foreach {
         $OU = "$($_.Name)"
         $name = "Contas de Servico"
		 New-ADOrganizationalUnit -Name "$name" -Path "OU=Usuarios,OU=$OU,OU=PMS,DC=ADPMS,DC=local" `
         -Verbose

}