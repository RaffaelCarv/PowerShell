Import-Csv "FilePath"\ContasServico.csv |
foreach {
         $OU = "$($_.Name)"
         $name = "Name"
		 New-ADOrganizationalUnit -Name "$name" -Path "OU=$OU,DC=,DC=" `
         -Verbose

}
