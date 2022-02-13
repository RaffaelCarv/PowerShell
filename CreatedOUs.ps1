Import-Csv C:\Users\rafael.carvalho\Desktop\Brasoftware\Groups\OUs.csv |
foreach {
         $name = "$($_.Name)"
		 New-ADOrganizationalUnit -Name "$name" -Path "OU=PMS,DC=ADPMS,DC=local" `
         -Verbose

}