Import-Csv FilePath\OUs.csv |
foreach {
         $name = "$($_.Name)"
		 New-ADOrganizationalUnit -Name "$name" -Path "DC=,DC=" `
         -Verbose

}
