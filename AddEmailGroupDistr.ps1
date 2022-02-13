Get-ADGroup -Filter 'GroupCategory -eq "Distribution"' | ForEach-Object {

    Set-ADGroup -Identity $_.DistinguishedName -Add @{mail="$($_.SamAccountName)@salvador.ba.gov.br"} #-WhatIf

}