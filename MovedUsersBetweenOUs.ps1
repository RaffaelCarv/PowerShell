﻿# Import AD Module
Import-Module ActiveDirectory

#Variable for filename
$reference = Read-Host "Nome do Arquivo"
 
# Import the data from CSV file and assign it to variable
$Import_csv = Import-Csv -Path "$reference.csv"

# Specify target OU where the users will be moved to
$TargetOU = "DC=,DC="
 
$Import_csv | ForEach-Object {

    # Retrieve DN of User
    $UserDN = (Get-ADUser -Identity $_.SamAccountName).distinguishedName

    Write-Host "Moving Accounts....."

    # Move user to target OU. Remove the -WhatIf parameter after you tested.
    Move-ADObject -Identity $UserDN -TargetPath $TargetOU -WhatIf
} 
Write-Host "Completed move"
