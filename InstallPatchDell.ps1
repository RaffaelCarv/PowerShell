<#
.SYNOPSIS
    Install built-in apps (Dell) from Windows 10.
.DESCRIPTION
    This script was created to install Patch Dell Security Advisory Update - DSA-2021-088 remotely 
.EXAMPLE
    .\InstallPatchDell.ps1
.NOTES
    FileName:    InstallPatchDell.ps1
    Author:      Rafael Carvalho
    Created:     2021-05-20
    Version history:
    1.0.0 - (2021-05-20) Initial script updated with help section and a fix for randomly freezing
    
#>
# Path for the workdir
$workdir = "C:\Dell\"

$sixtyFourBit = Test-Path -Path "C:\Program Files\Dell"

$DellInstalled = Test-Path -Path "C:\ProgramData\Dell\UpdatePackage\log\Dell-Security-Advisory-Update-DSA-2021-088_7PR57_WIN_1.0.0_A00.txt"

#Verified Manufacturer
$BIOSManufacturer = (get-ciminstance win32_bios).Manufacturer
$status = $BIOSManufacturer -contains 'Dell Inc.'

#If it's a Dell
If ($status -eq ‘true’){
Write-Host "It's a Dell Computer!" -ForegroundColor Green
} ELSE {
Write-Host "Not a Dell" -ForegroundColor Red
Exit
}

If ($DellInstalled){
Write-Host "Patch Already Installed!" -ForegroundColor Yellow
Exit
} ELSE {
Write-Host "Begining the installation" -ForegroundColor Green

# Check if work directory exists if not create it

If (Test-Path -Path $workdir -PathType Container){
Write-Host "$workdir already exists" -ForegroundColor Green
} ELSE {
New-Item -Path $workdir -ItemType directory
}

# Download the installer

$source = "https://dl.dell.com/FOLDER07312946M/1/Dell-Security-Advisory-Update-DSA-2021-088_7PR57_WIN_1.0.0_A00.EXE"
$destination = "$workdir\Dell-Security-Advisory-Update-DSA-2021-088_7PR57_WIN_1.0.0_A00.EXE"

# Check if Invoke-Webrequest exists otherwise execute WebClient

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (Get-Command 'Invoke-Webrequest'){
Invoke-WebRequest $source -OutFile $destination
} else {
$WebClient = New-Object System.Net.WebClient
$webclient.DownloadFile($source, $destination)
}

# Start the installation
Start-Process -FilePath "$workdir\Dell-Security-Advisory-Update-DSA-2021-088_7PR57_WIN_1.0.0_A00.EXE" -ArgumentList '/s'
}