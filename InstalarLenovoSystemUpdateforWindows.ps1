<#
.SYNOPSIS
    Install built-in apps (Lenovo) from Windows 10.
.DESCRIPTION
    This script was created to install Lenovo System Update for Windows remotely 
.EXAMPLE
    .\InstalarLenovoSystemUpdateforWindows.ps1
.NOTES
    FileName:    InstalarLenovoSystemUpdateforWindows.ps1
    Author:      Rafael Carvalho
    Created:     2021-05-21
    Version history:
    1.0.0 - (2021-05-21) Initial script updated with help section and a fix for randomly freezing
    
#>
# Path for the workdir
$workdir = "C:\Lenovo\"

$sixtyFourBit = Test-Path -Path "C:\Program Files\Lenovo"

$LenovoInstalled = Test-Path -Path "C:\Program Files (x86)\Lenovo\System Update"

If ($LenovoInstalled){
Write-Host "App Already Installed!" -ForegroundColor Yellow
Exit
} ELSE {
Write-Host "Begining the installation" -ForegroundColor Green

# Check if work directory exists if not create it

If (Test-Path -Path $workdir -PathType Container){
Write-Host "$workdir Already Exists" -ForegroundColor Green
} ELSE {
New-Item -Path $workdir -ItemType directory
}

# Download the installer

$source = "https://download.lenovo.com/pccbbs/thinkvantage_en/system_update_5.07.0118.exe"
$destination = "$workdir\system_update_5.07.0118.exe"

# Check if Invoke-Webrequest exists otherwise execute WebClient

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (Get-Command 'Invoke-Webrequest'){
Invoke-WebRequest $source -OutFile $destination
} else {
$WebClient = New-Object System.Net.WebClient
$webclient.DownloadFile($source, $destination)
}

# Start the installation
Start-Process -FilePath "$workdir\system_update_5.07.0118.exe" -ArgumentList "/verysilent"
}