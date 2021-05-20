<#
.SYNOPSIS
    Install built-in apps (Sophos) from Windows 10.
.DESCRIPTION
    This script was created to install Sophos Endpoint remotely 
.EXAMPLE
    .\InstalacaoSophosPowerShell.ps1
.NOTES
    FileName:    InstalacaoSophosPowerShell.ps1
    Author:      Rafael Carvalho
    Created:     2021-04-26
    Version history:
    1.0.0 - (2021-04-26) Initial script updated with help section and a fix for randomly freezing
    
#>
# Path for the workdir
$workdir = "C:\Sophos\"

$sixtyFourBit = Test-Path -Path "C:\Program Files"

$SophosInstalled = Test-Path -Path "C:\Program Files\Sophos"

If ($SophosInstalled){
Write-Host "Sophos Already Installed!"
} ELSE {
Write-Host "Begining the installation"

# Check if work directory exists if not create it

If (Test-Path -Path $workdir -PathType Container){
Write-Host "$workdir already exists" -ForegroundColor Green
} ELSE {
New-Item -Path $workdir -ItemType directory
}

# Download the installer

$source = "https://dzr-api-amzn-us-west-2-fa88.api-upe.p.hmr.sophos.com/api/download/affffec0bccbdb0e100d7962c1c7b7dd/SophosSetup.exe"
$destination = "$workdir\SophosSetup.exe"

# Check if Invoke-Webrequest exists otherwise execute WebClient

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (Get-Command 'Invoke-Webrequest'){
Invoke-WebRequest $source -OutFile $destination
} else {
$WebClient = New-Object System.Net.WebClient
$webclient.DownloadFile($source, $destination)
}

# Start the installation
Start-Process -FilePath "$workdir\SophosSetup.exe" -ArgumentList "--quiet"

Start-Sleep -s 360

Start-Process -FilePath "C:\Program Files\Sophos\Sophos UI\Sophos UI.exe" -ArgumentList "/AUTO"
}
