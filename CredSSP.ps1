<#
.SYNOPSIS
    Erro ao tentar RDP para uma VM do Windows.
.DESCRIPTION
    Correção oracle de criptografia CredSSP 
.EXAMPLE
    .\CredSSP.ps1
.NOTES
    FileName:    CredSSP.ps1
    Author:      Rafael Carvalho
    Created:     2020-07-21
    Version history:
    1.0.0 - (2020-07-21) Resolvendo o problema definitivamente nas máquinas que geram o erro.
    
#>
REG ADD HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters\ /v AllowEncryptionOracle /t REG_DWORD /d 2
