<#
    .SYNOPSIS
        Script para atualizacao em lote do UserPrincipalName no Active Directory.
        Permite definir um dominio atual e um novo, atualizar UPNs com opcao de simulacao,
        aplicar restricao por OU e registrar alteracoes em log.

    .CRIACAO
        12 de maio de 2022

    .ULTIMA ATUALIZACAO
        21 de julho de 2025

    .AUTOR
        Rafael Carvalho - GitHub: https://github.com/RaffaelCarv/PowerShell
#>

Import-Module ActiveDirectory

# Define atributos padroes de pais
$atributosPais = @{
    c = "BR"
    co = "Brasil"
    countrycode = 76
}

Write-Host "Atualizando atributos de localizacao padroes em todos os usuarios..." -ForegroundColor Cyan

# Atualiza todos os usuarios com os atributos de pais padrao
Get-ADUser -Filter * -Properties c | ForEach-Object {
    try {
        Set-ADUser $_ -Replace $atributosPais -ErrorAction Stop
    } catch {
        Write-Host "Erro ao atualizar atributos de $_ : $_" -ForegroundColor Red
    }
}

Write-Host "Atributos atualizados com sucesso.`n" -ForegroundColor Green

# Prompt: simulacao
$simulacao = Read-Host "Deseja simular as alteracoes sem aplica-las? (S/N)"
$usarWhatIf = if ($simulacao -match '^[sS]') { $true } else { $false }

# Prompt: definir escopo
$usarOU = Read-Host "Deseja limitar a alteracao a uma OU especifica? (S/N)"
if ($usarOU -match '^[sS]') {
    $searchBase = Read-Host "Informe o caminho da OU (ex: OU=TI,DC=contoso,DC=local)"
} else {
    $searchBase = $null
}

# Prompt: domínios
$dominioAtual = Read-Host "Informe o dominio atual do UPN (ex: contoso.local)"
$novoDominio  = Read-Host "Informe o novo dominio desejado (ex: contoso.com)"

# Busca usuarios com base na presenca do SearchBase
if ($searchBase) {
    $usuarios = Get-ADUser -Filter "UserPrincipalName -like '*@$dominioAtual'" -SearchBase $searchBase -Properties UserPrincipalName, SamAccountName
} else {
    $usuarios = Get-ADUser -Filter "UserPrincipalName -like '*@$dominioAtual'" -Properties UserPrincipalName, SamAccountName
}

if ($usuarios.Count -eq 0) {
    Write-Host "Nenhum usuario encontrado com o sufixo $dominioAtual." -ForegroundColor Yellow
    return
}

# Lista para log
$logAlteracoes = @()

Write-Host "`nIniciando processamento de $($usuarios.Count) usuarios..." -ForegroundColor Yellow

foreach ($usuario in $usuarios) {
    $UPNnovo = $usuario.UserPrincipalName.Replace($dominioAtual, $novoDominio)

    if ($usarWhatIf) {
        Write-Host "Simulando: $($usuario.UserPrincipalName) -> $UPNnovo" -ForegroundColor Cyan
        $logAlteracoes += "Simulacao: $($usuario.SamAccountName): $($usuario.UserPrincipalName) => $UPNnovo"
    } else {
        try {
            Set-ADUser -Identity $usuario.DistinguishedName -UserPrincipalName $UPNnovo -ErrorAction Stop
            $logAlteracoes += "$($usuario.SamAccountName): $($usuario.UserPrincipalName) => $UPNnovo"
            Write-Host "Alterado: $($usuario.UserPrincipalName) -> $UPNnovo" -ForegroundColor Green
        } catch {
            $logAlteracoes += "ERRO: $($usuario.SamAccountName): $($_.Exception.Message)"
            Write-Host "Erro ao alterar $($usuario.SamAccountName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Gera log
$dataLog = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$caminhoDesktop = [Environment]::GetFolderPath("Desktop")
$arquivoLog = "$caminhoDesktop\Alteracoes_UPN_$dataLog.txt"

if ($logAlteracoes.Count -gt 0) {
    [System.IO.File]::WriteAllLines($arquivoLog, $logAlteracoes, [System.Text.Encoding]::UTF8)
    Write-Host "`nLog salvo em: $arquivoLog" -ForegroundColor Yellow
} else {
    Write-Host "`nNenhuma alteracao realizada, nenhum log gerado." -ForegroundColor Yellow
}

Write-Host "`nProcessamento concluido." -ForegroundColor Cyan
