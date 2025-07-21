<#
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Criacao: 21 de julho de 2022
    Ultima atualizacao: 21 de julho de 2025

    Descricao:
    Este script realiza duas operacoes principais sobre contas de usuario no Active Directory:

    1. Atualiza os atributos de pais (c, co, countryCode) para padrao brasileiro.
    2. Altera o sufixo UPN dos usuarios (a parte apos o @ no UserPrincipalName), com base em um sufixo atual informado e um novo sufixo desejado.

    Funcionamento:
    - O script solicita ao administrador que informe o sufixo UPN atual (ex: @contoso.local) e o novo sufixo desejado (ex: @contoso.com).
    - Todos os usuarios com o sufixo atual informado terao o UPN ajustado para usar o novo sufixo.
    - O script tambem pergunta se deseja apenas simular as alteracoes (modo WhatIf).
    - Um log detalhado e gerado no Desktop somente se as alteracoes forem reais.

    Termos utilizados:
    - Sufixo UPN: parte do UserPrincipalName que vem apos o caractere @ (ex: em usuario@contoso.local, o sufixo UPN e "@contoso.local").
    - SearchBase: caminho LDAP que define o escopo da busca no AD (ex: OU=Usuarios,DC=empresa,DC=com).
#>

Import-Module ActiveDirectory

# Entrada dos sufixos
$sufixoAtual = Read-Host "Informe o sufixo UPN atual (ex: @dominio.local)"
$sufixoNovo = Read-Host "Informe o novo sufixo UPN desejado (ex: @dominio.com)"

# Confirmacao para simular
$simular = Read-Host "Deseja apenas simular as alteracoes? (S/N)"
$whatIf = $simular -match "^[Ss]"

# Opcional: SearchBase (comente se quiser todo o AD)
$searchBase = Read-Host "Deseja definir uma OU especifica para a busca? (Deixe em branco para todo o dominio)"
$usuarios = if ($searchBase) {
    Get-ADUser -Filter * -SearchBase $searchBase -Properties UserPrincipalName,c,co,countryCode
} else {
    Get-ADUser -Filter * -Properties UserPrincipalName,c,co,countryCode
}

# Filtro baseado no sufixo atual
$usuariosAlvo = $usuarios | Where-Object { $_.UserPrincipalName -like "*$sufixoAtual" }

if ($usuariosAlvo.Count -eq 0) {
    Write-Host "`nNenhum usuario encontrado com sufixo UPN $sufixoAtual" -ForegroundColor Yellow
    return
}

# Lista para log
$log = @()
$alteracoes = 0

foreach ($usuario in $usuariosAlvo) {
    $novoUPN = $usuario.UserPrincipalName -replace [regex]::Escape($sufixoAtual), $sufixoNovo

    Write-Host "`nUsuario: $($usuario.SamAccountName)" -ForegroundColor Cyan
    Write-Host "UPN atual : $($usuario.UserPrincipalName)"
    Write-Host "Novo UPN : $novoUPN"
    Write-Host "c / co / countryCode serao ajustados para BR / Brasil / 76"

    if (-not $whatIf) {
        try {
            Set-ADUser -Identity $usuario.DistinguishedName `
                       -UserPrincipalName $novoUPN `
                       -Replace @{ c = "BR"; co = "Brasil"; countryCode = 76 }

            $log += @(
                "Usuario: $($usuario.SamAccountName)",
                "UPN antigo: $($usuario.UserPrincipalName)",
                "UPN novo  : $novoUPN",
                "Atributos pais atualizados para BR/Brasil/76",
                "--------------------------------------------"
            )
            $alteracoes++
            Write-Host "Alteracoes aplicadas com sucesso." -ForegroundColor Green
        }
        catch {
            $log += @(
                "Usuario: $($usuario.SamAccountName)",
                "Erro ao aplicar alteracoes: $_",
                "--------------------------------------------"
            )
            Write-Host "Erro ao aplicar alteracoes: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "(Simulacao - nenhuma alteracao aplicada)" -ForegroundColor Yellow
    }
}

# Salvar log se houve alteracoes
if (-not $whatIf -and $alteracoes -gt 0) {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $caminho = "$([Environment]::GetFolderPath("Desktop"))\log_alteracao_upn_$timestamp.txt"
    [System.IO.File]::WriteAllLines($caminho, $log, [System.Text.Encoding]::UTF8)
    Write-Host "`n*** Log salvo em: $caminho ***" -ForegroundColor Yellow
} elseif ($whatIf) {
    Write-Host "`nSimulacao finalizada. Nenhum log foi gerado." -ForegroundColor Gray
} else {
    Write-Host "`nNenhuma alteracao foi registrada. Nenhum log gerado." -ForegroundColor Gray
}
