<#
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Criacao: 21 de julho de 2022
    Ultima atualizacao: 21 de julho de 2025

    Descricao:
    Este script realiza duas acoes principais em usuarios do Active Directory:
    1. Atualiza os atributos de pais (`c`, `co`, `countryCode`) para padrao brasileiro.
    2. Atualiza os sufixos de UserPrincipalName (UPN) de um dominio antigo para um novo, com entrada via prompt.
    3. Gera um log das alteracoes reais realizadas.

    Importante:
    - O administrador informa o dominio antigo, novo e o SearchBase.
    - O script pergunta se deve simular ou aplicar as alteracoes.
    - O log so sera gerado se as alteracoes forem reais.
    - Exemplos:
        - Dominio antigo (interno): @contoso.local
        - Dominio novo (nuvem/publico): @contoso.com
#>

# Carrega o modulo do Active Directory
Import-Module ActiveDirectory

# Entrada de dados via prompt
$dominioAntigo = Read-Host "Informe o sufixo antigo do UPN (ex: @contoso.local)"
$dominioNovo   = Read-Host "Informe o novo sufixo do UPN (ex: @contoso.com)"
$searchBase    = Read-Host "Informe o caminho do SearchBase (ex: OU=usuarios,DC=empresa,DC=local)"

# Pergunta se deseja simular
$simular = Read-Host "Deseja apenas simular as alteracoes (S/N)?"

# Define flag de WhatIf
$usarWhatIf = $false
if ($simular.ToUpper() -eq "S") {
    $usarWhatIf = $true
    Write-Host "`n*** Executando em modo de simulacao (WhatIf habilitado) ***`n" -ForegroundColor Yellow
} else {
    Write-Host "`n*** Executando alteracoes reais (WhatIf desabilitado) ***`n" -ForegroundColor Cyan
}

# Lista de mensagens para log
$log = @()

# Atualiza os atributos de pais para todos os usuarios
Get-ADUser -Filter * -Properties c,co,countryCode | ForEach-Object {
    Set-ADUser $_ -Replace @{
        c = "BR";
        co = "Brasil";
        countryCode = 76
    }
}

# Atualiza o sufixo do UserPrincipalName
Get-ADUser -Filter {UserPrincipalName -like "*$dominioAntigo"} -SearchBase $searchBase | ForEach-Object {
    $novoUPN = $_.UserPrincipalName.Replace($dominioAntigo, $dominioNovo)

    if ($usarWhatIf) {
        Set-ADUser $_ -UserPrincipalName $novoUPN -WhatIf -Verbose
    } else {
        try {
            Set-ADUser $_ -UserPrincipalName $novoUPN -Verbose
            $mensagem = "Sucesso: $($_.UserPrincipalName) alterado para $novoUPN"
            $log += $mensagem
            Write-Host $mensagem -ForegroundColor Green
        } catch {
            $erro = "Erro: falha ao atualizar $($_.UserPrincipalName) - $($_.Exception.Message)"
            $log += $erro
            Write-Host $erro -ForegroundColor Red
        }
        $log += ""  # linha em branco entre usuarios
    }
}

# Gera o log somente se for execucao real
if (-not $usarWhatIf -and $log.Count -gt 0) {
    $desktop = [Environment]::GetFolderPath("Desktop")
    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
    $caminhoLog = "$desktop\Log_Atualizacao_UPN_$timestamp.txt"
    [System.IO.File]::WriteAllText($caminhoLog, ($log -join "`n"), [System.Text.Encoding]::UTF8)

    Write-Host "`n`n*** Log gerado: $caminhoLog ***`n" -ForegroundColor Yellow
}
