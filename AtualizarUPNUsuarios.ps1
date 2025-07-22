<#
    .SYNOPSIS
        Atualiza em lote o sufixo UPN de usuarios no Active Directory.
        Permite simular ou executar a alteracao em toda a floresta ou em uma OU especifica.
        Gera log detalhado da operacao no Desktop.

    .AUTOR
        Rafael Carvalho - GitHub: https://github.com/RaffaelCarv/PowerShell

    .ULTIMA_ATUALIZACAO
        22/07/2025
#>

Import-Module ActiveDirectory -ErrorAction Stop

function Write-LineSeparator {
    Write-Host ("-" * 60) -ForegroundColor DarkGray
}

function Criar-Log {
    param(
        [string]$Caminho,
        [string[]]$Conteudo
    )
    try {
        [System.IO.File]::WriteAllLines($Caminho, $Conteudo, [System.Text.Encoding]::UTF8)
        Write-Host "`nLog salvo em: $Caminho" -ForegroundColor Yellow
    } catch {
        Write-Host "Erro ao salvar o log: $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host ">>> Atualizacao de sufixo UPN no Active Directory <<<" -ForegroundColor Cyan
Write-LineSeparator
Write-Host ""

do {
    $UPNAntigo = Read-Host "Informe o sufixo UPN antigo (ex: @empresa.local)"
} while ([string]::IsNullOrWhiteSpace($UPNAntigo))

Write-Host ""

do {
    $UPNDestino = Read-Host "Informe o novo sufixo UPN (ex: @empresa.com)"
} while ([string]::IsNullOrWhiteSpace($UPNDestino))

Write-Host ""

do {
    Write-Host "Deseja alterar usuarios em toda a floresta/dominio ou limitar a uma OU?" -ForegroundColor White
    Write-Host "  1 - Toda a floresta/dominio" -ForegroundColor Yellow
    Write-Host "  2 - Somente em uma OU especifica" -ForegroundColor Yellow
    $escopoInput = Read-Host "Digite 1 ou 2"
    if ($escopoInput -eq "1") {
        $usarOU = $false
        $searchBase = $null
        break
    }
    elseif ($escopoInput -eq "2") {
        $usarOU = $true
        do {
            $searchBase = Read-Host "Informe o caminho da OU (ex: OU=Usuarios,DC=empresa,DC=local)"
        } while ([string]::IsNullOrWhiteSpace($searchBase))
        break
    }
    else {
        Write-Host "Entrada invalida. Digite 1 ou 2." -ForegroundColor Red
        Write-Host ""
    }
} while ($true)

Write-Host ""

do {
    Write-Host "Escolha o modo:" -ForegroundColor White
    Write-Host "  1 - Simular (nao altera nada)" -ForegroundColor Yellow
    Write-Host "  2 - Executar (altera de fato)" -ForegroundColor Green
    $modoInput = Read-Host "Digite 1 ou 2"
    if ($modoInput -eq "1") { $modo = "SIMULACAO"; break }
    elseif ($modoInput -eq "2") { $modo = "EXECUCAO"; break }
    else {
        Write-Host "Entrada invalida. Digite 1 ou 2." -ForegroundColor Red
        Write-Host ""
    }
} while ($true)

Write-Host ""
Write-LineSeparator
Write-Host "Buscando usuarios no(s) local(is) selecionado(s)..." -ForegroundColor Cyan

$filtro = "UserPrincipalName -like '*$UPNAntigo'"

try {
    if ($usarOU) {
        $Usuarios = Get-ADUser -Filter $filtro -SearchBase $searchBase -Properties UserPrincipalName
    }
    else {
        $Usuarios = Get-ADUser -Filter $filtro -Properties UserPrincipalName
    }
} catch {
    Write-Host "Erro ao buscar usuarios. Verifique a conexao e os parametros." -ForegroundColor Red
    exit
}

if (-not $Usuarios) {
    Write-Host "Nenhum usuario encontrado com o sufixo $UPNAntigo no local selecionado." -ForegroundColor Yellow
    exit
}

Write-Host ""
Write-Host "Total de usuarios encontrados: $($Usuarios.Count)" -ForegroundColor Green
Write-LineSeparator
Write-Host ""

if ($modo -eq "SIMULACAO") {
    Write-Host ">>> SIMULACAO <<<" -ForegroundColor Cyan
    Write-Host ""

    foreach ($User in $Usuarios) {
        $NovoUPN = $User.UserPrincipalName -replace [regex]::Escape($UPNAntigo), $UPNDestino
        Write-Host "[SIMULACAO] $($User.UserPrincipalName)  -->  $NovoUPN" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "Total de contas que seriam alteradas: $($Usuarios.Count)" -ForegroundColor Green
    Write-Host ""
    Write-LineSeparator
    exit
}

elseif ($modo -eq "EXECUCAO") {
    Write-Host ">>> EXECUCAO <<<" -ForegroundColor Yellow
    Write-Host ""

    $DataHora = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $LogPath = [IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "Log_Atualizacao_UPN_$DataHora.txt")
    $Contador = 0
    $Log = @()

    $total = $Usuarios.Count
    $contadorProgresso = 0

    foreach ($User in $Usuarios) {
        $contadorProgresso++
        $NovoUPN = $User.UserPrincipalName -replace [regex]::Escape($UPNAntigo), $UPNDestino
        try {
            Set-ADUser -Identity $User -UserPrincipalName $NovoUPN
            $Contador++
            $Log += "[$Contador] $($User.UserPrincipalName)  -->  $NovoUPN"
            Write-Progress -Activity "Atualizando UPNs" -Status "Processando $contadorProgresso de $total" -PercentComplete (($contadorProgresso / $total) * 100)
        } catch {
            $Log += "[ERRO] $($User.UserPrincipalName)  -->  Falha: $_"
            Write-Host "[ERRO] ao alterar $($User.UserPrincipalName): $_" -ForegroundColor Red
        }
    }

    Write-Progress -Activity "Atualizando UPNs" -Completed

    Write-Host ""
    Write-Host "Total de contas alteradas: $Contador" -ForegroundColor Green
    Write-Host ""

    $escopoTexto = if ($usarOU) { $searchBase } else { 'Toda a floresta/dominio' }

    $LogHeader = @"
###############################################
#   LOG DE ATUALIZACAO DE UPN - EXECUCAO
#   Data: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
#   Escopo: $escopoTexto
#   UPN Antigo: $UPNAntigo
#   UPN Novo: $UPNDestino
#   Total de contas alteradas: $Contador
###############################################

"@

    $LogFinal = $LogHeader + ($Log -join "`r`n")
    Criar-Log -Caminho $LogPath -Conteudo $LogFinal
    Write-LineSeparator
}
