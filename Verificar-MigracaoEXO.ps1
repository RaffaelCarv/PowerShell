<#
    .SYNOPSIS
        Script para análise de batches de migração no Exchange Online.
        Lista batches disponíveis, permite escolher um batch e consultar status e erros.

    .DESCRIPTION
        Conecta ao Exchange Online e lista os batches disponíveis.
        O usuário escolhe o batch desejado.
        Exibe um menu para verificar status, erros, mudar de lote ou sair.
        Gera logs no Desktop (TXT para status, CSV para erros).

    .NOTES
        Autor: Rafael Carvalho
        Data de criação: 30/06/2025
        Última atualização: 01/07/2025
        Requisitos:
          - PowerShell 5.1 ou superior
          - Módulo ExchangeOnlineManagement instalado
          - Permissões adequadas para consultar batches e usuários
#>

function Verificar-StatusDoBatch {
    param ($batchId)

    $statusBatch = Get-MigrationBatch -Identity $batchId

    $statusUsuarios = Get-MigrationUser -BatchId $batchId | ForEach-Object {
        $stats = Get-MigrationUserStatistics -Identity $_.Identity

        $percentual = if ($stats.PercentageComplete -ne $null) { $stats.PercentageComplete } else { 0 }

        [PSCustomObject]@{
            Usuario             = $_.Identity
            Status              = $stats.Status
            Percentual          = $percentual  # valor numerico puro
            BytesTransferidos   = "$($stats.BytesTransferred)"
            TamanhoEstimado     = "$($stats.EstimatedTotalTransferSize)"
            ItensTransferidos   = $stats.SyncedItemCount
            ItensEstimados      = $stats.TotalItemsInSourceMailboxCount
            UltimoSync          = $stats.LastUpdatedTime
        }
    } | Sort-Object Percentual  # ordena crescente (padrão)

    Write-Host "`nStatus geral do batch:" -ForegroundColor Cyan
    $statusBatch | Format-List Identity, Status, TotalCount, InitialSyncDuration, CreationDateTime, LastSyncedTime

    Write-Host "`nStatus detalhado dos usuarios no batch:" -ForegroundColor Cyan
    $statusUsuarios | Select-Object Usuario, Status,
        @{Name="Percentual %"; Expression = { "$($_.Percentual)%" }},
        BytesTransferidos, TamanhoEstimado, ItensTransferidos, ItensEstimados, UltimoSync |
        Format-Table -AutoSize

    $salvar = Read-Host "`nDeseja salvar essas informacoes em um arquivo txt no Desktop? (S/N)"
    if ($salvar -match '^[sS]$') {
        $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
        $desktop = [Environment]::GetFolderPath("Desktop")
        $path = "$desktop\Status_Migracao_$batchId`_$timestamp.txt"

        $logContent = @()
        $logContent += "Status geral do batch '$batchId'"
        $logContent += "-----------------------------------"
        $logContent += ($statusBatch | Format-List Identity, Status, TotalCount, InitialSyncDuration, CreationDateTime, LastSyncedTime | Out-String)

        $logContent += "`nStatus detalhado dos usuarios:"
        $logContent += "-------------------------------"
        $logContent += ($statusUsuarios | Select-Object Usuario, Status,
            @{Name="Percentual %"; Expression = { "$($_.Percentual)%" }},
            BytesTransferidos, TamanhoEstimado, ItensTransferidos, ItensEstimados, UltimoSync |
            Format-Table -AutoSize | Out-String)

        Set-Content -Path $path -Value $logContent
        Write-Host "`nLog salvo em: $path" -ForegroundColor Green
    } else {
        Write-Host "`nLog nao gerado." -ForegroundColor DarkGray
    }
}

function Verificar-ErrosDoBatch {
    param ($batchId)

    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
    $desktop = [Environment]::GetFolderPath("Desktop")
    $logPath = "$desktop\Falhas_Migracao_$batchId`_$timestamp.csv"

    Write-Host "`nColetando usuarios com status 'Failed' no batch '$batchId'..." -ForegroundColor Yellow
    $failedUsers = Get-MigrationUser -BatchId $batchId | Where-Object { $_.Status -eq "Failed" }

    if (-not $failedUsers) {
        Write-Host "Nenhum usuario com erro encontrado." -ForegroundColor Green
        return
    }

    $logData = foreach ($user in $failedUsers) {
        $stats = Get-MigrationUserStatistics -Identity $user.Identity
        Write-Host "`nUsuario: $($user.Identity)" -ForegroundColor Yellow
        Write-Host "Erro resumido: $($stats.ErrorSummary)" -ForegroundColor Red
        Write-Host "Mensagem de erro: $($stats.Error)" -ForegroundColor DarkRed

        [PSCustomObject]@{
            Usuario         = $user.Identity
            Status          = $stats.Status
            UltimaFalha     = $stats.LastFailureTime
            ErrorSummary    = $stats.ErrorSummary
            ErrorMessage    = $stats.Error
            ErrorHelpUrl    = $stats.ErrorHelpUrl
            ErrorTime       = $stats.ErrorTime
            ErrorCode       = $stats.ErrorCode
            ErrorType       = $stats.ErrorType
            ErrorSide       = $stats.ErrorSide
            ErrorHash       = $stats.ErrorHash
        }
    }

    $logData | Export-Csv -Path $logPath -NoTypeInformation
    Write-Host "`nLog de erros exportado para: $logPath" -ForegroundColor Green
}

function BatchExiste {
    param (
        [string]$nomeBatch,
        [string[]]$listaBatches
    )
    $nomeBatchTrim = $nomeBatch.Trim().ToLower()
    foreach ($b in $listaBatches) {
        if ($b.Trim().ToLower() -eq $nomeBatchTrim) {
            return $true
        }
    }
    return $false
}

# Inicio do script

Write-Host "`nConectando ao Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Conectado com sucesso!" -ForegroundColor Green
} catch {
    Write-Host "Erro ao conectar ao Exchange Online: $_" -ForegroundColor Red
    exit
}

function ListarBatches {
    Write-Host "`nListando batches disponiveis..." -ForegroundColor Cyan
    $global:batches = Get-MigrationBatch | Sort-Object CreationDateTime -Descending

    if (-not $global:batches -or $global:batches.Count -eq 0) {
        Write-Host "Nenhum batch encontrado. Encerrando script." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        exit
    }

    $global:batches | Format-Table Identity, Status, MigrationType, TotalCount, CreationDateTime -AutoSize
}

# Lista batches inicialmente
ListarBatches

# Escolher batch válido
do {
    $batchId = Read-Host "`nDigite exatamente o nome do batch para analisar"
    if (-not (BatchExiste -nomeBatch $batchId -listaBatches $global:batches.Identity)) {
        Write-Host "Batch nao encontrado. Tente novamente." -ForegroundColor Red
        $batchId = $null
    }
} while (-not $batchId)

# Menu principal
do {
    Write-Host "`n========= MENU PRINCIPAL =========" -ForegroundColor Cyan
    Write-Host "Batch atual: $batchId" -ForegroundColor Yellow
    Write-Host "1. Verificar status completo do batch"
    Write-Host "2. Verificar erros detalhados dos usuarios com falha"
    Write-Host "3. Mudar lote"
    Write-Host "4. Sair"
    $opcao = Read-Host "`nEscolha uma opcao (1-4)"

    switch ($opcao) {
        "1" {
            Verificar-StatusDoBatch -batchId $batchId
        }
        "2" {
            Verificar-ErrosDoBatch -batchId $batchId
        }
        "3" {
            ListarBatches
            do {
                $novoBatch = Read-Host "`nDigite exatamente o nome do novo batch para analisar"
                if (-not (BatchExiste -nomeBatch $novoBatch -listaBatches $global:batches.Identity)) {
                    Write-Host "Batch nao encontrado. Tente novamente." -ForegroundColor Red
                    $novoBatch = $null
                }
            } while (-not $novoBatch)
            $batchId = $novoBatch
            Write-Host "`nBatch alterado para: $batchId" -ForegroundColor Green
        }
        "4" {
            Write-Host "`nEncerrando script..." -ForegroundColor Green
            Disconnect-ExchangeOnline -Confirm:$false
        }
        default {
            Write-Host "Opcao invalida. Tente novamente." -ForegroundColor Red
        }
    }
} while ($opcao -ne "4")
