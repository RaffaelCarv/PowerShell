<#
    .SYNOPSIS
        Script para análise de batches de migração no Exchange Online.
        Lista batches disponíveis, permite escolher um batch e consultar status e erros.

    .DESCRIPTION
        Conecta ao Exchange Online e lista os batches disponíveis.
        O usuário escolhe o batch desejado.
        Exibe um menu para verificar status, erros, concluir lote, mudar de lote ou sair.
        Gera logs no Desktop (TXT para status, CSV para erros).

    .NOTES
        Autor: Rafael Carvalho
        Data de criação: 30/06/2025
        Última atualização: 10/07/2025
        Requisitos:
          - PowerShell 5.1 ou superior
          - Módulo ExchangeOnlineManagement instalado
          - Permissões adequadas para consultar e concluir batches de migração
#>

function Verificar-StatusDoBatch {
    param ($batchId)

    $statusBatch = Get-MigrationBatch -Identity $batchId

    $usuarios = Get-MigrationUser -BatchId $batchId
    $statusUsuarios = $usuarios | ForEach-Object {
        $stats = Get-MigrationUserStatistics -Identity $_.Identity

        $percentual = if ($stats.PercentageComplete -ne $null) { $stats.PercentageComplete } else { 0 }

        [PSCustomObject]@{
            Usuario               = $_.Identity
            Status                = $stats.Status
            Percentual            = $percentual
            BytesTransferidos     = "$($stats.BytesTransferred)"
            TamanhoEstimado       = "$($stats.EstimatedTotalTransferSize)"
            ItensTransferidos     = $stats.SyncedItemCount
            ItensEstimados        = $stats.TotalItemsInSourceMailboxCount
            TaxaTransferencia     = $stats.CurrentBytesTransferredPerMinute
            ConclusaoSyncInicial  = $stats.InitialSeedingCompletedTime
            UltimoSync            = $stats.LastUpdatedTime
        }
    } | Sort-Object Percentual

    Write-Host "`nStatus geral do batch:" -ForegroundColor Cyan
    $statusBatch | Format-List Identity, Status, TotalCount, InitialSyncDuration, CreationDateTime, LastSyncedTime, CompleteAfter

    Write-Host "`nStatus detalhado dos usuarios no batch:" -ForegroundColor Cyan
    $statusUsuarios | Select-Object Usuario, Status,
        @{Name="Percentual %";         Expression = { "$($_.Percentual)%" }},
        @{Name="Bytes/min";            Expression = { $_.TaxaTransferencia }},
        BytesTransferidos, TamanhoEstimado, ItensTransferidos, ItensEstimados,
        @{Name="ConclusaoSyncInicial"; Expression = { $_.ConclusaoSyncInicial }},
        UltimoSync |
        Format-Table -AutoSize

    # Verificar se todos os usuarios estao com falha
    if ($statusUsuarios.Count -gt 0 -and ($statusUsuarios | Where-Object { $_.Status -ne 'Failed' }).Count -eq 0) {
        Write-Host "`nTodos os usuarios estao com falha." -ForegroundColor Red
        $reiniciar = Read-Host "Deseja reiniciar todos agora? (S/N)"
        if ($reiniciar -match '^[sS]$') {
            foreach ($u in $usuarios) {
                try {
                    Resume-MigrationUser -Identity $u.Identity -ErrorAction Stop
                    Write-Host "Usuario $($u.Identity) reiniciado com sucesso." -ForegroundColor Green
                } catch {
                    Write-Host "Erro ao reiniciar o usuario $($u.Identity): $_" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "Nenhum usuario foi reiniciado." -ForegroundColor DarkGray
        }
    }

    $salvar = Read-Host "`nDeseja salvar essas informacoes em um arquivo txt no Desktop? (S/N)"
    if ($salvar -match '^[sS]$') {
        $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
        $desktop   = [Environment]::GetFolderPath("Desktop")
        $path      = "$desktop\Status_Migracao_$batchId`_$timestamp.txt"

        $logContent  = @()
        $logContent += "Status geral do batch '$batchId'"
        $logContent += "-----------------------------------"
        $logContent += ($statusBatch | Format-List Identity, Status, TotalCount, InitialSyncDuration, CreationDateTime, LastSyncedTime, CompleteAfter | Out-String)

        $logContent += "`nStatus detalhado dos usuarios:"
        $logContent += "-------------------------------"
        $logContent += ($statusUsuarios | Select-Object Usuario, Status,
            @{Name="Percentual %";         Expression = { "$($_.Percentual)%" }},
            @{Name="Bytes/min";            Expression = { $_.TaxaTransferencia }},
            BytesTransferidos, TamanhoEstimado, ItensTransferidos, ItensEstimados,
            @{Name="ConclusaoSyncInicial"; Expression = { $_.ConclusaoSyncInicial }},
            UltimoSync |
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

    $salvar = Read-Host "`nDeseja salvar o log de erros em CSV na area de trabalho? (S/N)"
    if ($salvar -match '^[sS]$') {
        $logData | Export-Csv -Path $logPath -NoTypeInformation
        Write-Host "`nLog de erros exportado para: $logPath" -ForegroundColor Green
    } else {
        Write-Host "`nLog nao gerado." -ForegroundColor DarkGray
    }

    $resumir = Read-Host "`nDeseja reiniciar os usuarios com falha? (S/N)"
    if ($resumir -match '^[sS]$') {
        foreach ($user in $failedUsers) {
            try {
                $moveRequest = Get-MoveRequest -Identity $user.Identity -ErrorAction SilentlyContinue
                if ($moveRequest) {
                    Resume-MoveRequest -Identity $user.Identity -ErrorAction Stop
                    Write-Host "Usuario $($user.Identity) reiniciado com sucesso via Resume-MoveRequest." -ForegroundColor Green
                } else {
                    Start-MigrationUser -Identity $user.Identity -ErrorAction Stop
                    Write-Host "Usuario $($user.Identity) iniciado via Start-MigrationUser." -ForegroundColor Green
                }
            }
            catch {
                Write-Host "Erro ao reiniciar o usuario $($user.Identity): $_" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "`nNenhum usuario foi reiniciado." -ForegroundColor DarkGray
    }
}

function Concluir-Lote {
    param ($batchId)

    Write-Host "`nATENCAO: Voce esta prestes a concluir o batch '$batchId'." -ForegroundColor Yellow
    $confirmar = Read-Host "Deseja realmente concluir o batch? (S/N)"
    if ($confirmar -match '^[sS]$') {
        try {
            Complete-MigrationBatch -Identity $batchId -Confirm:$false -ErrorAction Stop
            Write-Host "Batch '$batchId' concluido com sucesso!" -ForegroundColor Green
        } catch {
            Write-Host "Erro ao concluir o batch: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Conclusao cancelada pelo usuario." -ForegroundColor DarkGray
    }
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
    Write-Host ""
    Write-Host "Batch atual: $batchId" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. Verificar status completo do batch"
    Write-Host "2. Verificar erros detalhados dos usuarios com falha"
    Write-Host "3. Mudar lote"
    Write-Host "4. Concluir lote"
    Write-Host "5. Sair"
    $opcao = Read-Host "`nEscolha uma opcao (1-5)"

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
            Concluir-Lote -batchId $batchId
        }
        "5" {
            Write-Host "`nEncerrando script..." -ForegroundColor Green
            Disconnect-ExchangeOnline -Confirm:$false
        }
        default {
            Write-Host "Opcao invalida. Tente novamente." -ForegroundColor Red
        }
    }
} while ($opcao -ne "5")
