<#
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Ultima atualizacao: 30 de junho de 2025

    Descricao:
    Este script conecta ao Exchange Online, solicita o nome de um batch de migracao,
    identifica os usuarios com falha e exporta os detalhes completos de erro para um arquivo CSV no Desktop.

    Requisitos:
    - PowerShell 5.1 ou superior
    - Modulo ExchangeOnlineManagement instalado (Install-Module ExchangeOnlineManagement)
    - Permissoes apropriadas no EXO para visualizar lotes de migracao
#>

# Solicita o nome do batch ao usuario
$batchId = Read-Host "Informe o nome do batch de migracao a ser analisado"

# Caminho do arquivo de log no Desktop
$timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$desktop = [Environment]::GetFolderPath("Desktop")
$logPath = "$desktop\Falhas_Migracao_$batchId`_$timestamp.csv"

# Conectar ao Exchange Online
Write-Host "`nConectando ao Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Conectado com sucesso!" -ForegroundColor Green
} catch {
    Write-Host "Erro ao conectar ao Exchange Online: $_" -ForegroundColor Red
    exit
}

# Verifica existencia do batch
Write-Host "`nVerificando existencia do batch '$batchId'..." -ForegroundColor Cyan
$batch = Get-MigrationBatch -Identity $batchId -ErrorAction SilentlyContinue
if (-not $batch) {
    Write-Host "Batch '$batchId' nao encontrado. Verifique o nome e tente novamente." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Obter usuarios com falha
Write-Host "Listando usuarios com status 'Failed' no batch '$batchId'..." -ForegroundColor Yellow
$failedUsers = Get-MigrationUser -BatchId $batchId | Where-Object { $_.Status -eq "Failed" }

if ($failedUsers.Count -eq 0) {
    Write-Host "Nenhum usuario com status 'Failed' encontrado no batch." -ForegroundColor Green
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Coleta detalhes de erro
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

# Exporta para CSV
$logData | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8

Write-Host "`nLog detalhado exportado para:" -ForegroundColor Green
Write-Host "$logPath" -ForegroundColor Yellow

# Desconecta
Disconnect-ExchangeOnline -Confirm:$false
