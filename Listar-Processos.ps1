<# 
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Ultima atualizacao: 19 de setembro de 2025

    Descricao:
    Este script solicita uma palavra-chave, lista os processos do Windows que correspondem,
    exibe informacoes detalhadas (similar ao Task Manager) e gera um log no Desktop em UTF-8.

    Release Notes:
    1.0 (19/09/2025) - Versao inicial com filtro de processos, exibicao detalhada e exportacao de log.
#>

# Funcao para gerar o log no Desktop
function Create-Log {
    param (
        [string]$logContent
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
    $desktopPath = [System.Environment]::GetFolderPath("Desktop")

    # Nome do log baseado no nome do script
    $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
    if (-not $scriptName) { $scriptName = "Processos" }
    $logFileName = "${scriptName}_$timestamp.txt"
    $logFilePath = Join-Path $desktopPath $logFileName

    [System.IO.File]::WriteAllText($logFilePath, $logContent, [System.Text.Encoding]::UTF8)

    Write-Host "`n*** Log gerado: $logFilePath ***`n" -ForegroundColor Yellow
}

# Solicita palavra-chave
$keyword = Read-Host "Informe a palavra-chave"

# Busca processos com informacoes completas
$processes = Get-Process | Where-Object { $_.ProcessName -ilike "*$keyword*" } | ForEach-Object {
    $cim = Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        Id             = $_.Id
        ServiceName    = $_.ProcessName
        ExecutableName = if ($_.Path) { [System.IO.Path]::GetFileName($_.Path) } else { "$($_.ProcessName).exe" }
        Path           = $_.Path
        CommandLine    = $cim.CommandLine
    }
}

# Exibe resultado no console
$processes | Format-Table -AutoSize

# Gera log (se houver processos encontrados)
if ($processes) {
    $logEntries = $processes | Out-String
    Create-Log -logContent $logEntries
} else {
    Write-Host "Nenhum processo encontrado com a palavra-chave '$keyword'." -ForegroundColor Red
    Create-Log -logContent "Nenhum processo encontrado com a palavra-chave '$keyword'."
}
