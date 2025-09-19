<# 
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Criado em: 19/09/2025
    Ultima atualizacao: 19/09/2025

    Descricao:
    Este script solicita uma palavra-chave, lista os processos correspondentes
    (incluindo ServiceName, ExecutableName, Path e CommandLine),
    e pergunta se deseja gerar log somente quando houver resultados.

    Release Notes:
    1.0 (19/09/2025) - Versao inicial com filtro por palavra-chave, exibicao detalhada e opcao de gerar log.
#>

# Solicita palavra-chave
$keyword = Read-Host "Informe a palavra-chave"

# Busca processos com informacoes completas
$processes = Get-CimInstance Win32_Process | Where-Object {
    $_.Name -ilike "*$keyword*" -or
    $_.CommandLine -ilike "*$keyword*"
} | ForEach-Object {
    $service = Get-CimInstance Win32_Service -Filter "ProcessId=$($_.ProcessId)" -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        Id             = $_.ProcessId
        ServiceName    = if ($service) { $service.DisplayName } else { $_.Name }
        ExecutableName = $_.Name
        Path           = $_.ExecutablePath
        CommandLine    = $_.CommandLine
    }
}

if ($processes) {
    $processes | Format-Table -AutoSize

    $resposta = Read-Host "Deseja gerar log? (S/N)"
    if ($resposta -match "^[Ss]$") {
        $timestamp  = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
        $desktop    = [System.Environment]::GetFolderPath("Desktop")
        $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
        if (-not $scriptName) { $scriptName = "Listar-Processos" }
        $logPath    = Join-Path $desktop "$scriptName`_$timestamp.txt"

        $processes | Out-String | Out-File -FilePath $logPath -Encoding utf8
        Write-Host "`n*** Log gerado: $logPath ***`n" -ForegroundColor Yellow
    }
} else {
    Write-Host "Nenhum processo encontrado com a palavra-chave '$keyword'." -ForegroundColor Red
}
