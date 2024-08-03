<# 
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Ultima atualizacao: 03 de agosto de 2024

    Descricao:
    Este script baixa e instala o arquivo AcessoRemotoGratis.exe de forma silenciosa.

    Requisitos:
    1. Permissões de administrador para instalar o software.
    2. Acesso à internet para baixar o arquivo.

    Observacoes:
    - O script não requer interação do usuário e realiza todas as operações de forma automática.
    - O log de execução é salvo no desktop do usuário, no formato UTF-8, e inclui data e hora.
#>

# Função para gerar o log
function Create-Log {
    param (
        [string]$logContent
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
    $desktopPath = [System.Environment]::GetFolderPath("Desktop")
    $logFileName = "Instalacao_AcessoRemotoGratis_$timestamp.txt"
    $logFilePath = "$desktopPath\$logFileName"

    # Salva o conteudo no arquivo de log com codificacao UTF-8
    [System.IO.File]::WriteAllText($logFilePath, $logContent, [System.Text.Encoding]::UTF8)
    Write-Host "`n`n`n*** Log gerado: $logFilePath ***`n`n" -ForegroundColor Yellow
}

# Define o URL do arquivo a ser baixado e o caminho de destino
$url = "https://github.com/acessoremotogratis/acessoremoto/releases/download/1.2.4/AcessoRemotoGratis.exe"
$output = "$env:TEMP\AcessoRemotoGratis.exe"

# Configura o ServicePointManager para usar todos os protocolos
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

# Baixa o arquivo usando System.Net.WebClient
try {
    Write-Host "Baixando o arquivo de $url para $output..." -ForegroundColor Green
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($url, $output)
    Write-Host "Download concluído com sucesso." -ForegroundColor Green
} catch {
    Write-Host "Erro ao baixar o arquivo: $_" -ForegroundColor Red
    Create-Log "Erro ao baixar o arquivo: $_"
    exit 1
}

# Verifica se o arquivo foi baixado com sucesso
if (-not (Test-Path $output)) {
    Write-Host "Arquivo não encontrado: $output" -ForegroundColor Red
    Create-Log "Arquivo não encontrado: $output"
    exit 1
}

# Instala o arquivo de forma silenciosa
try {
    Write-Host "Instalando o arquivo de forma silenciosa..." -ForegroundColor Green
    Start-Process -FilePath $output -ArgumentList "--silent-install" -Wait -NoNewWindow -PassThru | 
    ForEach-Object {
        $process = $_
        $process | Wait-Process
        Write-Host "Código de saída do processo: $($process.ExitCode)" -ForegroundColor Yellow
        if ($process.ExitCode -eq 0) {
            # Adiciona um sleep de 5 segundos antes de mostrar o popup
            Start-Sleep -Seconds 5
            
            # Adiciona um popup indicando sucesso
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show("Instalação concluída com sucesso.", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            Write-Host "Erro na instalação. Código de saída: $($process.ExitCode)" -ForegroundColor Red
            Create-Log "Erro na instalação. Código de saída: $($process.ExitCode)"
        }
    }
} catch {
    Write-Host "Erro na instalação: $_" -ForegroundColor Red
    Create-Log "Erro na instalação: $_"
    exit 1
}
