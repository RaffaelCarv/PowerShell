# Verificar se os modulos do Exchange Online Management estao instalados
$moduleName = "ExchangeOnlineManagement"
if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    Write-Host "Modulo $moduleName nao encontrado. Instalando..."
    Install-Module -Name $moduleName -Force -AllowClobber
} else {
    Write-Host "Modulo $moduleName ja esta instalado."
}

# Importar o modulo do Exchange Online Management
Import-Module $moduleName

# Conectar ao Exchange Online com MFA
Write-Host "Conectando ao Exchange Online..."
Connect-ExchangeOnline

# Definir o caminho para o arquivo CSV e o arquivo de log no desktop do usuario
$desktopPath = [System.Environment]::GetFolderPath('Desktop')
$csvPath = Join-Path -Path $desktopPath -ChildPath "email_forwarding.csv"
$logPath = Join-Path -Path $desktopPath -ChildPath "log.txt"

# Limpar ou criar o arquivo de log
Clear-Content $logPath -ErrorAction SilentlyContinue
Add-Content $logPath "Inicio do processo: $(Get-Date)"

# Importar o arquivo CSV
$emailMappings = Import-Csv -Path $csvPath

# Numero total de entradas para calculo de progresso
$totalEntries = $emailMappings.Count
$currentEntry = 0

# Loop atraves de cada linha no CSV
foreach ($mapping in $emailMappings) {
    $currentEntry++
    $oldEmail = $mapping.OldEmail
    $newEmail = $mapping.NewEmail

    # Calcular e exibir progresso
    $progress = [math]::Round(($currentEntry / $totalEntries) * 100, 2)
    Write-Progress -Activity "Criando regras de transporte" -Status "$progress% completo" -PercentComplete $progress

    try {
        # Criar a regra de transporte
        New-TransportRule -Name "Inform Non-Monitored - $($oldEmail)" `
                          -Comments "Informar remetentes que o email $($oldEmail) nao e monitorado e encaminhar para $($newEmail)" `
                          -FromScope "NotInOrganization" `
                          -SentTo $oldEmail `
                          -SetAuditSeverity Low `
                          -RejectMessageReasonText "Este email nao e monitorado. Por favor, envie sua mensagem para $($newEmail)" `
                          -RedirectMessageTo $newEmail

        # Log de sucesso
        Add-Content $logPath "[$(Get-Date)] Sucesso: Regra criada para $oldEmail redirecionando para $newEmail."
    } catch {
        # Log de falha
        Add-Content $logPath "[$(Get-Date)] Falha: Nao foi possivel criar a regra para $oldEmail redirecionando para $newEmail. Erro: $_"
    }
}

Add-Content $logPath "Fim do processo: $(Get-Date)"
Write-Host "Processo concluido. Verifique o arquivo de log em $logPath para detalhes."
