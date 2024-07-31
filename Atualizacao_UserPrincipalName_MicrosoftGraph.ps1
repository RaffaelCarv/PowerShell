<# 
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Ultima atualizacao: 30 de julho de 2024

    Descricao:
    Este script verifica a instalacao do modulo Microsoft.Graph, instala-o se necessario,
    importa o modulo, conecta-se ao Microsoft Graph API usando os escopos fornecidos,
    e atualiza o UserPrincipalName dos usuarios especificados em um arquivo CSV.

    Requisitos:
    1. O modulo Microsoft.Graph PowerShell deve estar instalado e atualizado.
    2. A autenticacao deve ser realizada com permissoes adequadas para ler e atualizar usuarios.
    3. O usuario deve ter permissoes de administracao necessárias para a operacao.
    4. O script utiliza o metodo de autenticacao moderna.
    5. As permissoes de escopo "User.Read.All" e "Group.ReadWrite.All" sao necessarias para executar a atualizacao.

    Instrucoes para criar o arquivo CSV:
    1. Crie um arquivo CSV com o nome "usuarios.csv".
    2. O arquivo deve conter as colunas "UserId" e "UserPrincipalName".
    3. Exemplo de estrutura do arquivo CSV:
       UserId,UserPrincipalName
       usuario.exemplo@dominio.com,nome.alternativo@dominio.com
       outro.usuario@dominio.com,nome.alternativo2@dominio.com

    Observacoes:
    - O script solicita ao usuario para selecionar o arquivo CSV atraves de uma caixa de dialogo.
    - O log de execucao é salvo no desktop do usuario, no formato UTF-8, e inclui data e hora.
#>

# Funcao para gerar o log
function Create-Log {
    param (
        [string]$logContent
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
    $desktopPath = [System.Environment]::GetFolderPath("Desktop")
    $logFileName = "Atualizacao_UserPrincipalName_$timestamp.txt"
    $logFilePath = "$desktopPath\$logFileName"

    # Salva o conteudo no arquivo de log com codificacao UTF-8
    [System.IO.File]::WriteAllText($logFilePath, $logContent, [System.Text.Encoding]::UTF8)
    Write-Output "`n`n`n`n`n`n`n`n*** Log gerado: $logFilePath ***"
}

# Verifica se o modulo Microsoft.Graph esta instalado e importa-o uma unica vez
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Output "O modulo Microsoft.Graph nao esta instalado. Instalando agora..."
    Install-Module -Name Microsoft.Graph -Scope AllUsers -Force
}

# Importa o modulo Microsoft.Graph uma unica vez
Import-Module Microsoft.Graph

# Conecta-se ao Microsoft Graph API
Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All" | Out-Null

# Abre uma caixa de dialogo para o usuario selecionar o arquivo CSV
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Arquivos CSV (*.csv)|*.csv"
$openFileDialog.Title = "Selecione o arquivo CSV"
if ($openFileDialog.ShowDialog() -eq "OK") {
    $csvPath = $openFileDialog.FileName
    $logEntries = @()

    # Verifica se o arquivo CSV existe e contem as colunas esperadas
    if (Test-Path $csvPath) {
        $usuarios = Import-Csv -Path $csvPath
        if ($usuarios[0].PSObject.Properties.Name -contains "UserId" -and $usuarios[0].PSObject.Properties.Name -contains "UserPrincipalName") {
            $totalUsuarios = $usuarios.Count
            $contador = 0

            # Atualiza os usuarios
            foreach ($usuario in $usuarios) {
                $contador++
                $percentual = [math]::Round(($contador / $totalUsuarios) * 100, 2)
                $logEntry = "Atualizando UserPrincipalName para o usuario $($usuario.UserId)... $percentual% concluido"
                Write-Output $logEntry
                $logEntries += $logEntry
                try {
                    # Verifica se o UserPrincipalName ja esta em uso
                    $existingUser = Get-MgUser -Filter "userPrincipalName eq '$($usuario.UserPrincipalName)'" -ErrorAction Stop
                    if ($existingUser) {
                        $logEntries += "Erro: UserPrincipalName $($usuario.UserPrincipalName) ja esta em uso."
                        Write-Output "Erro: UserPrincipalName $($usuario.UserPrincipalName) ja esta em uso."
                    } else {
                        Update-MgUser -UserId $usuario.UserId -UserPrincipalName $usuario.UserPrincipalName -ErrorAction Stop
                        $logEntries += "Sucesso: UserPrincipalName atualizado para $($usuario.UserPrincipalName)"
                        Write-Output "Sucesso: UserPrincipalName atualizado para $($usuario.UserPrincipalName)"
                    }
                } catch {
                    $logEntries += "Erro: Falha ao atualizar UserPrincipalName para $($usuario.UserId) - $_"
                    Write-Output "Erro: Falha ao atualizar UserPrincipalName para $($usuario.UserId) - $_"
                }
                Write-Output "`n"
            }
        } else {
            $errorMessage = "O arquivo CSV nao contem as colunas necessarias 'UserId' e 'UserPrincipalName'."
            Write-Output $errorMessage
            $logEntries += $errorMessage
        }
    } else {
        $errorMessage = "O arquivo CSV nao foi encontrado no caminho especificado: $csvPath"
        Write-Output $errorMessage
        $logEntries += $errorMessage
    }

    # Gera o log
    Create-Log -logContent ($logEntries -join "`n`n")
} else {
    $errorMessage = "Nenhum arquivo CSV foi selecionado. Por favor, selecione um arquivo CSV valido para continuar."
    Write-Output $errorMessage
    Create-Log -logContent $errorMessage
}
