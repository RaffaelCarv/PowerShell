<#
    Script para atualizar o e-mail principal de usuários no Active Directory
    Autor: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell/blob/dbebb7c781a3ac38565f94903559eb25b6d633ab/Update-ADUserPrimaryEmail.ps1

    Data de criacao: 30/01/2022
    Ultima atualizacao: 30/07/2024

    Instruções para o arquivo CSV:

    O arquivo CSV deve ter o seguinte formato:

    sAMAccountName,PrimaryEmail
    usuario1,novoemail1@dominio.com
    usuario2,novoemail2@dominio.com

    Onde:
    - "sAMAccountName" deve ser o nome de conta do usuário no Active Directory.
    - "PrimaryEmail" deve ser o novo e-mail principal que você deseja definir para o usuário.

    Certifique-se de que o arquivo CSV tenha o cabeçalho exato e que os dados estejam corretamente formatados.
#>

# Verifica se o módulo ActiveDirectory está instalado e o importa
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "O módulo ActiveDirectory não está instalado. Instalando..."
    Install-WindowsFeature RSAT-AD-PowerShell
}

Import-Module ActiveDirectory

# Função para mostrar uma caixa de diálogo de entrada
function Get-InputBox ($message, $title) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    return [Microsoft.VisualBasic.Interaction]::InputBox($message, $title, "")
}

# Solicita o caminho do arquivo CSV
$csvPath = Get-InputBox "Informe o caminho completo do arquivo incluindo a extensao .csv" "Caminho do Arquivo CSV"

# Verifica se o caminho não está vazio
if (-not [string]::IsNullOrWhiteSpace($csvPath) -and (Test-Path $csvPath)) {
    try {
        # Importa os dados do CSV
        $users = Import-Csv -Path $csvPath -Encoding UTF8

        # Verifica se o CSV tem os cabeçalhos necessários
        if ($users -and $users.PSObject.Properties.Name -contains "sAMAccountName" -and $users.PSObject.Properties.Name -contains "PrimaryEmail") {
            # Define o caminho do arquivo de log no Desktop do usuário com base na data e hora atuais
            $desktopPath = [System.Environment]::GetFolderPath('Desktop')
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $logPath = Join-Path -Path $desktopPath -ChildPath "Update-ADUserPrimaryEmail_$timestamp.log"
            
            # Inicia o arquivo de log
            New-Item -Path $logPath -ItemType File -Force -Encoding UTF8 | Out-Null
            Add-Content -Path $logPath -Value "Log de atualizacao de emails principais - $(Get-Date)" -Encoding UTF8
            Add-Content -Path $logPath -Value "=======================================" -Encoding UTF8

            # Itera sobre cada linha do CSV
            foreach ($user in $users) {
                # Obtém o usuário do AD pelo sAMAccountName
                $adUser = Get-ADUser -Filter {sAMAccountName -eq $user.sAMAccountName}
                
                if ($adUser) {
                    # Define o novo email principal
                    $newPrimaryEmail = $user.PrimaryEmail

                    # Adiciona o novo email ao atributo proxyAddresses
                    Set-ADUser -Identity $adUser -Add @{proxyAddresses="SMTP:$newPrimaryEmail"}

                    # Remove o antigo email principal e adiciona como secundário (opcional)
                    $oldPrimaryEmail = $adUser.EmailAddress
                    if ($oldPrimaryEmail) {
                        Set-ADUser -Identity $adUser -Add @{proxyAddresses="smtp:$oldPrimaryEmail"}
                        Set-ADUser -Identity $adUser -Remove @{proxyAddresses="SMTP:$oldPrimaryEmail"}
                    }

                    # Define o novo email principal no atributo mail
                    Set-ADUser -Identity $adUser -EmailAddress $newPrimaryEmail

                    $logMessage = "Email principal do usuario $($adUser.sAMAccountName) alterado para $newPrimaryEmail"
                    Write-Host $logMessage
                    Add-Content -Path $logPath -Value $logMessage -Encoding UTF8
                } else {
                    $logMessage = "Usuario com sAMAccountName $($user.sAMAccountName) nao encontrado."
                    Write-Host $logMessage
                    Add-Content -Path $logPath -Value $logMessage -Encoding UTF8
                }
            }

            Add-Content -Path $logPath -Value "=======================================" -Encoding UTF8
            Add-Content -Path $logPath -Value "Fim do log - $(Get-Date)" -Encoding UTF8
        } else {
            $errorMessage = "O arquivo CSV nao contem os cabecalhos necessarios: 'sAMAccountName' e 'PrimaryEmail'."
            Write-Host $errorMessage

            # Garante que um arquivo de log seja criado mesmo em caso de erro
            $desktopPath = [System.Environment]::GetFolderPath('Desktop')
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $logPath = Join-Path -Path $desktopPath -ChildPath "Update-ADUserPrimaryEmail_$timestamp.log"

            Add-Content -Path $logPath -Value $errorMessage -Encoding UTF8
            exit
        }
    } catch {
        $errorMessage = "Erro ao importar o arquivo CSV. Verifique o caminho e o formato do arquivo."
        Write-Host $errorMessage

        # Garante que um arquivo de log seja criado mesmo em caso de erro
        $desktopPath = [System.Environment]::GetFolderPath('Desktop')
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $logPath = Join-Path -Path $desktopPath -ChildPath "Update-ADUserPrimaryEmail_$timestamp.log"

        Add-Content -Path $logPath -Value $errorMessage -Encoding UTF8
        exit
    }
} else {
    $errorMessage = "Caminho do arquivo CSV nao fornecido ou arquivo nao encontrado."
    Write-Host $errorMessage

    # Garante que um arquivo de log seja criado mesmo em caso de erro
    $desktopPath = [System.Environment]::GetFolderPath('Desktop')
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logPath = Join-Path -Path $desktopPath -ChildPath "Update-ADUserPrimaryEmail_$timestamp.log"

    Add-Content -Path $logPath -Value $errorMessage -Encoding UTF8
    exit
}
