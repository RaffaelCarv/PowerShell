<#
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Criado em: 15 de março de 2024
    Ultima atualizacao: 15 de junho de 2025

    Descricao:
    Este script realiza uma varredura no Microsoft Entra ID (Cloud Only) para identificar
    usuarios que apresentam erros de provisionamento relacionados ao Cloud Sync (Agente Leve).

    O objetivo principal e gerar um relatorio com dados essenciais para analise de falhas na
    sincronizacao de identidade, como UPNs duplicados, atributos invalidos ou conflitos
    de provisionamento no tenant.

    Requisitos:
    - PowerShell 5.1 ou superior
    - Modulo Microsoft.Graph instalado
    - Permissoes no Microsoft Graph: "User.Read.All", "Directory.Read.All"

    Resultado:
    - Gera um arquivo de log no Desktop do usuario, com codificacao UTF-8
    - Informa usuarios afetados, detalhes dos erros e status da sincronizacao

    Observacoes:
    - Script projetado para ambientes com Microsoft Entra Cloud Sync (nao requer AD Connect)
    - Utiliza o campo OnPremisesProvisioningErrors para validar inconsistencias
    - Pode ser adaptado para exportar tambem em .csv conforme necessidade futura
#>

# ----------------------------
# Declaracao de Variaveis
# ----------------------------
$timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")        # Data e hora para nome de arquivo
$desktop = [Environment]::GetFolderPath("Desktop")             # Caminho para Desktop
$logFile = "$desktop\Erros_CloudSync_$timestamp.txt"          # Caminho completo do log
$log = @()                                                    # Array de log
$scopes = "User.Read.All", "Directory.Read.All"               # Escopos para Graph

# ----------------------------
# Funcao para gerar log em UTF-8
# ----------------------------
function Create-Log {
    param (
        [string[]]$linhas
    )
    [System.IO.File]::WriteAllText($logFile, ($linhas -join "`n"), [System.Text.Encoding]::UTF8)
    Write-Host "`n*** Log gerado: $logFile ***`n" -ForegroundColor Yellow
}

# ----------------------------
# Verifica e instala o modulo Microsoft.Graph, se necessario
# ----------------------------
$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Authentication")
foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Instalando o modulo $mod..." -ForegroundColor Yellow
        Install-Module $mod -Scope AllUsers -Force
    }
}

# Importa os modulos com seguranca, sem registrar todas as funcoes
foreach ($mod in $requiredModules) {
    Import-Module $mod -DisableNameChecking -Force -ErrorAction Stop
}

# ----------------------------
# Autenticacao no Microsoft Graph
# ----------------------------
Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes | Out-Null

# ----------------------------
# Consulta de usuarios com erros de provisionamento
# ----------------------------
Write-Host "Consultando usuarios com erros de provisionamento..." -ForegroundColor Cyan

$usuarios = Get-MgUser -All -Property "DisplayName,UserPrincipalName,Mail,OnPremisesProvisioningErrors,onPremisesSyncEnabled"
$usuariosComErro = $usuarios | Where-Object { $_.OnPremisesProvisioningErrors.Count -gt 0 }

# ----------------------------
# Geracao de relatorio
# ----------------------------
$log += "RELATORIO DE ERROS DE CLOUD SYNC - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$log += ""

if ($usuariosComErro.Count -eq 0) {
    $log += "Nenhum erro de provisionamento encontrado."
} else {
    foreach ($user in $usuariosComErro) {
        $log += "Usuario: $($user.DisplayName)"
        $log += "UPN: $($user.UserPrincipalName)"
        $log += "Email: $($user.Mail)"
        $log += "Sincronizacao habilitada: $($user.OnPremisesSyncEnabled)"
        foreach ($erro in $user.OnPremisesProvisioningErrors) {
            $log += "  - Erro: $($erro.ErrorDetail)"
            $log += "  - Ocorreu em: $($erro.OccurredDateTime)"
        }
        $log += ""
    }
}

# ----------------------------
# Gera o arquivo de log
# ----------------------------
Create-Log -linhas $log
