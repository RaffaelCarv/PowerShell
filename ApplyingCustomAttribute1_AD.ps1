# Importa o módulo Active Directory
Import-Module ActiveDirectory

# Caminho do arquivo CSV
$csvPath = "$home\desktop\Cross_Users_AD.csv"

# Define o nome do atributo personalizado e o valor
$attributeName = "CustomAttribute1"
$attributeValue = "Cross-Tenant-Project"

# Caminho do arquivo de log
$logPath = "$home\desktop\AD_Update_Log.txt"

# Lê o arquivo CSV
$users = Import-Csv -Path $csvPath

# Função para adicionar o atributo personalizado ao esquema do AD
function Add-CustomAttribute {
    param (
        [string]$attributeName
    )

    # Verifica se o atributo já existe
    $existingAttribute = Get-ADObject -Filter { Name -eq $attributeName } -SearchBase (Get-ADRootDSE).SchemaNamingContext -SearchScope Base

    if (-not $existingAttribute) {
        # Comando para adicionar o atributo ao esquema do AD
        Write-Output "Atributo '$attributeName' não encontrado no esquema. Adicionando o atributo..."
        
        # Parâmetros do novo atributo
        $schemaAttrParams = @{
            Name           = $attributeName
            LdapDisplayName = $attributeName
            AttributeSyntax = "2.5.5.12"   # String (Unicode) syntax
            OMObjectClass  = 1.3.12.2.1011.28.1.4
            IsSingleValued = $true
        }
        
        # Adiciona o novo atributo ao esquema
        New-ADObject -Type attributeSchema @schemaAttrParams
        Write-Output "Atributo '$attributeName' adicionado ao esquema."
    } else {
        Write-Output "Atributo '$attributeName' já existe no esquema."
    }
}

# Adiciona o atributo personalizado ao esquema se necessário
Add-CustomAttribute -attributeName $attributeName

# Inicializa variáveis para logs
$successLog = @()
$failureLog = @()
$errorLog = @()

$totalUsers = $users.Count
$currentCount = 0

foreach ($user in $users) {
    # Atualiza o progresso
    $currentCount++
    $percentComplete = [math]::Round(($currentCount / $totalUsers) * 100)
    Write-Progress -Activity "Atualizando atributos dos usuários" -Status "$percentComplete% completo" -PercentComplete ($currentCount / $totalUsers * 100)

    try {
        # Obtém o SamAccountName do usuário a partir do CSV
        $samAccountName = $user.SamAccountName

        # Verifica se o usuário existe no AD
        $adUser = Get-ADUser -Filter {SamAccountName -eq $samAccountName} -Properties $attributeName

        if ($adUser) {
            # Adiciona ou atualiza o atributo personalizado para cada usuário
            Set-ADUser -Identity $adUser -Replace @{$attributeName = $attributeValue}
            $successLog += "Atributo '$attributeName' atualizado para o usuário: $($adUser.SamAccountName)"
        } else {
            $failureLog += "Usuário com SamAccountName '$samAccountName' não encontrado."
        }
    } catch {
        $errorLog += "Erro ao processar usuário com SamAccountName '$samAccountName': $_"
    }
}

# Cria o arquivo de log
$logContent = @(
    "----- Usuários Atualizados com Sucesso -----",
    $successLog,
    "",
    "----- Usuários Não Encontrados -----",
    $failureLog,
    "",
    "----- Erros Durante o Processo -----",
    $errorLog
)

$logContent | Out-File -FilePath $logPath -Encoding UTF8

Write-Output "Processo concluído. Verifique o arquivo de log em: $logPath"
