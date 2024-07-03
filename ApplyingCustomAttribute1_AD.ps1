# Importa o modulo Active Directory
Import-Module ActiveDirectory

# Caminho do arquivo CSV
$csvPath = "$home\desktop\Cross_Users_AD.csv"

# Define o nome do atributo personalizado e o valor
$attributeName = "CustomAttribute1"
$attributeValue = "Cross-Tenant-Project"

# Obtem a data e hora atual para o nome do arquivo de log
$currentDateTime = Get-Date -Format "ddMMyyyy_HHmmss"
$logFileName = "AD_Update_Log_$currentDateTime.txt"
$logPath = "$home\desktop\$logFileName"

# Le o arquivo CSV
$users = Import-Csv -Path $csvPath

# Funcao para adicionar o atributo personalizado ao esquema do AD
function Add-CustomAttribute {
    param (
        [string]$attributeName
    )

    # Verifica se o atributo ja existe
    $existingAttribute = Get-ADObject -Filter { Name -eq $attributeName } -SearchBase (Get-ADRootDSE).SchemaNamingContext -SearchScope Base

    if (-not $existingAttribute) {
        # Comando para adicionar o atributo ao esquema do AD
        Write-Output "Atributo '$attributeName' nao encontrado no esquema. Adicionando o atributo..."
        
        # Parametros do novo atributo
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
        Write-Output "Atributo '$attributeName' ja existe no esquema."
    }
}

# Adiciona o atributo personalizado ao esquema se necessario
Add-CustomAttribute -attributeName $attributeName

# Inicializa variaveis para logs
$successLog = @()
$failureLog = @()
$errorLog = @()

$totalUsers = $users.Count
$currentCount = 0

foreach ($user in $users) {
    # Atualiza o progresso
    $currentCount++
    $percentComplete = [math]::Round(($currentCount / $totalUsers) * 100)
    Write-Progress -Activity "Atualizando atributos dos usuarios" -Status "$percentComplete% completo" -PercentComplete ($currentCount / $totalUsers * 100)

    try {
        # Obtem o SamAccountName do usuario a partir do CSV
        $samAccountName = $user.SamAccountName

        # Verifica se o usuario existe no AD
        $adUser = Get-ADUser -Filter {SamAccountName -eq $samAccountName} -Properties $attributeName

        if ($adUser) {
            # Adiciona ou atualiza o atributo personalizado para cada usuario
            Set-ADUser -Identity $adUser -Replace @{$attributeName = $attributeValue}
            $successLog += "Atributo '$attributeName' atualizado para o usuario: $($adUser.SamAccountName)"
        } else {
            $failureLog += "Usuario com SamAccountName '$samAccountName' nao encontrado."
        }
    } catch {
        $errorLog += "Erro ao processar usuario com SamAccountName '$samAccountName': $_"
    }
}

# Cria o arquivo de log
$logContent = @(
    "----- Usuarios Atualizados com Sucesso -----",
    $successLog,
    "",
    "----- Usuarios Nao Encontrados -----",
    $failureLog,
    "",
    "----- Erros Durante o Processo -----",
    $errorLog
)

$logContent | Out-File -FilePath $logPath -Encoding UTF8

Write-Output "Processo concluido. Verifique o arquivo de log em: $logPath"
