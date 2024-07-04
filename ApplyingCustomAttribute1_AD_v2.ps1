# Importa o modulo Active Directory
Import-Module ActiveDirectory

# Caminho do arquivo CSV
$csvPath = "$home\desktop\Cross_Users_AD.csv"

# Define o nome do atributo personalizado e o valor
$attributeName = "ExtensionCustomAttribute1"
$attributeValue = "Cross-Tenant-Project-HMG"

# Obtem a data e hora atual para o nome do arquivo de log
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFileName = "AD_Update_Log_$currentDateTime.txt"
$logPath = "$home\desktop\$logFileName"

# Le o arquivo CSV
$users = Import-Csv -Path $csvPath

# Funcao para verificar se o atributo personalizado já existe no esquema do AD
function Check-CustomAttribute {
    param (
        [string]$attributeName
    )

    # Verifica se o atributo já existe
    $existingAttribute = Get-ADObject -Filter { Name -eq $attributeName } -SearchBase (Get-ADRootDSE).SchemaNamingContext -SearchScope Base

    if ($existingAttribute) {
        Write-Output "Atributo '$attributeName' já existe no esquema."
        return $true
    } else {
        Write-Output "Atributo '$attributeName' não encontrado no esquema."
        return $false
    }
}

# Verifica se o atributo personalizado já existe no esquema
$attributeExists = Check-CustomAttribute -attributeName $attributeName

if (-not $attributeExists) {
    Write-Output "Adicionando o atributo '$attributeName' ao esquema não é necessário, script será encerrado."
    exit
}

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
    Write-Progress -Activity "Atualizando atributos dos usuários" -Status "$percentComplete% completo" -PercentComplete ($currentCount / $totalUsers * 100)

    try {
        # Obtem o SamAccountName do usuário a partir do CSV
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
