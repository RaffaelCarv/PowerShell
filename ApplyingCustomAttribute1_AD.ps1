# Importa o módulo Active Directory
Import-Module ActiveDirectory

# Caminho do arquivo CSV
$csvPath = "$home\desktop\Cross_Users_AD.csv"

# Define o nome do atributo personalizado e o valor
$attributeName = "CustomAttribute1"
$attributeValue = "Cross-Tenant-Project"

# Lê o arquivo CSV
$users = Import-Csv -Path $csvPath

foreach ($user in $users) {
    # Obtém o SamAccountName do usuário a partir do CSV
    $samAccountName = $user.SamAccountName
    
    # Verifica se o usuário existe no AD
    $adUser = Get-ADUser -Filter {SamAccountName -eq $samAccountName} -Properties $attributeName

    if ($adUser) {
        # Adiciona ou atualiza o atributo personalizado para cada usuário
        Set-ADUser -Identity $adUser -Replace @{$attributeName = $attributeValue}
        Write-Output "Atributo '$attributeName' atualizado para o usuário: $($adUser.SamAccountName)"
    } else {
        Write-Output "Usuário com SamAccountName '$samAccountName' não encontrado."
    }
}

Write-Output "Processo concluído."