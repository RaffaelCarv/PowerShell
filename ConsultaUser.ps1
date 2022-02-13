$reference = Read-Host "Nome do Usuario"
Get-AdUser -Filter "UserPrincipalName -like '$reference@*'"