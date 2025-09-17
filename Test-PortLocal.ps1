<#
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Ultima atualizacao: 17 de setembro de 2025

    Descricao:
    Script para verificar se portas especificas estao em uso no Windows.
    Permite a verificacao tanto de uma lista definida de portas quanto de um range.

    Requisitos:
    - PowerShell 5.1 ou superior
    - Permissoes administrativas para consulta das conexoes locais

    Exemplo de uso:
    Test-PortLocal -Ports 80,443,3389
    Test-PortLocal -Range 8000..8010
#>

function Test-PortLocal {
    param (
        [int[]]$Ports,
        [object]$Range
    )

    if ($Ports) {
        Write-Host "`nVerificando lista de portas especificas...`n" -ForegroundColor Yellow
        $Ports | ForEach-Object {
            [PSCustomObject]@{
                Porta  = $_
                Status = if (Get-NetTCPConnection -LocalPort $_ -ErrorAction SilentlyContinue) {
                    "Em uso"
                } else {
                    "Nao listada"
                }
            }
        } | Format-Table -AutoSize
    }

    if ($Range) {
        $expandedRange = @()
        if ($Range -is [Array]) {
            $expandedRange = $Range
        } else {
            $expandedRange = Invoke-Expression $Range
        }

        Write-Host "`nVerificando range de portas...`n" -ForegroundColor Yellow
        $expandedRange | ForEach-Object {
            [PSCustomObject]@{
                Porta  = $_
                Status = if (Get-NetTCPConnection -LocalPort $_ -ErrorAction SilentlyContinue) {
                    "Em uso"
                } else {
                    "Nao listada"
                }
            }
        } | Format-Table -AutoSize
    }
}

# Exemplos rapidos:
# Test-PortLocal -Ports 9393,9392,9401,9419
# Test-PortLocal -Range 8000..8010
