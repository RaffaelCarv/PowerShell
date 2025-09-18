<#
    .TITULO
        Teste de Portas Locais com Menu Interativo

    .AUTOR
        Rafael Carvalho
        GitHub: https://github.com/RaffaelCarv/PowerShell

    .DATA
        Criado: 17/09/2025
        Ultima atualizacao: 18/09/2025

    .DESCRICAO
        Script para verificar se portas especificas estao em uso no Windows.
        Suporta lista de portas (-Ports), range em formato "x-y" (-Range)
        ou ambos em conjunto. Retorna porta, status, PID, processo e caminho.
        Inclui menu interativo com loop e exemplos de preenchimento.

    .REQUISITOS
        - PowerShell 5.1 ou superior
        - Permissoes administrativas para consultar processos

    .USO
        . .\Test-PortLocal.ps1
        # O menu interativo sera exibido automaticamente.
        # Escolha a opcao desejada e siga as instrucoes.

    .EXEMPLOS
        1 - Somente lista de portas (ex.: 80,443)
        2 - Somente range de portas (ex.: 8000-8010)
        3 - Lista + range combinados (ex.: 80,443 e 8000-8010)
        4 - Sair

    .RELEASE NOTES
        1.3 (18/09/2025)
            - Adicionados exemplos visiveis em cada opcao do menu
            - Menu em loop continuo ate a escolha de sair
        1.2 (18/09/2025)
            - Adicionado menu interativo com loop e opcao de sair
        1.1 (18/09/2025)
            - Range no formato "x-y"
            - Uso combinado de lista e range
            - Inclusao de PID, nome do processo e caminho
        1.0 (17/09/2025)
            - Versao inicial com lista de portas e range via operador ".."
#>

function Test-PortLocal {
    param (
        [int[]]$Ports,
        [string]$Range
    )

    $allPorts = @()

    if ($Ports) { $allPorts += $Ports }

    if ($Range) {
        if ($Range -match '^\d+-\d+$') {
            $start, $end = $Range -split '-'
            $start = [int]$start
            $end   = [int]$end
            if ($start -gt $end) { $tmp = $start; $start = $end; $end = $tmp }
            $allPorts += ($start..$end)
        } else {
            Write-Host 'Formato invalido para -Range. Use ex: 8000-8010' -ForegroundColor Red
            return
        }
    }

    if (-not $allPorts) {
        Write-Host 'Nenhuma porta especificada.' -ForegroundColor Red
        return
    }

    $allPorts |
    Sort-Object -Unique |
    ForEach-Object {
        $port = $_
        $conn = Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue

        if ($conn) {
            $conn | Group-Object OwningProcess | ForEach-Object {
                $pid = $_.Name
                try {
                    $proc = Get-Process -Id $pid -ErrorAction Stop
                    [PSCustomObject]@{
                        Porta    = $port
                        Status   = 'Em uso'
                        PID      = $pid
                        Processo = $proc.ProcessName
                        Caminho  = $proc.Path
                    }
                } catch {
                    [PSCustomObject]@{
                        Porta    = $port
                        Status   = 'Em uso'
                        PID      = $pid
                        Processo = 'N/A'
                        Caminho  = 'N/A'
                    }
                }
            }
        } else {
            [PSCustomObject]@{
                Porta    = $port
                Status   = 'Nao listada'
                PID      = ''
                Processo = ''
                Caminho  = ''
            }
        }
    } | Format-Table -AutoSize
}

# ====================
# Menu Interativo com Loop
# ====================
do {
    Write-Host "`nSelecione uma opcao:" -ForegroundColor Cyan
    Write-Host "1 - Somente lista de portas"
    Write-Host "2 - Somente range de portas (x-y)"
    Write-Host "3 - Lista + range combinados"
    Write-Host "4 - Sair"
    $choice = Read-Host "Digite a opcao desejada (1, 2, 3 ou 4)"

    switch ($choice) {
        "1" {
            Write-Host "`nExemplo: 80,443,3389" -ForegroundColor Yellow
            $portsInput = Read-Host "Informe as portas separadas por virgula"
            $ports = $portsInput -split "," | ForEach-Object { [int]$_ }
            Test-PortLocal -Ports $ports
        }
        "2" {
            Write-Host "`nExemplo: 8000-8010" -ForegroundColor Yellow
            $rangeInput = Read-Host "Informe o range de portas no formato x-y"
            Test-PortLocal -Range $rangeInput
        }
        "3" {
            Write-Host "`nExemplo lista: 80,443" -ForegroundColor Yellow
            $portsInput = Read-Host "Informe as portas separadas por virgula"
            Write-Host "`nExemplo range: 8000-8010" -ForegroundColor Yellow
            $rangeInput = Read-Host "Informe o range de portas no formato x-y"
            $ports = $portsInput -split "," | ForEach-Object { [int]$_ }
            Test-PortLocal -Ports $ports -Range $rangeInput
        }
        "4" {
            Write-Host "`nSaindo...`n" -ForegroundColor Yellow
        }
        Default {
            Write-Host "Opcao invalida." -ForegroundColor Red
        }
    }

} while ($choice -ne "4")
