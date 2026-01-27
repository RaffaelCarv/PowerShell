chcp 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

# =============================================
# FERRAMENTA INTERATIVA DE REPARO DE REDE
# =============================================

If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Host "Execute este script como ADMINISTRADOR." -ForegroundColor Red
    Pause
    Exit
}

function Show-ProgressStep {
    param (
        [string]$Activity,
        [int]$Percent
    )
    Write-Progress -Activity "Reparo de Rede em Andamento" -Status $Activity -PercentComplete $Percent
}

function Pause-Step {
    Write-Host ""
    Read-Host "Pressione ENTER para continuar"
}

do {
    Clear-Host
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "     FERRAMENTA DE LIMPEZA E REPARO DE REDE" -ForegroundColor Cyan
    Write-Host "============================================="
    Write-Host ""
    Write-Host "Escolha uma opcao:`n"
    Write-Host "1  - Limpar cache DNS"
    Write-Host "2  - Renovar IP (DHCP)"
    Write-Host "3  - Reiniciar adaptadores de rede (DESLIGA/LIGA)"
    Write-Host "4  - Limpar ARP e resetar NetBIOS"
    Write-Host "5  - Resetar Winsock (Requer reinicio)"
    Write-Host "6  - Resetar pilha TCP/IP (Requer reinicio)"
    Write-Host "7  - Limpeza COMPLETA guiada (recomendado)"
    Write-Host "8  - Gerar relatorio detalhado de Wi-Fi (WLAN Report)"
    Write-Host "0  - Sair"
    Write-Host ""

    $choice = Read-Host "Digite o numero da opcao desejada"

    switch ($choice) {

        1 {
            Show-ProgressStep "Limpando cache DNS..." 100
            ipconfig /flushdns
            Pause-Step
        }

        2 {
            Show-ProgressStep "Liberando IP atual..." 40
            ipconfig /release
            Show-ProgressStep "Renovando IP..." 80
            ipconfig /renew
            Pause-Step
        }

        3 {
            Write-Host "`nIsso ira reiniciar TODOS os adaptadores de rede." -ForegroundColor Yellow
            $confirm = Read-Host "Deseja continuar? (S/N)"
            if ($confirm -match "^[sS]") {
                Show-ProgressStep "Desativando adaptadores..." 40
                Get-NetAdapter | Disable-NetAdapter -Confirm:$false -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
                Show-ProgressStep "Reativando adaptadores..." 80
                Get-NetAdapter | Enable-NetAdapter -Confirm:$false -ErrorAction SilentlyContinue
            }
            Pause-Step
        }

        4 {
            Show-ProgressStep "Limpando cache ARP..." 40
            arp -d *
            Show-ProgressStep "Resetando NetBIOS..." 80
            nbtstat -R
            Pause-Step
        }

        5 {
            Write-Host "`nATENCAO: Isso resetara o Winsock e pode afetar VPNs." -ForegroundColor Yellow
            $confirm = Read-Host "Deseja continuar? (S/N)"
            if ($confirm -match "^[sS]") {
                netsh winsock reset
                Write-Host "Reinicie o computador apos concluir." -ForegroundColor Green
            }
            Pause-Step
        }

        6 {
            Write-Host "`nATENCAO: Isso resetara a pilha TCP/IP." -ForegroundColor Yellow
            $confirm = Read-Host "Deseja continuar? (S/N)"
            if ($confirm -match "^[sS]") {
                netsh int ip reset
                Write-Host "Reinicie o computador apos concluir." -ForegroundColor Green
            }
            Pause-Step
        }

        7 {
            Write-Host "`nModo GUIADO de limpeza completa iniciado..." -ForegroundColor Cyan

            Show-ProgressStep "Limpando DNS..." 10
            ipconfig /flushdns

            Show-ProgressStep "Renovando IP..." 25
            ipconfig /release
            ipconfig /renew

            Show-ProgressStep "Limpando ARP..." 40
            arp -d *

            Show-ProgressStep "Resetando NetBIOS..." 55
            nbtstat -R

            Show-ProgressStep "Reiniciando adaptadores..." 75
            Get-NetAdapter | Disable-NetAdapter -Confirm:$false -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
            Get-NetAdapter | Enable-NetAdapter -Confirm:$false -ErrorAction SilentlyContinue

            Write-Host "`nDeseja tambem resetar Winsock e TCP/IP? (Requer reinicio)" -ForegroundColor Yellow
            $deep = Read-Host "(S/N)"
            if ($deep -match "^[sS]") {
                Show-ProgressStep "Reset profundo de rede..." 90
                netsh winsock reset
                netsh int ip reset
                Write-Host "Reinicie o computador apos concluir." -ForegroundColor Green
            }

            Show-ProgressStep "Processo finalizado" 100
            Pause-Step
        }

        8 {
            Write-Host "`nGerando relatorio detalhado de Wi-Fi..." -ForegroundColor Cyan
            Show-ProgressStep "Coletando dados de conectividade Wi-Fi..." 60

            netsh wlan show wlanreport | Out-Null
            $reportPath = "$env:ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html"

            Show-ProgressStep "Relatorio gerado" 100

            if (Test-Path $reportPath) {
                Write-Host "`nRelatorio WLAN criado com sucesso." -ForegroundColor Green
                Write-Host "Local do arquivo:" -ForegroundColor Yellow
                Write-Host $reportPath -ForegroundColor White

                $openNow = Read-Host "`nDeseja abrir o relatorio agora? (S/N)"
                if ($openNow -match "^[sS]") {
                    Start-Process $reportPath
                    Write-Host "Abrindo relatorio no navegador..." -ForegroundColor Green
                }
                else {
                    Write-Host "Voce pode abrir o arquivo manualmente depois." -ForegroundColor Cyan
                }
            }
            else {
                Write-Host "Nao foi possivel localizar o relatorio WLAN." -ForegroundColor Red
            }

            Pause-Step
        }

        0 {
            Write-Host "Saindo..."
        }

        Default {
            Write-Host "Opcao invalida."
            Pause-Step
        }
    }

} while ($choice -ne 0)

Write-Progress -Activity "Reparo de Rede" -Completed
