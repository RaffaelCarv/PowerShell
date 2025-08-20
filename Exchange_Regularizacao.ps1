<#
    Criado por: Rafael Carvalho
    Ultima atualizacao: 20 de agosto de 2025

    Descricao:
    Script para verificacao e regularizacao basica de servidor Exchange:
    1. Lista discos e memoria
    2. Configura Pagefile conforme recomendacao Microsoft
    3. Configura KeepAlive (TCP/IP) conforme recomendacao Microsoft
    4. Gera relatorio no Desktop
    5. Exibe alerta de necessidade de reinicio apos alteracoes
#>

# Funcao para salvar log no Desktop
function Save-Log {
    param (
        [string]$Content
    )
    $timestamp = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
    $desktop = [Environment]::GetFolderPath("Desktop")
    $file = "$desktop\Exchange_Check_$timestamp.txt"
    [IO.File]::WriteAllText($file, $Content, [Text.Encoding]::UTF8)
    Write-Host "`n*** Relatorio salvo em: $file ***`n" -ForegroundColor Yellow
}

# --- Inventario inicial ---
$log = @()
$log += "===== INVENTARIO DO SERVIDOR ====="
$log += "Data: $(Get-Date)"

# Discos
$log += "`n-- Discos --"
Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
    $sizeGB = [math]::Round($_.Size / 1GB, 0)
    $freeGB = [math]::Round($_.FreeSpace / 1GB, 0)
    $log += "$($_.DeviceID) - Total: $sizeGB GB - Livre: $freeGB GB"
}
# Memoria
$ram = (Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory
$ramGB = [math]::Round($ram / 1GB, 0)
$log += "`n-- Memoria --"
$log += "Total instalada: $ramGB GB"

# --- Funcao Pagefile ---
function Config-Pagefile {
    param([int]$ramGB)

    Write-Host "`n===== Pagefile =====" -ForegroundColor Cyan
    Write-Host "Memoria instalada: $ramGB GB"

    # Exibe configuracao atual do pagefile
    $currentPagefiles = Get-WmiObject Win32_PageFileUsage
    if ($currentPagefiles) {
        Write-Host "`nConfiguracao atual do Pagefile:" -ForegroundColor Yellow
        foreach ($pf in $currentPagefiles) {
            $initialGB = [math]::Round($pf.AllocatedBaseSize / 1024, 2)
            $usedGB = [math]::Round($pf.CurrentUsage / 1024, 2)
            $maxGB = [math]::Round($pf.MaximumSize / 1024, 2)
            Write-Host "Disco: $($pf.Name) | Inicial: $initialGB GB | Maximo: $maxGB GB | Em uso: $usedGB GB"
            $script:log += "Pagefile atual -> Disco: $($pf.Name) | Inicial: $initialGB GB | Maximo: $maxGB GB | Em uso: $usedGB GB"
        }
    } else {
        Write-Host "Nenhum Pagefile configurado no momento." -ForegroundColor Yellow
        $script:log += "Nenhum Pagefile configurado atualmente."
    }

    # Lista discos com total e livre
    Write-Host "`nDiscos disponiveis para configuracao:" -ForegroundColor Cyan
    $disks = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3"
    $i = 1
    foreach ($d in $disks) {
        $size = [math]::Round($d.Size / 1GB, 0)
        $free = [math]::Round($d.FreeSpace / 1GB, 0)
        Write-Host "[$i] $($d.DeviceID) - Total: $size GB | Livre: $free GB"
        $i++
    }

    # Pergunta se deseja aplicar
    $apply = Read-Host "`nDeseja aplicar a configuracao de Pagefile? (S/N)"
    if ($apply -match "^[Ss]$") {

        # Escolha do disco
        $choice = Read-Host "Digite o numero do disco para configurar"
        if ($choice -lt 1 -or $choice -gt $disks.Count) {
            Write-Host "Opcao invalida. Abortando configuracao." -ForegroundColor Red
            $script:log += "Pagefile nao alterado (opcao de disco invalida)."
            return
        }
        $selected = $disks[[int]$choice - 1]

        # Pergunta quantidade desejada
        $maxPage = $ramGB + 1
        $pageSize = Read-Host "Digite o tamanho do Pagefile em GB (min 1 / max $maxPage)"
        $pageSizeInt = [int]$pageSize

        if ($pageSizeInt -lt 1 -or $pageSizeInt -gt $maxPage) {
            Write-Host "Valor invalido. Abortando configuracao." -ForegroundColor Red
            $script:log += "Pagefile nao alterado (valor invalido informado)."
            return
        }

        # Verifica espaço livre corretamente
        $freeGB = [math]::Floor($selected.FreeSpace / 1GB)
        if ($pageSizeInt -gt $freeGB) {
            Write-Host "Espaco livre insuficiente no disco $($selected.DeviceID). Abortando configuracao." -ForegroundColor Red
            $script:log += "Pagefile nao alterado (espaco livre insuficiente no disco)."
            return
        }

        Write-Host "Configurando Pagefile em $($selected.DeviceID) com $pageSizeInt GB..." -ForegroundColor Yellow

        # Remove pagefile antigo e aplica novo
        $pfPath = "$($selected.DeviceID)\pagefile.sys"
        $pfPathEscaped = $pfPath -replace '\\','\\\\'
        $pagefile = Get-WmiObject -Query "Select * from Win32_PageFileSetting Where Name='$pfPathEscaped'"
        if ($pagefile) { $pagefile.Delete() | Out-Null }

        Set-WmiInstance -Class Win32_PageFileSetting -Arguments @{
            Name="$($selected.DeviceID)\pagefile.sys"; 
            InitialSize=$pageSizeInt*1024; 
            MaximumSize=$pageSizeInt*1024
        } | Out-Null

        $logEntry = "Pagefile configurado em $($selected.DeviceID) com $pageSizeInt GB"
        $script:log += $logEntry
        Write-Host "Pagefile configurado com sucesso." -ForegroundColor Green
        $script:NeedReboot = $true
    } else {
        $script:log += "Pagefile nao alterado."
    }
}

# --- Funcao KeepAlive ---
function Config-KeepAlive {
    Write-Host "`n===== KeepAlive =====" -ForegroundColor Cyan
    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"

    $KeepAliveTime = Get-ItemProperty -Path $key -Name "KeepAliveTime" -ErrorAction SilentlyContinue
    $KeepAliveInterval = Get-ItemProperty -Path $key -Name "KeepAliveInterval" -ErrorAction SilentlyContinue

    if ($KeepAliveTime -and $KeepAliveInterval) {
        Write-Host "KeepAlive ja configurado." -ForegroundColor Green
        $script:log += "KeepAlive ja configurado."
    } else {
        Write-Host "KeepAlive nao configurado." -ForegroundColor Red
        $script:log += "KeepAlive nao configurado."
        $apply = Read-Host "Deseja aplicar configuracao recomendada? (S/N)"
        if ($apply -match "^[Ss]$") {
            New-ItemProperty -Path $key -Name "KeepAliveTime" -Value 7200000 -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $key -Name "KeepAliveInterval" -Value 1000 -PropertyType DWord -Force | Out-Null
            Write-Host "KeepAlive configurado (Time=7200000 ms / Interval=1000 ms)" -ForegroundColor Green
            $script:log += "KeepAlive configurado."
            $script:NeedReboot = $true
        } else {
            $script:log += "KeepAlive nao alterado."
        }
    }
}

# --- Execucao ---
Config-Pagefile -ramGB $ramGB
Config-KeepAlive

# Salva relatorio
Save-Log ($log -join "`r`n")

# Alerta de reinicio
if ($script:NeedReboot) {
    Write-Host "`n*** IMPORTANTE: Um reinicio do servidor sera necessario para aplicar as alteracoes. ***`n" -ForegroundColor Yellow
} else {
    Write-Host "`nNenhuma alteracao que exija reinicio foi aplicada.`n" -ForegroundColor Cyan
}
