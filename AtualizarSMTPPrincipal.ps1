<#
    Criado por: Rafael Carvalho
    Ultima atualizacao: 24 de julho de 2025

    Descricao:
    Script interativo para atualizar SMTP principal de usuarios no Active Directory.
    Permite operar em um usuario especifico, via CSV, em toda a floresta ou em uma OU especifica.
    Inclui modo simulacao (dry run), validacao de duplicacao de SMTP e barra de progresso.

    Requisitos:
    - Modulo ActiveDirectory
    - Permissoes para alterar usuarios no AD

    Estrutura do CSV:
    UserPrincipalName,NewSMTP
    usuario1@dominio.com,novo1@dominio.com
    usuario2@dominio.com,novo2@dominio.com
#>

Import-Module ActiveDirectory

function Create-Log {
    param([string]$content)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
    $desktop = [Environment]::GetFolderPath("Desktop")
    $logFile = "$desktop\AtualizacaoSMTP_$timestamp.txt"
    [System.IO.File]::WriteAllText($logFile, $content, [System.Text.Encoding]::UTF8)
    Write-Host "`n*** Log salvo em: $logFile ***`n" -ForegroundColor Yellow
}

function Atualizar-SMTP {
    param (
        [string]$UserId,
        [string]$NovoSMTP,
        [bool]$Simulacao
    )

    try {
        $user = Get-ADUser -Identity $UserId -Properties ProxyAddresses
        if (-not $user) {
            Write-Host "Usuario $UserId nao encontrado." -ForegroundColor Red
            return "ERRO: Usuario $UserId nao encontrado"
        }

        $proxyAddresses = $user.ProxyAddresses
        $smtpPrincipal = ($proxyAddresses | Where-Object {$_ -cmatch '^SMTP:'})
        $smtpPrincipalClean = $smtpPrincipal -replace '^SMTP:', ''

        # Validacao: ja possui o SMTP informado?
        if ($proxyAddresses -match [regex]::Escape($NovoSMTP)) {
            Write-Host "ERRO: $UserId ja possui o SMTP $NovoSMTP" -ForegroundColor Red
            return "ERRO: $UserId ja possui o SMTP $NovoSMTP"
        }

        # Validacao: outro usuario ja possui esse SMTP?
        $existente = Get-ADUser -Filter {ProxyAddresses -like "*$NovoSMTP*"} -Properties ProxyAddresses
        if ($existente) {
            Write-Host "ERRO: SMTP $NovoSMTP ja esta em uso no usuario $($existente.SamAccountName)" -ForegroundColor Red
            return "ERRO: SMTP $NovoSMTP ja em uso por $($existente.SamAccountName)"
        }

        if ($Simulacao) {
            Write-Host "[SIMULACAO] Usuario: $UserId | Atual: $smtpPrincipalClean | Novo: $NovoSMTP" -ForegroundColor Cyan
            return
        }

        # Remove SMTP atual e adiciona como alias
        $proxyAddresses = $proxyAddresses | Where-Object {$_ -ne $smtpPrincipal}
        $proxyAddresses += "smtp:$smtpPrincipalClean"
        # Adiciona novo SMTP principal
        $proxyAddresses += "SMTP:$NovoSMTP"

        Set-ADUser -Identity $UserId -Replace @{ProxyAddresses = $proxyAddresses}

        Write-Host "Sucesso: $UserId -> Novo SMTP: $NovoSMTP" -ForegroundColor Green
        return "SUCESSO: $UserId -> Novo SMTP: $NovoSMTP"
    } catch {
        Write-Host "Erro ao atualizar $UserId: $_" -ForegroundColor Red
        return "ERRO: $UserId -> $_"
    }
}

function Mostrar-Progresso {
    param (
        [int]$Atual,
        [int]$Total
    )
    $percentual = [math]::Round(($Atual / $Total) * 100, 2)
    $barra = "#" * ($percentual / 5) + "." * (20 - ($percentual / 5))
    Write-Progress -Activity "Atualizando usuarios..." -Status "$percentual% concluido" -PercentComplete $percentual
}

# Menu interativo
Write-Host "Selecione uma opcao:" -ForegroundColor Cyan
Write-Host "1 - Usuario especifico"
Write-Host "2 - Arquivo CSV"
Write-Host "3 - Toda a floresta"
Write-Host "4 - Somente uma OU"
$opcao = Read-Host "Digite o numero da opcao"

# Perguntar se deseja simular ou aplicar
$modo = Read-Host "Deseja simular (S) ou aplicar (A)? [S/A]"
$Simulacao = $modo -eq "S"

$log = @()

switch ($opcao) {
    "1" {
        $usuario = Read-Host "Informe o UPN ou sAMAccountName"
        $novoSMTP = Read-Host "Informe o novo SMTP principal"
        Atualizar-SMTP -UserId $usuario -NovoSMTP $novoSMTP -Simulacao:$Simulacao | Out-Null
        if (-not $Simulacao) {
            $log += "Usuario: $usuario -> Novo SMTP: $novoSMTP"
        }
    }
    "2" {
        $csvPath = Read-Host "Informe o caminho completo do arquivo CSV"
        if (Test-Path $csvPath) {
            $usuarios = Import-Csv -Path $csvPath
            $total = $usuarios.Count
            $i = 0
            foreach ($linha in $usuarios) {
                $i++
                Mostrar-Progresso -Atual $i -Total $total
                Atualizar-SMTP -UserId $linha.UserPrincipalName -NovoSMTP $linha.NewSMTP -Simulacao:$Simulacao | Out-Null
                if (-not $Simulacao) {
                    $log += "Usuario: $($linha.UserPrincipalName) -> Novo SMTP: $($linha.NewSMTP)"
                }
            }
        } else {
            Write-Host "Arquivo CSV nao encontrado." -ForegroundColor Red
            if (-not $Simulacao) { $log += "ERRO: CSV nao encontrado" }
        }
    }
    "3" {
        $novoSMTP = Read-Host "Informe o novo SMTP para todos os usuarios"
        $todosUsuarios = Get-ADUser -Filter * -Properties ProxyAddresses
        $total = $todosUsuarios.Count
        $i = 0
        foreach ($usuario in $todosUsuarios) {
            $i++
            Mostrar-Progresso -Atual $i -Total $total
            Atualizar-SMTP -UserId $usuario.SamAccountName -NovoSMTP $novoSMTP -Simulacao:$Simulacao | Out-Null
            if (-not $Simulacao) {
                $log += "Usuario: $($usuario.SamAccountName) -> Novo SMTP: $novoSMTP"
            }
        }
    }
    "4" {
        $ouDN = Read-Host "Informe o DistinguishedName da OU"
        $novoSMTP = Read-Host "Informe o novo SMTP para os usuarios da OU"
        $usuariosOU = Get-ADUser -SearchBase $ouDN -Filter * -Properties ProxyAddresses
        $total = $usuariosOU.Count
        $i = 0
        foreach ($usuario in $usuariosOU) {
            $i++
            Mostrar-Progresso -Atual $i -Total $total
            Atualizar-SMTP -UserId $usuario.SamAccountName -NovoSMTP $novoSMTP -Simulacao:$Simulacao | Out-Null
            if (-not $Simulacao) {
                $log += "Usuario: $($usuario.SamAccountName) -> Novo SMTP: $novoSMTP"
            }
        }
    }
    Default {
        Write-Host "Opcao invalida." -ForegroundColor Red
    }
}

if (-not $Simulacao -and $log.Count -gt 0) {
    Create-Log -content ($log -join "`r`n")
} elseif ($Simulacao) {
    Write-Host "`n*** SIMULACAO concluida. Nenhuma alteracao foi feita. ***" -ForegroundColor Yellow
}
