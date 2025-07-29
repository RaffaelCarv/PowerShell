<#
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Criacao do script: 05 de agosto de 2024
    Ultima atualizacao: 29 de julho de 2025

    Descricao:
    Script para criar lotes de migracao no Exchange Online com base em uma lista de usuarios.
    Recursos:
     - Seleciona arquivo CSV via Explorer
     - Lista e permite selecionar Migration Endpoint
     - Lista dominios aceitos no tenant para evitar erros
     - Valida CSV, remove duplicados, adiciona dominio se faltar
     - Suporte a modo simulacao
     - Opcao de criar multiplos lotes ou apenas um
     - Exporta CSVs individuais por lote e erros
#>

# ============================
# VARIAVEIS PERSONALIZAVEIS
# ============================
$prefix = "MXOPS"
$tamanhoSufixoAleatorio = 6

# ============================
# FUNCOES AUXILIARES
# ============================

function Get-RandomString {
    param ([int]$length = $tamanhoSufixoAleatorio)
    $characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
    $random = [System.Random]::new()
    -join (1..$length | ForEach-Object { $characters[$random.Next(0, $characters.Length)] })
}

function Export-CSVFile {
    param (
        [Parameter(Mandatory=$true)][array]$data,
        [Parameter(Mandatory=$true)][string]$filenameSuffix
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
    $scriptPath = $PSScriptRoot
    $csvFileName = "MigrationBatch_${timestamp}${filenameSuffix}.csv"
    $csvFilePath = Join-Path -Path $scriptPath -ChildPath $csvFileName
    $data | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    Write-Host "*** CSV exportado: $csvFilePath ***" -ForegroundColor Yellow
}

function Validate-CSVContent {
    param ([string[]]$lines)
    $cleanLines = $lines | Where-Object { $_ -and $_.Trim().Length -gt 0 }
    if ($cleanLines.Count -eq 0) {
        Write-Host "O arquivo CSV esta vazio apos limpeza." -ForegroundColor Red
        return $null
    }
    foreach ($line in $cleanLines) {
        if ($line -match '["]{3,}') {
            Write-Host "Linha contem caracteres invalidos: $line" -ForegroundColor Red
            return $null
        }
    }
    return $cleanLines
}

function Append-DomainIfMissing {
    param (
        [Parameter(Mandatory=$true)][string[]]$emails,
        [Parameter(Mandatory=$true)][string]$domain
    )
    $result = @()
    foreach ($email in $emails) {
        $trimmed = $email.Trim()
        if ($trimmed -notmatch '@') {
            $result += "$trimmed@$domain"
        } else {
            $result += $trimmed
        }
    }
    return $result
}

function Show-Summary {
    param (
        [string]$batchName,
        [string]$domain,
        [int]$totalCSV,
        [int]$alreadyInBatch,
        [int]$includedInBatch,
        [int]$duplicatesInCSV
    )
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    Write-Host "Lote criado: $batchName" -ForegroundColor Cyan
    Write-Host "Dominio: $domain" -ForegroundColor Cyan
    Write-Host "Total no CSV: $totalCSV" -ForegroundColor Cyan
    Write-Host "Ja em lote: $alreadyInBatch" -ForegroundColor Cyan
    Write-Host "Incluidos no lote: $includedInBatch" -ForegroundColor Cyan
    Write-Host "Duplicados internos no CSV: $duplicatesInCSV" -ForegroundColor Cyan
    Write-Host "----------------------------------------`n" -ForegroundColor Cyan
}

function Chunk-List {
    param (
        [Parameter(Mandatory=$true)][array]$list,
        [Parameter(Mandatory=$true)][int]$chunkSize
    )
    $chunks = @()
    for ($i = 0; $i -lt $list.Count; $i += $chunkSize) {
        $chunks += ,($list[$i..([Math]::Min($i + $chunkSize - 1, $list.Count - 1))])
    }
    return $chunks
}

# ============================
# INICIO DO SCRIPT
# ============================

# Verifica modulo ExchangeOnlineManagement
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Modulo ExchangeOnlineManagement nao encontrado. Instalando..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Scope AllUsers -Force
}
Import-Module ExchangeOnlineManagement

# Conectar Exchange Online
Write-Host "Conectando ao Exchange Online..." -ForegroundColor Yellow
Connect-ExchangeOnline -ShowBanner:$false
Write-Host "Conectado ao Exchange Online." -ForegroundColor Green

# Seleciona Migration Endpoint
$endpoints = @(Get-MigrationEndpoint)
if ($endpoints.Count -eq 0) {
    Write-Host "Nenhum Migration Endpoint encontrado. Verifique configuracao." -ForegroundColor Red
    exit
}
Write-Host "`nEndpoints de migracao disponiveis:" -ForegroundColor Cyan
for ($i = 0; $i -lt $endpoints.Count; $i++) {
    Write-Host "[$($i+1)] $($endpoints[$i].Identity)" -ForegroundColor Yellow
}
Write-Host  # Espaço extra

$endpointIndex = Read-Host "Informe o numero do Migration Endpoint desejado"
if (-not ($endpointIndex -match '^\d+$') -or $endpointIndex -lt 1 -or $endpointIndex -gt $endpoints.Count) {
    Write-Host "Indice invalido." -ForegroundColor Red
    exit
}
$nomeMigrationEndpoint = $endpoints[$endpointIndex - 1].Identity
Write-Host "Migration Endpoint selecionado: $nomeMigrationEndpoint" -ForegroundColor Green

# Lista dominios aceitos
$acceptedDomains = @(Get-AcceptedDomain)
Write-Host "`nDominios aceitos no tenant:" -ForegroundColor Cyan
for ($i = 0; $i -lt $acceptedDomains.Count; $i++) {
    Write-Host "[$($i+1)] $($acceptedDomains[$i].DomainName)" -ForegroundColor Yellow
}
Write-Host  # Espaço extra

$domainChoice = Read-Host "Informe o numero do dominio desejado ou digite manualmente"
if ($domainChoice -match '^\d+$' -and $domainChoice -ge 1 -and $domainChoice -le $acceptedDomains.Count) {
    $dominioEntregaDestino = $acceptedDomains[$domainChoice - 1].DomainName
} else {
    $dominioEntregaDestino = $domainChoice
}
Write-Host "Dominio selecionado: $dominioEntregaDestino" -ForegroundColor Green

# Mensagem explicativa ANTES do explorer CSV
Write-Host "`nPor favor, selecione o arquivo CSV com a lista de usuarios para migracao." -ForegroundColor Cyan

# Seleciona arquivo CSV via Explorer
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Arquivos CSV (*.csv)|*.csv"
$openFileDialog.Title = "Selecione o arquivo CSV"
if ($openFileDialog.ShowDialog() -eq "OK") {
    $csvPath = $openFileDialog.FileName
} else {
    Write-Host "Nenhum arquivo CSV selecionado. Saindo." -ForegroundColor Red
    exit
}
Write-Host "Arquivo CSV selecionado: $csvPath" -ForegroundColor Cyan

# Leitura e validacao CSV
$rawLines = Get-Content -Path $csvPath
$linhas = Validate-CSVContent -lines $rawLines
if ($null -eq $linhas) { exit }

# Remove cabecalho se existir
if ($linhas[0] -match 'emailaddress|email|upn') {
    $linhasSemCabecalho = $linhas[1..($linhas.Count - 1)]
} else {
    $linhasSemCabecalho = $linhas
}

# Acrescenta dominio se faltar
$emailsTratados = Append-DomainIfMissing -emails $linhasSemCabecalho -domain $dominioEntregaDestino

# Remove duplicados
$emailsDistinct = $emailsTratados | Select-Object -Unique
$duplicatesCount = $emailsTratados.Count - $emailsDistinct.Count

# Monta lista de objetos
$dadosOriginais = $emailsDistinct | ForEach-Object { [PSCustomObject]@{ EmailAddress = $_.ToLower().Trim() } }

# Consulta usuarios ja em lotes
$usuariosEmLotes = Get-MigrationUser | Select-Object Identity, BatchId
$usuariosEmLotesMap = @{}
foreach ($user in $usuariosEmLotes) {
    $key = $user.Identity.ToLower().Trim()
    $usuariosEmLotesMap[$key] = $user.BatchId
}

$usuariosJaEmLote = @()
$usuariosNaoEmLote = @()
foreach ($usuario in $dadosOriginais) {
    if ($usuariosEmLotesMap.ContainsKey($usuario.EmailAddress)) {
        $batchIdValue = $usuariosEmLotesMap[$usuario.EmailAddress]
        if (-not $batchIdValue) { $batchIdValue = "(BatchId nao informado)" }
        $usuariosJaEmLote += [PSCustomObject]@{
            EmailAddress = $usuario.EmailAddress
            BatchId = $batchIdValue
        }
    } else {
        $usuariosNaoEmLote += $usuario
    }
}

# Resumo inicial
Write-Host "`n===============================================" -ForegroundColor Cyan
Write-Host "Usuarios no CSV: $($dadosOriginais.Count)" -ForegroundColor Cyan
Write-Host "Usuarios ja em lote: $($usuariosJaEmLote.Count)" -ForegroundColor DarkYellow
Write-Host "Usuarios que podem ser usados no novo lote: $($usuariosNaoEmLote.Count)" -ForegroundColor Green
Write-Host "Duplicados internos no CSV: $duplicatesCount" -ForegroundColor Magenta
Write-Host "===============================================`n" -ForegroundColor Cyan

$detalharJa = Read-Host "Deseja exibir detalhes dos usuarios ja em lotes? (S/N)"
if ($detalharJa.ToUpper() -eq 'S') {
    foreach ($u in $usuariosJaEmLote) {
        Write-Host " - $($u.EmailAddress) (Batch: $($u.BatchId))" -ForegroundColor Yellow
    }
    Write-Host
}

$detalharNovos = Read-Host "Deseja exibir detalhes dos usuarios que serao processados? (S/N)"
if ($detalharNovos.ToUpper() -eq 'S') {
    foreach ($u in $usuariosNaoEmLote) {
        Write-Host " + $($u.EmailAddress)" -ForegroundColor Green
    }
    Write-Host
}

if ($usuariosNaoEmLote.Count -eq 0) {
    Write-Host "Todos os usuarios do CSV ja estao em lotes. Nada a processar." -ForegroundColor Red
    exit
}

# Pergunta modo simulacao
$entrada = Read-Host "Executar em modo de simulacao? (S = Sim / N = Nao / C = Cancelar)"
switch ($entrada.ToUpper()) {
    "S" { $modoSimulacao = $true }
    "N" { $modoSimulacao = $false }
    "C" { Write-Host "Execucao cancelada." -ForegroundColor Yellow; exit }
    Default { Write-Host "Entrada invalida. Cancelando." -ForegroundColor Red; exit }
}

if ($modoSimulacao) {
    Write-Host "`n-------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "*** MODO SIMULACAO ATIVO - Nenhum lote sera criado ***" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------------`n" -ForegroundColor Cyan
}

# Pergunta multiplos lotes
$criarMultiplosLotes = Read-Host "Deseja criar multiplos lotes automaticamente? (S/N)"
if ($criarMultiplosLotes.ToUpper() -eq "S") {
    do {
        $limiteLote = Read-Host "Informe o numero maximo de usuarios por lote (exemplo: 264)"
    } while (-not ([int]::TryParse($limiteLote, [ref]$null)) -or $limiteLote -lt 1)

    $chunks = Chunk-List -list $usuariosNaoEmLote -chunkSize $limiteLote
    $usuariosComErro = @()
    $contadorLotes = 0

    foreach ($chunk in $chunks) {
        $contadorLotes++
        $batchName = "$prefix" + "_" + (Get-RandomString)
        Show-Summary -batchName $batchName -domain $dominioEntregaDestino `
            -totalCSV $dadosOriginais.Count -alreadyInBatch $usuariosJaEmLote.Count `
            -includedInBatch $chunk.Count -duplicatesInCSV $duplicatesCount

        if ($modoSimulacao) {
            Write-Host "Simulacao: Lote '$batchName' com $($chunk.Count) usuarios criado (simulado)." -ForegroundColor Cyan
            continue
        }

        try {
            $csvData = $chunk | Select-Object -Property EmailAddress | ConvertTo-Csv -NoTypeInformation | Out-String
            $csvBytes = [System.Text.Encoding]::UTF8.GetBytes($csvData)
            New-MigrationBatch -Name $batchName -SourceEndpoint $nomeMigrationEndpoint -CSVData $csvBytes -Autostart -TargetDeliveryDomain $dominioEntregaDestino
            Write-Host "Sucesso: Lote criado com nome $batchName." -ForegroundColor Green
            Export-CSVFile -data $chunk -filenameSuffix "_$batchName"
        } catch {
            Write-Host "Erro ao criar lote $batchName - $_" -ForegroundColor Red
            $usuariosComErro += $chunk
        }
    }

    if ($usuariosComErro.Count -gt 0) {
        Export-CSVFile -data $usuariosComErro -filenameSuffix "_Erros"
        Write-Host "Usuarios com erro foram salvos em CSV." -ForegroundColor Red
    }
} else {
    # Lote unico
    $batchName = "$prefix" + "_" + (Get-RandomString)
    Show-Summary -batchName $batchName -domain $dominioEntregaDestino `
        -totalCSV $dadosOriginais.Count -alreadyInBatch $usuariosJaEmLote.Count `
        -includedInBatch $usuariosNaoEmLote.Count -duplicatesInCSV $duplicatesCount

    if ($modoSimulacao) {
        Write-Host "Simulacao: Lote '$batchName' com $($usuariosNaoEmLote.Count) usuarios criado (simulado)." -ForegroundColor Cyan
        exit
    }

    try {
        $csvData = $usuariosNaoEmLote | Select-Object -Property EmailAddress | ConvertTo-Csv -NoTypeInformation | Out-String
        $csvBytes = [System.Text.Encoding]::UTF8.GetBytes($csvData)
        New-MigrationBatch -Name $batchName -SourceEndpoint $nomeMigrationEndpoint -CSVData $csvBytes -Autostart -TargetDeliveryDomain $dominioEntregaDestino
        Write-Host "Sucesso: Lote criado com nome $batchName." -ForegroundColor Green
        Export-CSVFile -data $usuariosNaoEmLote -filenameSuffix "_$batchName"
    } catch {
        Write-Host "Erro ao criar lote $batchName - $_" -ForegroundColor Red
        Export-CSVFile -data $usuariosNaoEmLote -filenameSuffix "_Erros"
    }
}
