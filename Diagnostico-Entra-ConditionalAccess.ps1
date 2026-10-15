#requires -Version 5.1
<#
.SYNOPSIS
Diagnostica falhas de autenticacao no Microsoft Entra ID e identifica politicas
de Acesso Condicional que bloquearam o sign-in.

.DESCRIPTION
Conecta ao Microsoft Graph em modo somente leitura e pesquisa os logs de sign-in
usando as informacoes disponiveis na tela de erro, como UPN, App ID, endereco IP,
Correlation ID, Error Code e Timestamp.

Prioriza a busca pelo Correlation ID. Se o evento ainda nao estiver disponivel ou
a consulta nao retornar resultado, utiliza uma janela de tempo ao redor do
Timestamp e aplica os demais filtros localmente.

Exibe os dados do evento, dispositivo, resultado de cada politica de Acesso
Condicional e destaca politicas com resultado failure. Opcionalmente consulta o
estado atual das politicas bloqueadoras. Ao final, permite repetir a mesma busca,
ampliar a janela, iniciar uma nova consulta ou encerrar o script.

O script aceita respostas incompletas do Microsoft Graph. Propriedades opcionais,
como conditionsSatisfied e conditionsNotSatisfied, sao exibidas como Nao informado
quando nao estiverem presentes, sem interromper a execucao.

.REQUISITOS
- Windows PowerShell 5.1.
- Modulo Microsoft.Graph.Authentication.
- Permissoes delegadas AuditLog.Read.All e Policy.Read.All.
- Funcao do Microsoft Entra que permita visualizar logs de entrada e politicas.
- Licenciamento e retencao de logs compativeis com a consulta de sign-ins.
- Conectividade com Microsoft Graph.
- O script utiliza somente requisicoes GET e nao altera objetos do tenant.

.PARAMETER ValidarScript
Executa testes locais com dados simulados para validar o tratamento de propriedades
opcionais e a identificacao de politicas bloqueadoras. Nao conecta ao Microsoft
Graph e nao consulta o tenant.

.EXAMPLE
.\Diagnostico-Entra-ConditionalAccess.ps1

Inicia o diagnostico interativo e solicita somente as informacoes disponiveis.

.EXAMPLE
.\Diagnostico-Entra-ConditionalAccess.ps1 -ValidarScript

Executa a validacao local do script sem conectar ao Microsoft Graph.

.NOTES
Autor: Rafael Carvalho
GitHub: https://github.com/RaffaelCarv/PowerShell

Criacao: 15/10/2026
Ultima atualizacao: 15/10/2026
Versao: 1.2.0

Release Notes:
1.2.0
- Corrigido o erro PropertyNotFoundStrict em propriedades opcionais.
- Adicionada leitura segura para Hashtable, OrderedDictionary e PSCustomObject.
- Adicionado valor Nao informado para propriedades ausentes no Microsoft Graph.
- Corrigida a normalizacao de eventos e politicas retornados pelo Graph.
- Mantida a repeticao da mesma consulta sem redigitar os dados.
- Mantida a ampliacao da janela de pesquisa durante a execucao.
- Adicionado parametro ValidarScript para testes locais sem conexao ao tenant.
- Mantido o modo somente leitura com requisicoes GET.

1.1.0
- Adicionado cabecalho operacional na tela.
- Adicionado fluxo de repeticao e ampliacao da janela de pesquisa.
- Corrigidos conflitos com a variavel automatica PID.
- Removido o uso de labels e goto.

1.0.0
- Versao inicial.
- Adicionada pesquisa por Correlation ID e janela temporal.
- Adicionada identificacao de politicas com resultado failure.
- Adicionada consulta opcional do estado atual da politica bloqueadora.
#>

[CmdletBinding()]
param(
    [switch]$ValidarScript
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-SafeValue {
    param(
        [Parameter(Mandatory = $false)]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )

    if ($null -eq $InputObject) { return $Default }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            if ([string]$key -ieq $Name) { return $InputObject[$key] }
        }
        return $Default
    }

    $property = $InputObject.PSObject.Properties |
        Where-Object { $_.Name -ieq $Name } |
        Select-Object -First 1

    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-SafeText {
    param($InputObject, [string]$Name, [string]$Default = 'Nao informado')
    $value = Get-SafeValue -InputObject $InputObject -Name $Name -Default $null
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return $Default }
    return [string]$value
}

function Get-SafeArray {
    param($InputObject, [string]$Name)
    $value = Get-SafeValue -InputObject $InputObject -Name $Name -Default $null
    if ($null -eq $value) { return @() }
    return @($value)
}

function Read-Answer {
    param([string]$Prompt, [string]$Default = '')
    $suffix = if ($Default) { " [$Default]" } else { '' }
    $answer = Read-Host "$Prompt$suffix"
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    return $answer.Trim()
}

function Write-ScreenHeader {
    Clear-Host
    Write-Host '================================================================================' -ForegroundColor Cyan
    Write-Host ' DIAGNOSTICO MICROSOFT ENTRA ID / CONDITIONAL ACCESS' -ForegroundColor Cyan
    Write-Host '================================================================================' -ForegroundColor Cyan
    Write-Host ' Script      : Diagnostico-Entra-ConditionalAccess.ps1'
    Write-Host ' Autor       : Rafael Carvalho'
    Write-Host ' Versao      : 1.2.0'
    Write-Host ' Objetivo    : Identificar politicas que bloquearam uma autenticacao'
    Write-Host ' Modo        : SOMENTE LEITURA - nenhuma politica sera alterada' -ForegroundColor Green
    Write-Host '================================================================================' -ForegroundColor Cyan
}

function Test-LocalScript {
    Write-Host 'VALIDACAO LOCAL DO SCRIPT' -ForegroundColor Cyan
    Write-Host 'Nenhuma conexao com Microsoft Graph sera realizada.' -ForegroundColor DarkGray

    $samplePolicyMissingFields = [pscustomobject]@{
        displayName = 'Microsoft-managed: Block device code flow'
        id = 'a56ad16c-af51-4fdc-a51a-32075254e5c7'
        result = 'failure'
        enforcedGrantControls = @('Block')
    }

    $sampleEvent = [pscustomobject]@{
        createdDateTime = '2026-10-15T13:24:57Z'
        userPrincipalName = 'usuario@empresa.com'
        appDisplayName = 'Microsoft Exchange REST API Based Powershell'
        appId = 'fb78d390-0c51-40cd-8e17-fdbfab77341b'
        resourceDisplayName = 'Office 365 Exchange Online'
        ipAddress = '100.58.208.88'
        status = [pscustomobject]@{
            errorCode = 53003
            failureReason = 'Access has been blocked by Conditional Access policies.'
        }
        conditionalAccessStatus = 'failure'
        correlationId = '22b9a05d-2e5b-4176-a0a8-ec5e937d52b0'
        id = 'fd51469b-4052-4065-a3d3-2141a4541000'
        deviceDetail = [pscustomobject]@{
            operatingSystem = 'Linux'
            browser = 'Chrome'
            isManaged = $false
            isCompliant = $false
        }
        appliedConditionalAccessPolicies = @($samplePolicyMissingFields)
    }

    try {
        Show-SignInEvent -Event $sampleEvent
        $ids = @(Get-BlockingPolicyIds -Events @($sampleEvent))
        if ($ids.Count -ne 1) { throw 'A politica bloqueadora nao foi identificada.' }
        if ($ids[0] -ne $samplePolicyMissingFields.id) { throw 'O Policy ID identificado esta incorreto.' }
        Write-Host "`nVALIDACAO CONCLUIDA COM SUCESSO." -ForegroundColor Green
        Write-Host 'O tratamento de propriedades opcionais e Policy ID esta funcional.' -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "`nFALHA NA VALIDACAO: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Ensure-GraphConnection {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "`nModulo Microsoft.Graph.Authentication nao encontrado." -ForegroundColor Yellow
        $install = Read-Answer 'Instalar para o usuario atual? (S/N)' 'S'
        if ($install -notmatch '^(s|sim|y|yes)$') { throw 'Modulo necessario nao instalado.' }
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $context = Get-MgContext -ErrorAction SilentlyContinue
    $requiredScopes = @('AuditLog.Read.All', 'Policy.Read.All')
    $mustConnect = ($null -eq $context)

    if (-not $mustConnect) {
        foreach ($scope in $requiredScopes) {
            if ($context.Scopes -notcontains $scope) { $mustConnect = $true }
        }
    }

    if ($mustConnect) {
        Write-Host "`nConectando ao Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        $context = Get-MgContext
    }

    Write-Host "`nConectado como : $($context.Account)" -ForegroundColor Green
    Write-Host "Tenant ID      : $($context.TenantId)" -ForegroundColor Green
}

function ConvertTo-NormalizedObject {
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    return ($InputObject | ConvertTo-Json -Depth 50 | ConvertFrom-Json)
}

function Invoke-GraphGet {
    param([string]$Uri)
    $raw = Invoke-MgGraphRequest -Method GET -Uri $Uri
    return (ConvertTo-NormalizedObject -InputObject $raw)
}

function Read-SearchCriteria {
    Write-Host "`nINFORME SOMENTE OS DADOS DISPONIVEIS" -ForegroundColor Cyan
    Write-Host 'Todos os campos sao opcionais. Quanto mais dados, mais precisa sera a busca.' -ForegroundColor DarkGray
    Write-Host 'Timestamp: copie da tela do erro. Exemplo: 2026-10-15T13:24:59.823Z' -ForegroundColor DarkGray
    Write-Host 'Janela: quantidade de minutos pesquisados antes e depois do Timestamp.' -ForegroundColor DarkGray

    return [pscustomobject]@{
        Upn           = Read-Answer 'UPN do usuario'
        AppId         = Read-Answer 'App ID'
        Ip            = Read-Answer 'Endereco IP'
        CorrelationId = Read-Answer 'Correlation ID'
        RequestId     = Read-Answer 'Request ID (somente referencia)'
        ErrorCode     = Read-Answer 'Error Code (ex.: 53003)'
        Timestamp     = Read-Answer 'Timestamp UTC (vazio = agora)'
        Minutes       = [int](Read-Answer 'Janela em minutos' '10')
    }
}

function Invoke-SignInSearch {
    param($Criteria)

    $events = @()

    if (-not [string]::IsNullOrWhiteSpace($Criteria.CorrelationId)) {
        try {
            $escaped = $Criteria.CorrelationId.Replace("'", "''")
            $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=correlationId eq '$escaped'&`$top=50"
            $response = Invoke-GraphGet -Uri $uri
            $events = @(Get-SafeValue -InputObject $response -Name 'value' -Default @())
        }
        catch {
            Write-Warning 'A pesquisa por Correlation ID falhou. A janela temporal sera utilizada.'
        }
    }

    if ($events.Count -eq 0) {
        if ([string]::IsNullOrWhiteSpace($Criteria.Timestamp)) {
            $center = [DateTimeOffset]::UtcNow
        }
        else {
            try { $center = [DateTimeOffset]::Parse($Criteria.Timestamp).ToUniversalTime() }
            catch { throw 'Timestamp invalido. Exemplo: 2026-10-15T13:24:59.823Z' }
        }

        $start = $center.AddMinutes(-$Criteria.Minutes).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $end = $center.AddMinutes($Criteria.Minutes).ToString('yyyy-MM-ddTHH:mm:ssZ')
        Write-Host "`nPeriodo pesquisado: $start ate $end" -ForegroundColor DarkGray

        $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=createdDateTime ge $start and createdDateTime le $end&`$orderby=createdDateTime desc&`$top=1000"
        $response = Invoke-GraphGet -Uri $uri
        $events = @(Get-SafeValue -InputObject $response -Name 'value' -Default @())
    }

    if ($Criteria.Upn) {
        $events = @($events | Where-Object { (Get-SafeText $_ 'userPrincipalName' '') -eq $Criteria.Upn })
    }
    if ($Criteria.AppId) {
        $events = @($events | Where-Object { (Get-SafeText $_ 'appId' '') -eq $Criteria.AppId })
    }
    if ($Criteria.Ip) {
        $events = @($events | Where-Object { (Get-SafeText $_ 'ipAddress' '') -eq $Criteria.Ip })
    }
    if ($Criteria.ErrorCode) {
        $events = @($events | Where-Object {
            $status = Get-SafeValue $_ 'status'
            (Get-SafeText $status 'errorCode' '') -eq [string]$Criteria.ErrorCode
        })
    }

    return $events
}

function Show-SignInEvent {
    param($Event)

    $status = Get-SafeValue $Event 'status'
    $device = Get-SafeValue $Event 'deviceDetail'
    $policies = @(Get-SafeArray $Event 'appliedConditionalAccessPolicies')

    Write-Host "`n================================================================================" -ForegroundColor Cyan
    Write-Host "DATA           : $(Get-SafeText $Event 'createdDateTime')"
    Write-Host "USUARIO        : $(Get-SafeText $Event 'userPrincipalName')"
    Write-Host "APLICATIVO     : $(Get-SafeText $Event 'appDisplayName')"
    Write-Host "APP ID         : $(Get-SafeText $Event 'appId')"
    Write-Host "RECURSO        : $(Get-SafeText $Event 'resourceDisplayName')"
    Write-Host "IP             : $(Get-SafeText $Event 'ipAddress')"
    Write-Host "ERROR CODE     : $(Get-SafeText $status 'errorCode')"
    Write-Host "MOTIVO         : $(Get-SafeText $status 'failureReason')"
    Write-Host "DETALHE        : $(Get-SafeText $status 'additionalDetails')"
    Write-Host "CA STATUS      : $(Get-SafeText $Event 'conditionalAccessStatus')"
    Write-Host "CORRELATION ID : $(Get-SafeText $Event 'correlationId')"
    Write-Host "SIGN-IN ID     : $(Get-SafeText $Event 'id')"

    Write-Host "`nDISPOSITIVO" -ForegroundColor Yellow
    Write-Host "  Nome         : $(Get-SafeText $device 'displayName')"
    Write-Host "  Device ID    : $(Get-SafeText $device 'deviceId')"
    Write-Host "  Sistema      : $(Get-SafeText $device 'operatingSystem')"
    Write-Host "  Navegador    : $(Get-SafeText $device 'browser')"
    Write-Host "  Gerenciado   : $(Get-SafeText $device 'isManaged')"
    Write-Host "  Compativel   : $(Get-SafeText $device 'isCompliant')"
    Write-Host "  Trust Type   : $(Get-SafeText $device 'trustType')"

    Write-Host "`nPOLITICAS DE CONDITIONAL ACCESS" -ForegroundColor Yellow
    if ($policies.Count -eq 0) {
        Write-Host '  Nenhuma politica retornada ou sem permissao para visualizar detalhes.' -ForegroundColor DarkYellow
        return
    }

    foreach ($policy in $policies) {
        if ($null -eq $policy) { continue }

        $policyResult = Get-SafeText $policy 'result'
        $policyColor = switch ($policyResult) {
            'failure'           { 'Red' }
            'success'           { 'Green' }
            'reportOnlyFailure' { 'DarkYellow' }
            default             { 'Gray' }
        }
        $controls = @(Get-SafeArray $policy 'enforcedGrantControls')

        Write-Host "`n  Politica     : $(Get-SafeText $policy 'displayName')" -ForegroundColor $policyColor
        Write-Host "  Policy ID    : $(Get-SafeText $policy 'id')"
        Write-Host "  Resultado    : $policyResult"
        Write-Host "  Satisfeitas  : $(Get-SafeText $policy 'conditionsSatisfied')"
        Write-Host "  Nao satisf.  : $(Get-SafeText $policy 'conditionsNotSatisfied')"
        Write-Host "  Controles    : $(if ($controls.Count -gt 0) { $controls -join ', ' } else { 'Nao informado' })"
    }
}

function Get-BlockingPolicyIds {
    param([array]$Events)

    $ids = @()
    foreach ($eventItem in $Events) {
        $policies = @(Get-SafeArray $eventItem 'appliedConditionalAccessPolicies')
        foreach ($policy in $policies) {
            if ($null -eq $policy) { continue }
            $result = Get-SafeText $policy 'result' ''
            $policyIdValue = Get-SafeText $policy 'id' ''
            if ($result -eq 'failure' -and -not [string]::IsNullOrWhiteSpace($policyIdValue)) {
                $ids += $policyIdValue
            }
        }
    }
    return @($ids | Sort-Object -Unique)
}

function Show-CurrentPolicyState {
    param([string[]]$PolicyIds)

    if ($null -eq $PolicyIds -or $PolicyIds.Count -eq 0) {
        Write-Host "`nNenhuma politica com resultado FAILURE e Policy ID valido foi encontrada." -ForegroundColor Yellow
        return
    }

    $answer = Read-Answer 'Consultar o estado ATUAL das politicas bloqueadoras? (S/N)' 'S'
    if ($answer -notmatch '^(s|sim|y|yes)$') { return }

    foreach ($policyIdCurrent in $PolicyIds) {
        try {
            $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$policyIdCurrent"
            $policy = Invoke-GraphGet -Uri $uri
            $conditions = Get-SafeValue $policy 'conditions'
            $flows = Get-SafeValue $conditions 'authenticationFlows'
            $grantControls = Get-SafeValue $policy 'grantControls'
            $controls = @(Get-SafeArray $grantControls 'builtInControls')

            Write-Host "`nPOLITICA ATUAL: $(Get-SafeText $policy 'displayName')" -ForegroundColor Cyan
            Write-Host "ID             : $(Get-SafeText $policy 'id')"
            Write-Host "STATE          : $(Get-SafeText $policy 'state')"
            Write-Host "MODIFICADA     : $(Get-SafeText $policy 'modifiedDateTime')"
            Write-Host "TEMPLATE ID    : $(Get-SafeText $policy 'templateId')"
            Write-Host "AUTH FLOW      : $(Get-SafeText $flows 'transferMethods')"
            Write-Host "CONTROLES      : $(if ($controls.Count -gt 0) { $controls -join ', ' } else { 'Nao informado' })"
        }
        catch {
            Write-Warning "Nao foi possivel consultar a politica $policyIdCurrent. $($_.Exception.Message)"
        }
    }
}

if ($ValidarScript) {
    Write-ScreenHeader
    $ok = Test-LocalScript
    if ($ok) { exit 0 } else { exit 1 }
}

Write-ScreenHeader
Ensure-GraphConnection

$continueProgram = $true
while ($continueProgram) {
    $criteria = Read-SearchCriteria
    $searchAgain = $true

    while ($searchAgain) {
        try {
            $events = @(Invoke-SignInSearch -Criteria $criteria)
        }
        catch {
            Write-Host "`nERRO NA CONSULTA: $($_.Exception.Message)" -ForegroundColor Red
            $events = @()
        }

        Write-Host "`nEVENTOS ENCONTRADOS: $($events.Count)" -ForegroundColor Cyan

        if ($events.Count -gt 0) {
            foreach ($eventItem in $events) { Show-SignInEvent -Event $eventItem }
            $blockingPolicyIds = @(Get-BlockingPolicyIds -Events $events)

            Write-Host "`nRESUMO" -ForegroundColor Cyan
            Write-Host "  Eventos localizados                   : $($events.Count)"
            Write-Host "  Politicas bloqueadoras com ID valido : $($blockingPolicyIds.Count)"
            Show-CurrentPolicyState -PolicyIds $blockingPolicyIds
        }
        else {
            Write-Host 'Nenhum evento correspondente foi localizado.' -ForegroundColor Yellow
            Write-Host 'Eventos recentes podem demorar para aparecer nos logs.' -ForegroundColor Yellow
        }

        $defaultChoice = if ($events.Count -eq 0) { 'R' } else { 'N' }
        $choice = Read-Answer '[R] repetir mesmos dados | [A] ampliar janela | [N] nova consulta | [S] sair' $defaultChoice

        switch -Regex ($choice) {
            '^(r|repetir)$' {
                Write-Host "`nRepetindo com os mesmos dados..." -ForegroundColor Cyan
                continue
            }
            '^(a|ampliar)$' {
                $newWindow = [Math]::Max(20, ($criteria.Minutes * 2))
                $criteria.Minutes = [int](Read-Answer 'Nova janela em minutos' ([string]$newWindow))
                Write-Host "`nRepetindo com janela de $($criteria.Minutes) minutos..." -ForegroundColor Cyan
                continue
            }
            '^(n|nova|novo)$' {
                $searchAgain = $false
                continue
            }
            '^(s|sair)$' {
                $searchAgain = $false
                $continueProgram = $false
                continue
            }
            default {
                Write-Host 'Opcao invalida. A consulta sera repetida.' -ForegroundColor Yellow
                continue
            }
        }
    }
}

Write-Host "`nConcluido. O script e somente leitura e nao altera politicas de Conditional Access." -ForegroundColor Green
