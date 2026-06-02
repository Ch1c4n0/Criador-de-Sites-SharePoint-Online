<#
.SYNOPSIS
    Cria sites no SharePoint Online em lote (via CSV) ou de forma individual,
    definindo owners, membros e a privacidade (Publico/Privado).

.DESCRIPTION
    Usa o modulo PnP.PowerShell (requer PowerShell 7+). Suporta dois tipos de site:
      - TeamSite          : Site de Equipe conectado a um Grupo do Microsoft 365
                            (suporta Publico/Privado, owners e membros do grupo).
      - CommunicationSite : Site de Comunicacao (sem grupo M365; owners/membros
                            sao adicionados aos grupos associados do SharePoint).

    Recursos:
      - Retry com backoff exponencial para throttling (HTTP 429).
      - Espera inteligente pelo provisionamento do Grupo M365.
      - Suporte a Sensitivity Label e Classificacao.
      - Validacao previa do CSV (campos, duplicados, e-mails).
      - Log em arquivo de cada execucao.

    Autenticacao:
      - Interativo : abre o login da Microsoft no navegador (suporta MFA). Requer -ClientId.
      - AppOnly    : autenticacao sem interacao via ClientId + Certificado (ideal p/ agendamento).

.PARAMETER AdminUrl
    URL do centro de administracao do SharePoint. Ex.: https://contoso-admin.sharepoint.com

.PARAMETER CsvPath
    Caminho para um CSV. Se informado, executa direto em modo lote (sem menu).

.PARAMETER AuthMode
    Interactive (padrao) ou AppOnly.

.PARAMETER ClientId
    Client ID do app registrado no Entra ID. Obrigatorio no login interativo das
    versoes recentes do PnP.PowerShell.

.PARAMETER Tenant
    (AppOnly) Dominio do tenant. Ex.: contoso.onmicrosoft.com

.PARAMETER Thumbprint
    (AppOnly) Thumbprint do certificado instalado na maquina.

.PARAMETER CertificatePath
    (AppOnly) Alternativa ao Thumbprint: caminho para um arquivo .pfx.

.PARAMETER CertificatePassword
    (AppOnly) Senha do .pfx, como SecureString.

.PARAMETER DelaySeconds
    Pausa (segundos) entre a criacao de cada site no lote, para reduzir throttling. Padrao: 5.

.PARAMETER MaxRetries
    Numero maximo de tentativas em caso de throttling. Padrao: 5.

.PARAMETER GroupWaitSeconds
    Tempo maximo (segundos) de espera pelo provisionamento do Grupo M365. Padrao: 120.

.PARAMETER SensitivityLabel
    (Opcional) Id ou nome de um rotulo de confidencialidade aplicado a todos os sites.

.PARAMETER LogPath
    (Opcional) Pasta para os logs. Padrao: <pasta do script>\logs.

.EXAMPLE
    .\Criar-Sites.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "GUID"
    Abre o menu interativo (login no navegador).

.EXAMPLE
    .\Criar-Sites.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "GUID" -CsvPath ".\modelo-sites.csv"
    Cria em lote a partir do CSV.

.EXAMPLE
    .\Criar-Sites.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -CsvPath ".\modelo-sites.csv" `
        -AuthMode AppOnly -ClientId "GUID" -Tenant "contoso.onmicrosoft.com" -Thumbprint "ABC123..."
    Cria em lote usando autenticacao App-Only (sem interacao).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'AppOnly')]
    [string]$AuthMode = 'Interactive',

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $false)]
    [string]$CertificatePath,

    [Parameter(Mandatory = $false)]
    [securestring]$CertificatePassword,

    [Parameter(Mandatory = $false)]
    [int]$DelaySeconds = 5,

    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,

    [Parameter(Mandatory = $false)]
    [int]$GroupWaitSeconds = 120,

    [Parameter(Mandatory = $false)]
    [string]$SensitivityLabel,

    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------
# Log
# ----------------------------------------------------------------------------

$script:LogFile = $null

function Initialize-Log {
    param([string]$Dir)
    if ([string]::IsNullOrWhiteSpace($Dir)) { $Dir = Join-Path $PSScriptRoot 'logs' }
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile = Join-Path $Dir "Criar-Sites_$stamp.log"
    "==== Criar-Sites - inicio em $(Get-Date -Format 's') ====" | Out-File -FilePath $script:LogFile -Encoding utf8
}

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    if ($script:LogFile) {
        "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Out-File -FilePath $script:LogFile -Append -Encoding utf8
    }
}

# ----------------------------------------------------------------------------
# Funcoes auxiliares de saida
# ----------------------------------------------------------------------------

function Write-Step  { param([string]$Msg) Write-Host "==> $Msg" -ForegroundColor Cyan;   Write-Log $Msg 'STEP' }
function Write-Ok    { param([string]$Msg) Write-Host "  [OK] $Msg" -ForegroundColor Green; Write-Log $Msg 'OK' }
function Write-Warn2 { param([string]$Msg) Write-Host "  [!]  $Msg" -ForegroundColor Yellow; Write-Log $Msg 'WARN' }
function Write-Err2  { param([string]$Msg) Write-Host "  [ERRO] $Msg" -ForegroundColor Red; Write-Log $Msg 'ERRO' }

# ----------------------------------------------------------------------------
# Retry com backoff exponencial (para throttling / HTTP 429)
# ----------------------------------------------------------------------------

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock]$Script,
        [string]$What = 'operacao',
        [int]$Retries = $MaxRetries
    )
    $attempt = 0
    while ($true) {
        try {
            return & $Script
        }
        catch {
            $attempt++
            $msg = $_.Exception.Message
            $status = $null
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }

            $isTransient = ($status -eq 429) -or ($status -eq 503) -or
                           ($msg -match '(?i)throttl|too many requests|429|503|temporarily')

            if ($attempt -ge $Retries -or -not $isTransient) { throw }

            $wait = [math]::Min(60, [math]::Pow(2, $attempt))  # 2,4,8,16,32,60...
            Write-Warn2 "Throttling/erro transitorio em '$What' (tentativa $attempt/$Retries). Aguardando $wait s..."
            Start-Sleep -Seconds $wait
        }
    }
}

# ----------------------------------------------------------------------------
# Ambiente / conexao
# ----------------------------------------------------------------------------

function Assert-PowerShell7 {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw "Este script requer PowerShell 7+ (voce esta no $($PSVersionTable.PSVersion)). Abra o 'pwsh' e rode novamente."
    }
}

function Ensure-PnPModule {
    if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
        Write-Step "Modulo PnP.PowerShell nao encontrado. Preparando instalacao..."

        # Garante TLS 1.2 (relevante em hosts antigos) e o provedor NuGet, sem prompt.
        try {
            [Net.ServicePointManager]::SecurityProtocol = `
                [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        } catch { }
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
        Write-Step "Instalando PnP.PowerShell para o usuario atual..."
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-Ok "Modulo PnP.PowerShell carregado."
}

function Connect-Tenant {
    param([string]$Url)

    Write-Step "Conectando em $Url (modo: $AuthMode)..."
    Invoke-WithRetry -What "conexao em $Url" -Script {
        switch ($AuthMode) {
            'Interactive' {
                if ($ClientId) {
                    Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
                }
                else {
                    # Sem ClientId o login interativo falha nas versoes recentes do PnP.
                    Connect-PnPOnline -Url $Url -Interactive
                }
            }
            'AppOnly' {
                if (-not $ClientId -or -not $Tenant) { throw "AppOnly requer -ClientId e -Tenant." }
                if ($Thumbprint) {
                    Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint
                }
                elseif ($CertificatePath) {
                    Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $Tenant `
                        -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword
                }
                else { throw "AppOnly requer -Thumbprint OU -CertificatePath." }
            }
        }
    }
    Write-Ok "Conectado."
}

# Converte a URL de admin (https://contoso-admin.sharepoint.com) na URL raiz
# do tenant (https://contoso.sharepoint.com).
function Get-TenantRootUrl {
    param([string]$Admin)
    return ($Admin -replace '-admin\.sharepoint\.com', '.sharepoint.com').TrimEnd('/')
}

# Divide uma string de e-mails separados por ; ou , em um array limpo.
function Split-Emails {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    return $Value -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
}

# ----------------------------------------------------------------------------
# Validacao previa do CSV
# ----------------------------------------------------------------------------

function Test-CsvRows {
    param([object[]]$Rows)

    $errors = @()
    $seenAlias = @{}
    $seenUrl   = @{}
    $emailRx   = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    $i = 0

    foreach ($r in $Rows) {
        $i++
        $tipo = ($r.Tipo | ForEach-Object { $_ }) -as [string]
        if ($tipo) { $tipo = $tipo.Trim() }

        if ([string]::IsNullOrWhiteSpace($r.Titulo)) { $errors += "Linha ${i}: 'Titulo' vazio." }

        if ($tipo -notin @('Team', 'Communication')) {
            $errors += "Linha ${i}: 'Tipo' invalido ('$($r.Tipo)'). Use Team ou Communication."
        }

        if ($tipo -eq 'Team') {
            if ([string]::IsNullOrWhiteSpace($r.Alias)) {
                $errors += "Linha ${i}: 'Alias' e obrigatorio para Team Site."
            }
            else {
                $a = $r.Alias.Trim()
                if ($a -match '\s') { $errors += "Linha ${i}: 'Alias' nao pode conter espacos ('$a')." }
                $key = $a.ToLower()
                if ($seenAlias.ContainsKey($key)) { $errors += "Linha ${i}: 'Alias' duplicado no CSV ('$a') - tambem na linha $($seenAlias[$key])." }
                else { $seenAlias[$key] = $i }
            }
            if ($r.Privacidade -and ($r.Privacidade.Trim() -notin @('Public', 'Private'))) {
                $errors += "Linha ${i}: 'Privacidade' invalida ('$($r.Privacidade)'). Use Public ou Private."
            }
        }

        if ($tipo -eq 'Communication') {
            if ([string]::IsNullOrWhiteSpace($r.Url)) {
                $errors += "Linha ${i}: 'Url' e obrigatoria para Communication Site."
            }
            else {
                $key = $r.Url.Trim().ToLower()
                if ($seenUrl.ContainsKey($key)) { $errors += "Linha ${i}: 'Url' duplicada no CSV ('$($r.Url)') - tambem na linha $($seenUrl[$key])." }
                else { $seenUrl[$key] = $i }
            }
        }

        foreach ($e in (@(Split-Emails $r.Owners) + @(Split-Emails $r.Members))) {
            if ($e -notmatch $emailRx) { $errors += "Linha ${i}: e-mail invalido ('$e')." }
        }
    }
    return $errors
}

# ----------------------------------------------------------------------------
# Criacao de site (nucleo, usado tanto no modo unico quanto em lote)
# ----------------------------------------------------------------------------

function New-SharePointSite {
    param(
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [ValidateSet('Team', 'Communication')] [string]$Type,
        [string]$Alias,          # usado por Team Site (mailNickname do grupo)
        [string]$Url,            # usado por Communication Site (parte apos /sites/ ou URL completa)
        [ValidateSet('Public', 'Private')] [string]$Privacy = 'Private',
        [string[]]$Owners = @(),
        [string[]]$Members = @(),
        [string]$Description = '',
        [string]$Label = '',     # sensitivity label (id ou nome)
        [string]$Classification = ''
    )

    $tenantRoot = Get-TenantRootUrl -Admin $AdminUrl
    if ([string]::IsNullOrWhiteSpace($Label)) { $Label = $SensitivityLabel }

    if ($Type -eq 'Team') {
        if ([string]::IsNullOrWhiteSpace($Alias)) { throw "TeamSite '$Title' requer um Alias." }

        Write-Step "Criando Team Site '$Title' (alias: $Alias, $Privacy)..."
        $params = @{
            Type     = 'TeamSite'
            Title    = $Title
            Alias    = $Alias
            IsPublic = ($Privacy -eq 'Public')
        }
        if ($Owners.Count -gt 0)                              { $params['Owners'] = $Owners }
        if (-not [string]::IsNullOrWhiteSpace($Description))  { $params['Description'] = $Description }
        if (-not [string]::IsNullOrWhiteSpace($Label))        { $params['SensitivityLabel'] = $Label }
        if (-not [string]::IsNullOrWhiteSpace($Classification)) { $params['Classification'] = $Classification }

        $siteUrl = Invoke-WithRetry -What "criacao do Team Site '$Title'" -Script { New-PnPSite @params }
        Write-Ok "Site criado: $siteUrl"

        # O grupo M365 e provisionado de forma assincrona; aguardamos ate GroupWaitSeconds.
        $maxTries = [math]::Max(1, [int][math]::Ceiling($GroupWaitSeconds / 5))
        Write-Step "Aguardando provisionamento do Grupo M365 (ate $GroupWaitSeconds s)..."
        $group = $null
        for ($i = 0; $i -lt $maxTries -and -not $group; $i++) {
            Start-Sleep -Seconds 5
            try { $group = Get-PnPMicrosoft365Group -Identity $Alias } catch { $group = $null }
        }

        if ($group) {
            if ($Members.Count -gt 0) {
                Write-Step "Adicionando $($Members.Count) membro(s)..."
                foreach ($m in $Members) {
                    try {
                        Invoke-WithRetry -What "add membro '$m'" -Script {
                            Add-PnPMicrosoft365GroupMember -Identity $group.Id -Users $m -RemoveExisting:$false
                        }
                        Write-Ok "Membro: $m"
                    } catch { Write-Warn2 "Falha ao adicionar membro '$m': $($_.Exception.Message)" }
                }
            }
            # Garante owners (alem dos passados no New-PnPSite).
            foreach ($o in $Owners) {
                try {
                    Invoke-WithRetry -What "add owner '$o'" -Script {
                        Add-PnPMicrosoft365GroupOwner -Identity $group.Id -Users $o -RemoveExisting:$false
                    }
                } catch { Write-Warn2 "Falha ao garantir owner '$o': $($_.Exception.Message)" }
            }
        }
        else {
            Write-Warn2 "Grupo M365 nao localizado em $GroupWaitSeconds s. Adicione owners/membros manualmente ou rode novamente."
        }

        return $siteUrl
    }
    else {
        # Communication Site
        if ([string]::IsNullOrWhiteSpace($Url)) { throw "CommunicationSite '$Title' requer uma Url." }

        # Aceita tanto a URL completa quanto apenas o sufixo (ex.: "marketing").
        if ($Url -notmatch '^https?://') { $fullUrl = "$tenantRoot/sites/$($Url.TrimStart('/'))" }
        else { $fullUrl = $Url }

        Write-Step "Criando Communication Site '$Title' ($fullUrl)..."
        $params = @{
            Type  = 'CommunicationSite'
            Title = $Title
            Url   = $fullUrl
        }
        if (-not [string]::IsNullOrWhiteSpace($Description))    { $params['Description'] = $Description }
        if (-not [string]::IsNullOrWhiteSpace($Label))          { $params['SensitivityLabel'] = $Label }
        if (-not [string]::IsNullOrWhiteSpace($Classification)) { $params['Classification'] = $Classification }

        $siteUrl = Invoke-WithRetry -What "criacao do Communication Site '$Title'" -Script { New-PnPSite @params }
        Write-Ok "Site criado: $siteUrl"

        # Communication Site nao tem grupo M365. Adicionamos owners/membros aos
        # grupos associados do SharePoint e owners tambem como admin do site.
        Write-Step "Configurando permissoes do Communication Site..."
        Connect-Tenant -Url $siteUrl

        foreach ($o in $Owners) {
            try {
                Invoke-WithRetry -What "owner '$o'" -Script { Set-PnPSiteCollectionAdmin -Owners $o }
                Write-Ok "Owner (admin do site): $o"
            } catch { Write-Warn2 "Falha ao definir owner '$o': $($_.Exception.Message)" }
        }

        if ($Members.Count -gt 0) {
            try {
                $memberGroup = Get-PnPGroup -AssociatedMemberGroup
                foreach ($m in $Members) {
                    try {
                        Invoke-WithRetry -What "membro '$m'" -Script { Add-PnPGroupMember -LoginName $m -Group $memberGroup }
                        Write-Ok "Membro: $m"
                    } catch { Write-Warn2 "Falha ao adicionar membro '$m': $($_.Exception.Message)" }
                }
            } catch { Write-Warn2 "Nao foi possivel obter o grupo de membros: $($_.Exception.Message)" }
        }

        # Volta a conexao para o admin para continuar o lote.
        Connect-Tenant -Url $AdminUrl
        return $siteUrl
    }
}

# ----------------------------------------------------------------------------
# Modo LOTE (CSV)
# ----------------------------------------------------------------------------

function Invoke-BatchFromCsv {
    param([string]$Path)

    if (-not (Test-Path $Path)) { throw "CSV nao encontrado: $Path" }

    Write-Step "Lendo CSV: $Path"
    $rows = Import-Csv -Path $Path -Delimiter ';'
    if (-not $rows -or @($rows).Count -eq 0 -or -not ($rows | Get-Member -Name 'Titulo' -ErrorAction SilentlyContinue)) {
        # Tenta com virgula caso nao tenha ; ou nao tenha as colunas esperadas.
        $rows = Import-Csv -Path $Path
    }
    $rows = @($rows)
    Write-Ok "$($rows.Count) linha(s) encontrada(s)."

    # Validacao previa - aborta antes de criar qualquer coisa se houver erros.
    Write-Step "Validando o CSV..."
    $problems = Test-CsvRows -Rows $rows
    if ($problems.Count -gt 0) {
        Write-Err2 "Foram encontrados $($problems.Count) problema(s) no CSV:"
        $problems | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red; Write-Log $_ 'CSV' }
        throw "Corrija o CSV e tente novamente. Nenhum site foi criado."
    }
    Write-Ok "CSV valido."

    $resultados = @()
    $idx = 0
    foreach ($row in $rows) {
        $idx++
        Write-Host ""
        Write-Host "----- Linha $idx de $($rows.Count) -----" -ForegroundColor Magenta
        try {
            $url = New-SharePointSite `
                -Title          $row.Titulo `
                -Type           $row.Tipo.Trim() `
                -Alias          $row.Alias `
                -Url            $row.Url `
                -Privacy        ($(if ([string]::IsNullOrWhiteSpace($row.Privacidade)) { 'Private' } else { $row.Privacidade.Trim() })) `
                -Owners         (Split-Emails $row.Owners) `
                -Members        (Split-Emails $row.Members) `
                -Description    $row.Descricao `
                -Label          $row.SensitivityLabel `
                -Classification $row.Classificacao

            $resultados += [pscustomobject]@{ Titulo = $row.Titulo; Url = $url; Status = 'Criado' }
        }
        catch {
            Write-Err2 "Falha na linha $idx ('$($row.Titulo)'): $($_.Exception.Message)"
            $resultados += [pscustomobject]@{ Titulo = $row.Titulo; Url = ''; Status = "ERRO: $($_.Exception.Message)" }
        }

        # Pausa entre sites para reduzir throttling (exceto apos o ultimo).
        if ($idx -lt $rows.Count -and $DelaySeconds -gt 0) {
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    Write-Host ""
    Write-Step "Resumo do lote:"
    $resultados | Format-Table -AutoSize

    # Salva o resumo ao lado do log.
    if ($script:LogFile) {
        $csvOut = [System.IO.Path]::ChangeExtension($script:LogFile, 'resultado.csv')
        $resultados | Export-Csv -Path $csvOut -NoTypeInformation -Encoding utf8 -Delimiter ';'
        Write-Ok "Resumo salvo em: $csvOut"
    }
}

# ----------------------------------------------------------------------------
# Modo UNICO (interativo)
# ----------------------------------------------------------------------------

function Invoke-SingleInteractive {
    Write-Host ""
    Write-Step "Criacao de site individual"

    $title = Read-Host "Titulo do site"
    while ([string]::IsNullOrWhiteSpace($title)) { $title = Read-Host "Titulo do site (obrigatorio)" }

    $tipo = ''
    while ($tipo -notin @('Team', 'Communication')) { $tipo = Read-Host "Tipo (Team / Communication)" }

    $alias = ''
    $url   = ''
    $priv  = 'Private'
    if ($tipo -eq 'Team') {
        $alias = Read-Host "Alias (nome curto do grupo, sem espacos)"
        $p = Read-Host "Privacidade (Public / Private) [Private]"
        if ($p -in @('Public', 'Private')) { $priv = $p }
    }
    else {
        $url = Read-Host "Url (URL completa ou apenas o sufixo, ex.: marketing)"
    }

    $owners  = Read-Host "Owners (e-mails separados por ;)"
    $members = Read-Host "Membros (e-mails separados por ;)"
    $desc    = Read-Host "Descricao (opcional)"

    $url2 = New-SharePointSite `
        -Title $title `
        -Type $tipo `
        -Alias $alias `
        -Url $url `
        -Privacy $priv `
        -Owners (Split-Emails $owners) `
        -Members (Split-Emails $members) `
        -Description $desc

    Write-Host ""
    Write-Ok "Concluido: $url2"
}

# ----------------------------------------------------------------------------
# Fluxo principal
# ----------------------------------------------------------------------------

try {
    Assert-PowerShell7
    Initialize-Log -Dir $LogPath
    Write-Step "Log desta execucao: $($script:LogFile)"

    Ensure-PnPModule
    Connect-Tenant -Url $AdminUrl

    if ($CsvPath) {
        # Execucao direta em lote (sem menu)
        Invoke-BatchFromCsv -Path $CsvPath
    }
    else {
        # Menu interativo
        do {
            Write-Host ""
            Write-Host "===================================================" -ForegroundColor White
            Write-Host "   CRIADOR DE SITES - SHAREPOINT ONLINE" -ForegroundColor White
            Write-Host "===================================================" -ForegroundColor White
            Write-Host "  1) Criar sites EM LOTE (a partir de um CSV)"
            Write-Host "  2) Criar um site UNICO"
            Write-Host "  3) Sair"
            Write-Host "---------------------------------------------------"
            $opt = Read-Host "Escolha uma opcao"

            switch ($opt) {
                '1' {
                    $p = Read-Host "Caminho do CSV [.\modelo-sites.csv]"
                    if ([string]::IsNullOrWhiteSpace($p)) { $p = Join-Path $PSScriptRoot 'modelo-sites.csv' }
                    Invoke-BatchFromCsv -Path $p
                }
                '2' { Invoke-SingleInteractive }
                '3' { Write-Host "Saindo..." -ForegroundColor Gray }
                default { Write-Warn2 "Opcao invalida." }
            }
        } while ($opt -ne '3')
    }
}
catch {
    Write-Err2 $_.Exception.Message
    exit 1
}
finally {
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
}
