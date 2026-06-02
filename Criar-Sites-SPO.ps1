<#
.SYNOPSIS
    Cria sites no SharePoint Online em lote (CSV) ou individual usando o
    modulo OFICIAL Microsoft.Online.SharePoint.PowerShell (SPO Management Shell).

.DESCRIPTION
    Alternativa ao Criar-Sites.ps1 (que usa PnP.PowerShell). Use este quando
    quiser ficar apenas no modulo oficial da Microsoft e nao precisar dos recursos
    de Grupo M365.

    >>> LIMITACOES IMPORTANTES DO SPO MANAGEMENT SHELL <<<
      - NAO existe Publico/Privado (isso e propriedade de Grupo M365).
      - NAO adiciona "membros" (nao ha cmdlet no SPO Shell). Por isso este script
        e seu CSV (modelo-sites-spo.csv) NAO tem coluna/pergunta de Membros.
      - Apenas 1 owner principal (-Owner). Owners extras entram como
        ADMINISTRADORES da colecao (Set-SPOUser -IsSiteCollectionAdmin).
      - New-SPOSite nao tem -Description.
      - Cria Communication Site (SITEPAGEPUBLISHING#0) ou Team Site SEM grupo (STS#3).

    Para Publico/Privado, membros e Team Site com Grupo M365, use o Criar-Sites.ps1 (PnP).

.PARAMETER AdminUrl
    URL do centro de administracao do SharePoint. Ex.: https://contoso-admin.sharepoint.com

.PARAMETER CsvPath
    Caminho para um CSV. Se informado, executa direto em modo lote (sem menu).

.PARAMETER DefaultStorageQuota
    Cota de armazenamento padrao (MB) quando o CSV nao informar. Padrao: 1024.

.PARAMETER DelaySeconds
    Pausa (segundos) entre a criacao de cada site no lote. Padrao: 5.

.PARAMETER MaxRetries
    Tentativas em caso de throttling (429/503). Padrao: 5.

.PARAMETER LogPath
    Pasta para os logs. Padrao: <pasta do script>\logs.

.EXAMPLE
    .\Criar-Sites-SPO.ps1 -AdminUrl "https://contoso-admin.sharepoint.com"
    Abre o menu interativo (login no navegador, sem precisar de ClientId).

.EXAMPLE
    .\Criar-Sites-SPO.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -CsvPath ".\modelo-sites-spo.csv"
    Cria em lote a partir do CSV.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [int]$DefaultStorageQuota = 1024,

    [Parameter(Mandatory = $false)]
    [int]$DelaySeconds = 5,

    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 5,

    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'

# Mapeamento Tipo amigavel -> template do SharePoint
$script:Templates = @{
    'Team'          = 'STS#3'              # Team Site moderno SEM grupo M365
    'Communication' = 'SITEPAGEPUBLISHING#0'
}

# ----------------------------------------------------------------------------
# Log
# ----------------------------------------------------------------------------

$script:LogFile = $null

function Initialize-Log {
    param([string]$Dir)
    if ([string]::IsNullOrWhiteSpace($Dir)) { $Dir = Join-Path $PSScriptRoot 'logs' }
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile = Join-Path $Dir "Criar-Sites-SPO_$stamp.log"
    "==== Criar-Sites-SPO - inicio em $(Get-Date -Format 's') ====" | Out-File -FilePath $script:LogFile -Encoding utf8
}

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    if ($script:LogFile) {
        "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Out-File -FilePath $script:LogFile -Append -Encoding utf8
    }
}

function Write-Step  { param([string]$Msg) Write-Host "==> $Msg" -ForegroundColor Cyan;    Write-Log $Msg 'STEP' }
function Write-Ok    { param([string]$Msg) Write-Host "  [OK] $Msg" -ForegroundColor Green;  Write-Log $Msg 'OK' }
function Write-Warn2 { param([string]$Msg) Write-Host "  [!]  $Msg" -ForegroundColor Yellow; Write-Log $Msg 'WARN' }
function Write-Err2  { param([string]$Msg) Write-Host "  [ERRO] $Msg" -ForegroundColor Red;  Write-Log $Msg 'ERRO' }

# ----------------------------------------------------------------------------
# Retry com backoff (throttling 429/503)
# ----------------------------------------------------------------------------

function Invoke-WithRetry {
    param([Parameter(Mandatory)] [scriptblock]$Script, [string]$What = 'operacao', [int]$Retries = $MaxRetries)
    $attempt = 0
    while ($true) {
        try { return & $Script }
        catch {
            $attempt++
            $msg = $_.Exception.Message
            $isTransient = $msg -match '(?i)throttl|too many requests|429|503|temporarily'
            if ($attempt -ge $Retries -or -not $isTransient) { throw }
            $wait = [math]::Min(60, [math]::Pow(2, $attempt))
            Write-Warn2 "Throttling/erro transitorio em '$What' (tentativa $attempt/$Retries). Aguardando $wait s..."
            Start-Sleep -Seconds $wait
        }
    }
}

# ----------------------------------------------------------------------------
# Ambiente / conexao
# ----------------------------------------------------------------------------

function Ensure-SPOModule {
    $name = 'Microsoft.Online.SharePoint.PowerShell'

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        # O modulo SPO e baseado em .NET Framework e roda no Windows PowerShell.
        # No PS7 importamos em modo de compatibilidade (-UseWindowsPowerShell), que
        # cria uma sessao do Windows PowerShell 5.1 por baixo. Por isso o modulo
        # precisa estar instalado para o WINDOWS POWERSHELL, nao para o PS7 -- entao
        # verificamos/instalamos atraves do proprio powershell.exe.
        $winPS = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        if (-not (Test-Path $winPS)) { $winPS = 'powershell.exe' }

        $check = & $winPS -NoProfile -Command "if (Get-Module -ListAvailable -Name '$name') { 'OK' } else { 'MISSING' }"
        if ("$check" -notmatch 'OK') {
            Write-Step "Modulo $name nao encontrado no Windows PowerShell. Instalando..."
            $install = @"
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null }
Install-Module -Name '$name' -Scope CurrentUser -Force -AllowClobber
"@
            & $winPS -NoProfile -Command $install
            if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar '$name' no Windows PowerShell." }
        }
        Import-Module $name -UseWindowsPowerShell -ErrorAction Stop -WarningAction SilentlyContinue
    }
    else {
        if (-not (Get-Module -ListAvailable -Name $name)) {
            Write-Step "Modulo $name nao encontrado. Instalando..."
            try {
                [Net.ServicePointManager]::SecurityProtocol = `
                    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            } catch { }
            if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
            }
            Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module $name -ErrorAction Stop
    }
    Write-Ok "Modulo SPO Management Shell carregado."
}

function Connect-SPO {
    Write-Step "Conectando em $AdminUrl (login no navegador)..."
    Invoke-WithRetry -What "conexao SPO" -Script { Connect-SPOService -Url $AdminUrl }
    Write-Ok "Conectado."
}

function Get-TenantRootUrl {
    param([string]$Admin)
    return ($Admin -replace '-admin\.sharepoint\.com', '.sharepoint.com').TrimEnd('/')
}

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
    $seenUrl = @{}
    $emailRx = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    $i = 0
    foreach ($r in $Rows) {
        $i++
        $tipo = ($r.Tipo -as [string]); if ($tipo) { $tipo = $tipo.Trim() }

        if ([string]::IsNullOrWhiteSpace($r.Titulo)) { $errors += "Linha ${i}: 'Titulo' vazio." }
        if ($tipo -notin @('Team', 'Communication')) {
            $errors += "Linha ${i}: 'Tipo' invalido ('$($r.Tipo)'). Use Team ou Communication."
        }
        if ([string]::IsNullOrWhiteSpace($r.Url)) {
            $errors += "Linha ${i}: 'Url' e obrigatoria."
        }
        else {
            $key = $r.Url.Trim().ToLower()
            if ($seenUrl.ContainsKey($key)) { $errors += "Linha ${i}: 'Url' duplicada no CSV ('$($r.Url)') - tambem na linha $($seenUrl[$key])." }
            else { $seenUrl[$key] = $i }
        }
        if ([string]::IsNullOrWhiteSpace($r.Owner)) {
            $errors += "Linha ${i}: 'Owner' (owner principal) e obrigatorio no modo SPO."
        }
        elseif ((Split-Emails $r.Owner).Count -gt 1) {
            $errors += "Linha ${i}: 'Owner' aceita apenas UM usuario (use 'Admins' para os demais)."
        }
        foreach ($e in (@(Split-Emails $r.Owner) + @(Split-Emails $r.Admins))) {
            if ($e -notmatch $emailRx) { $errors += "Linha ${i}: e-mail invalido ('$e')." }
        }
        if ($r.StorageQuota -and ($r.StorageQuota.Trim() -notmatch '^\d+$')) {
            $errors += "Linha ${i}: 'StorageQuota' deve ser numero (MB)."
        }
    }
    return $errors
}

# ----------------------------------------------------------------------------
# Criacao de site (nucleo)
# ----------------------------------------------------------------------------

function New-SpoSiteCollection {
    param(
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [ValidateSet('Team', 'Communication')] [string]$Type,
        [Parameter(Mandatory)] [string]$Url,
        [Parameter(Mandatory)] [string]$Owner,
        [string[]]$Admins = @(),
        [int]$StorageQuota = 0
    )

    $tenantRoot = Get-TenantRootUrl -Admin $AdminUrl
    if ($Url -notmatch '^https?://') { $fullUrl = "$tenantRoot/sites/$($Url.TrimStart('/'))" }
    else { $fullUrl = $Url }

    if ($StorageQuota -le 0) { $StorageQuota = $DefaultStorageQuota }
    $template = $script:Templates[$Type]

    Write-Step "Criando $Type Site '$Title' ($fullUrl) [template $template, owner $Owner]..."
    Invoke-WithRetry -What "criacao de '$Title'" -Script {
        New-SPOSite -Url $fullUrl -Owner $Owner -Title $Title -Template $template -StorageQuota $StorageQuota
    }
    Write-Ok "Site criado: $fullUrl"

    # Owners adicionais -> administradores da colecao de sites.
    foreach ($a in $Admins) {
        try {
            Invoke-WithRetry -What "admin '$a'" -Script {
                Set-SPOUser -Site $fullUrl -LoginName $a -IsSiteCollectionAdmin $true | Out-Null
            }
            Write-Ok "Admin (owner adicional): $a"
        } catch { Write-Warn2 "Falha ao definir admin '$a': $($_.Exception.Message)" }
    }

    return $fullUrl
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
        $rows = Import-Csv -Path $Path
    }
    $rows = @($rows)
    Write-Ok "$($rows.Count) linha(s) encontrada(s)."

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
            $sq = 0
            if ($row.StorageQuota -and $row.StorageQuota.Trim() -match '^\d+$') { $sq = [int]$row.StorageQuota.Trim() }

            $url = New-SpoSiteCollection `
                -Title        $row.Titulo `
                -Type         $row.Tipo.Trim() `
                -Url          $row.Url `
                -Owner        ($row.Owner.Trim()) `
                -Admins       (Split-Emails $row.Admins) `
                -StorageQuota $sq

            $resultados += [pscustomobject]@{ Titulo = $row.Titulo; Url = $url; Status = 'Criado' }
        }
        catch {
            Write-Err2 "Falha na linha $idx ('$($row.Titulo)'): $($_.Exception.Message)"
            $resultados += [pscustomobject]@{ Titulo = $row.Titulo; Url = ''; Status = "ERRO: $($_.Exception.Message)" }
        }

        if ($idx -lt $rows.Count -and $DelaySeconds -gt 0) { Start-Sleep -Seconds $DelaySeconds }
    }

    Write-Host ""
    Write-Step "Resumo do lote:"
    $resultados | Format-Table -AutoSize
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
    Write-Step "Criacao de site individual (modo SPO)"

    $title = Read-Host "Titulo do site"
    while ([string]::IsNullOrWhiteSpace($title)) { $title = Read-Host "Titulo do site (obrigatorio)" }

    $tipo = ''
    while ($tipo -notin @('Team', 'Communication')) { $tipo = Read-Host "Tipo (Team / Communication)" }

    $url = Read-Host "Url (URL completa ou apenas o sufixo, ex.: marketing)"
    while ([string]::IsNullOrWhiteSpace($url)) { $url = Read-Host "Url (obrigatorio)" }

    $owner = Read-Host "Owner principal (UM e-mail)"
    while ([string]::IsNullOrWhiteSpace($owner)) { $owner = Read-Host "Owner principal (obrigatorio)" }

    $admins  = Read-Host "Admins / owners adicionais (e-mails separados por ;)"
    $sqText  = Read-Host "Cota de armazenamento em MB [$DefaultStorageQuota]"
    $sq = 0
    if ($sqText -match '^\d+$') { $sq = [int]$sqText }

    $url2 = New-SpoSiteCollection `
        -Title $title `
        -Type $tipo `
        -Url $url `
        -Owner $owner.Trim() `
        -Admins (Split-Emails $admins) `
        -StorageQuota $sq

    Write-Host ""
    Write-Ok "Concluido: $url2"
}

# ----------------------------------------------------------------------------
# Fluxo principal
# ----------------------------------------------------------------------------

try {
    Initialize-Log -Dir $LogPath
    Write-Step "Log desta execucao: $($script:LogFile)"
    Write-Warn2 "Modo SPO: sem Publico/Privado, sem membros e apenas 1 owner principal. Para isso use Criar-Sites.ps1 (PnP)."

    Ensure-SPOModule
    Connect-SPO

    if ($CsvPath) {
        Invoke-BatchFromCsv -Path $CsvPath
    }
    else {
        do {
            Write-Host ""
            Write-Host "===================================================" -ForegroundColor White
            Write-Host "   CRIADOR DE SITES - SPO MANAGEMENT SHELL" -ForegroundColor White
            Write-Host "===================================================" -ForegroundColor White
            Write-Host "  1) Criar sites EM LOTE (a partir de um CSV)"
            Write-Host "  2) Criar um site UNICO"
            Write-Host "  3) Sair"
            Write-Host "---------------------------------------------------"
            $opt = Read-Host "Escolha uma opcao"

            switch ($opt) {
                '1' {
                    $p = Read-Host "Caminho do CSV [.\modelo-sites-spo.csv]"
                    if ([string]::IsNullOrWhiteSpace($p)) { $p = Join-Path $PSScriptRoot 'modelo-sites-spo.csv' }
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
    try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
}
