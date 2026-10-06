#requires -Version 5.1

<#
.SYNOPSIS
    Recopie la navigation de gauche (liens personnalises : bibliotheques, listes, liens) du site source vers le site cible.

.DESCRIPTION
    Les listes et bibliotheques creees par Microsoft Graph (copieDesFichiers3-historique2.ps1)
    n'apparaissent pas dans la navigation du site. Ce script relit les noeuds de navigation
    "lancement rapide" du site source et cree ceux qui manquent sur la cible (meme titre = ignore),
    en remplacant l'URL du site source par celle du site cible.

    $Apply = $false : simulation, rien n'est modifie.
    Necessite PnP.PowerShell. Droits : FullControl sur les deux sites.
#>

# ============================================================================
# CONFIGURATION
# ============================================================================
$SourceSiteUrl = 'https://infoprodigital365.sharepoint.com/sites/fethygate'
$TargetSiteUrl = 'https://infoprodigital365.sharepoint.com/sites/fethiMercury'

$SourceTenantId = '2eea08b8-1972-447b-ad43-d044d042500a'
$SourceClientId = '821109a4-6e0e-48e8-b477-8f9b70aa32a4'
$SourceCertificatePath = 'C:\certs\MSGraphExchangeOnlineAuth20260717_2.pfx'
$SourceCertificatePassword = ''

# Vide = memes valeurs que la source.
$TargetTenantId = ''
$TargetClientId = ''
$TargetCertificatePath = ''
$TargetCertificatePassword = ''

$Apply = $false

# Emplacement a copier : QuickLaunch (menu de gauche) ou TopNavigationBar (barre du haut).
$Locations = @('QuickLaunch')

# ============================================================================
if ([string]::IsNullOrWhiteSpace($TargetTenantId)) { $TargetTenantId = $SourceTenantId }
if ([string]::IsNullOrWhiteSpace($TargetClientId)) { $TargetClientId = $SourceClientId }
if ([string]::IsNullOrWhiteSpace($TargetCertificatePath)) { $TargetCertificatePath = $SourceCertificatePath }
if ([string]::IsNullOrWhiteSpace($TargetCertificatePassword)) { $TargetCertificatePassword = $SourceCertificatePassword }

# PnP.PowerShell 2.x/3.x exige PowerShell 7 : relance automatique si pwsh.exe est installe.
if ($PSVersionTable.PSVersion.Major -lt 7) {
    $pwshCommand = Get-Command -Name pwsh.exe -ErrorAction SilentlyContinue
    if ($pwshCommand -and $PSCommandPath) {
        Write-Host "Windows PowerShell $($PSVersionTable.PSVersion) detecte : relance dans PowerShell 7 ($($pwshCommand.Source))..." -ForegroundColor Cyan
        & $pwshCommand.Source -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $color = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } 'SUCCESS' { 'Green' } default { 'White' } }
    Write-Host ("[{0}][{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message) -ForegroundColor $color
}

function Connect-Site {
    param([string]$Url, [string]$Tenant, [string]$ClientId, [string]$CertPath, [string]$CertPassword)
    if (-not (Test-Path -LiteralPath $CertPath -PathType Leaf)) { throw "Certificat introuvable: $CertPath" }
    if ([string]::IsNullOrWhiteSpace($CertPassword)) { throw "Mot de passe du certificat vide pour $CertPath" }
    return Connect-PnPOnline -Url $Url -Tenant $Tenant -ClientId $ClientId -CertificatePath $CertPath `
        -CertificatePassword (ConvertTo-SecureString -String $CertPassword -AsPlainText -Force) -ReturnConnection
}

function Convert-NodeUrl {
    param([string]$Url, [string]$SourceWebUrl, [string]$TargetWebUrl)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $Url }
    if ($Url.StartsWith($SourceWebUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $TargetWebUrl.TrimEnd('/') + $Url.Substring($SourceWebUrl.Length)
    }
    return $Url
}

function Add-MissingNodes {
    param($Nodes, $ParentId, [string]$Location, $Conn, [string]$SrcWeb, [string]$DstWeb, $ExistingByTitle)

    foreach ($node in @($Nodes)) {
        $title = [string]$node.Title
        $url = Convert-NodeUrl -Url ([string]$node.Url) -SourceWebUrl $SrcWeb -TargetWebUrl $DstWeb
        if ([string]::IsNullOrWhiteSpace($title) -or [string]::IsNullOrWhiteSpace($url)) {
            Write-Log ("noeud ignore (titre ou URL vide): '{0}' / '{1}'" -f $title, $node.Url) 'WARN'
            continue
        }

        $key = "$ParentId|$title".ToLowerInvariant()
        $newId = $null
        if ($ExistingByTitle.ContainsKey($key)) {
            $newId = $ExistingByTitle[$key]
            Write-Log ("deja presente: {0}" -f $title)
        } elseif (-not $Apply) {
            Write-Log ("[Simulation] ajout: {0} -> {1}" -f $title, $url)
        } else {
            $params = @{ Location = $Location; Title = $title; Url = $url; Connection = $Conn }
            if ($ParentId) { $params['Parent'] = [int]$ParentId }
            if ($url -match '^https?://' -and $url -notlike ($DstWeb.TrimEnd('/') + '*')) { $params['External'] = $true }
            $created = Add-PnPNavigationNode @params
            $newId = $created.Id
            Write-Log ("ajoutee: {0} -> {1}" -f $title, $url) 'SUCCESS'
        }

        $children = @($node.Children)
        if ($children.Count -gt 0 -and $newId) {
            Add-MissingNodes -Nodes $children -ParentId $newId -Location $Location -Conn $Conn -SrcWeb $SrcWeb -DstWeb $DstWeb -ExistingByTitle $ExistingByTitle
        }
    }
}

# Les bibliotheques/listes apparaissent dans le menu de gauche grace a leur propriete OnQuickLaunch
# ("Afficher dans la navigation"), pas via des noeuds de navigation : on recopie cette propriete.
function Sync-ListNavigation {
    param($Src, $Dst, [string]$SrcWeb, [string]$DstWeb)

    $props = 'Title', 'Hidden', 'IsSystemList', 'OnQuickLaunch', 'RootFolder'
    $srcLists = @(Get-PnPList -Includes $props -Connection $Src)
    $dstMap = @{}
    foreach ($l in @(Get-PnPList -Includes $props -Connection $Dst)) {
        $rel = ([string]$l.RootFolder.ServerRelativeUrl).Substring($DstWeb.Length).TrimStart('/').ToLowerInvariant()
        $dstMap[$rel] = $l
    }

    foreach ($l in $srcLists) {
        if ($l.Hidden -or $l.IsSystemList -or -not $l.OnQuickLaunch) { continue }
        $rel = ([string]$l.RootFolder.ServerRelativeUrl).Substring($SrcWeb.Length).TrimStart('/')
        $target = $dstMap[$rel.ToLowerInvariant()]
        if (-not $target) {
            Write-Log ("liste absente de la cible: {0}" -f $rel) 'WARN'
            continue
        }
        if ($target.OnQuickLaunch) {
            Write-Log ("deja dans la navigation: {0}" -f $rel)
        } elseif (-not $Apply) {
            Write-Log ("[Simulation] affichage dans la navigation: {0}" -f $rel)
        } else {
            Set-PnPList -Identity $target -OnQuickLaunch $true -Connection $Dst
            Write-Log ("ajoutee a la navigation: {0}" -f $rel) 'SUCCESS'
        }
    }
}

try {
    Write-Log ("Mode: {0}" -f $(if ($Apply) { 'APPLICATION' } else { 'SIMULATION (aucune modification)' }))
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) { throw 'Module PnP.PowerShell absent. Installez PowerShell 7 puis, dans pwsh : Install-Module PnP.PowerShell -Scope CurrentUser' }
    Import-Module PnP.PowerShell -DisableNameChecking

    $src = Connect-Site -Url $SourceSiteUrl -Tenant $SourceTenantId -ClientId $SourceClientId -CertPath $SourceCertificatePath -CertPassword $SourceCertificatePassword
    $dst = Connect-Site -Url $TargetSiteUrl -Tenant $TargetTenantId -ClientId $TargetClientId -CertPath $TargetCertificatePath -CertPassword $TargetCertificatePassword

    $srcWeb = (Get-PnPWeb -Connection $src).ServerRelativeUrl
    $dstWeb = (Get-PnPWeb -Connection $dst).ServerRelativeUrl

    Write-Log 'Listes et bibliotheques affichees dans la navigation'
    Sync-ListNavigation -Src $src -Dst $dst -SrcWeb $srcWeb -DstWeb $dstWeb

    foreach ($location in $Locations) {
        Write-Log ("Navigation {0}" -f $location)
        $srcNodes = @(Get-PnPNavigationNode -Location $location -Tree -Connection $src)
        Write-Log ("Source: {0} noeud(s) de premier niveau" -f $srcNodes.Count)

        # Index des noeuds existants sur la cible (parent|titre -> id).
        $existing = @{}
        foreach ($n in @(Get-PnPNavigationNode -Location $location -Tree -Connection $dst)) {
            $existing["|$($n.Title)".ToLowerInvariant()] = $n.Id
            foreach ($c in @($n.Children)) { $existing["$($n.Id)|$($c.Title)".ToLowerInvariant()] = $c.Id }
        }

        Add-MissingNodes -Nodes $srcNodes -ParentId $null -Location $location -Conn $dst -SrcWeb $srcWeb -DstWeb $dstWeb -ExistingByTitle $existing
    }
    Write-Log 'Termine' 'SUCCESS'
}
catch {
    Write-Log ("Erreur: {0}" -f $_.Exception.Message) 'ERROR'
    exit 1
}
