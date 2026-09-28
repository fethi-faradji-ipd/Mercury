# =========================
# CONFIGURATION
# =========================

# App Registration TENANT SOURCE
#$app_source = @{ #recette
#tenant_Id = "85a5a352-25d6-4894-a97d-221cd1712dd2"
#        client_Id = "684f670e-6931-4120-b053-f8c9538744f6"
#        client_Secret = "O5s8Q~Z~e_d1-dSQ.vt4uFTeR8c4Xssaa093Racz"
#}
#$app_source = @{
  #  tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
 #       client_Id = "28d46b2c-b527-4abb-9fb8-013656682b19"
    #    client_Secret = "l-f8Q~-NfTX__dizvu8u1dAWSzwlqBcM~unzpcQa"
#}
# App Registration TENANT CIBLE
#$app_cible = @{
#    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
#    client_id     = " 821109a4-6e0e-48e8-b477-8f9b70aa32a4"
#    client_secret = "Yhy8Q~mBxTrYqx~ve7enqTQatjwhEZ5.34-q-a4j"
#}
$app_source = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = "821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = "XYr8Q~jx8vDX30c97YP5M0-juBIahzni6WlDNdz8"  # laisser vide si utilisation du certificat
    # Authentification par certificat (recommande pour SharePoint REST)
    # Renseigner certificate_thumbprint (certificat dans le store Windows Cert:\CurrentUser\My)
    # OU certificate_path (chemin vers fichier .pfx)
    certificate_thumbprint = ""  # ex : "A1B2C3D4E5F6..."
    certificate_path       = "C:\certs\MSGraphExchangeOnlineAuth20260717_2.pfx"  # ex : "C:\certs\app_cible.pfx"
    certificate_password   = "MotDePasseFort123!"  # mot de passe du PFX si besoin
}

# App Registration TENANT CIBLE
$app_cible = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = "821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = "XYr8Q~jx8vDX30c97YP5M0-juBIahzni6WlDNdz8"  # laisser vide si utilisation du certificat
    # Authentification par certificat (recommande pour SharePoint REST)
    # Renseigner certificate_thumbprint (certificat dans le store Windows Cert:\CurrentUser\My)
    # OU certificate_path (chemin vers fichier .pfx)
    certificate_thumbprint = ""  # ex : "A1B2C3D4E5F6..."
    certificate_path       = "C:\certs\MSGraphExchangeOnlineAuth20260717_2.pfx"  # ex : "C:\certs\app_cible.pfx"
    certificate_password   = "MotDePasseFort123!"  # mot de passe du PFX si besoin
}

$CsvPath = Join-Path $PSScriptRoot 'migrations.csv'
$UserMappingCsvPath = Join-Path $PSScriptRoot 'user_mapping.csv'
# Groupes et comptes sans email : login_source;login_cible. Pas de correspondance par nom approchant.
# Pour un groupe SharePoint cible : login_cible=spgroup:Nom exact du groupe
$PrincipalMappingCsvPath = Join-Path $PSScriptRoot 'principal_mapping.csv'
$CopyMode = 'incremental'          # full | incremental ; premier passage = copie initiale
$CopyVersionHistory = $true
$CopyHistoricalMetadata = $true   # champs de chaque version via listItem/versions
$CopyItemMetadata = $true
$CopyPermissions = $true
$PreserveFileTimestamps = $true
$RestoreItemAuthorEditor = $true
$RequireExplicitUserMapping = $true # proprietaire de la ligne CSV mappe automatiquement
$ChunkSizeMB = 10                   # multiple de 320 Kio, < 60 Mio
$SmallFileThresholdMB = 4           # PUT direct en streaming, economise une requete
$MaxRetries = 5
$RequestTimeoutMinutes = 30
$PageSize = 200
$CacheLimit = 2000
$StateDirectory = Join-Path $PSScriptRoot 'OneDrive_State_v2'
$LogDirectory = Join-Path $PSScriptRoot 'OneDrive_Logs'
$TempTransferDir = Join-Path ([System.IO.Path]::GetTempPath()) 'OneDriveCopy_v2'
# Destination dediee a la migration. Les ACL uniques existantes sont reconciliees
# exactement (ajouts PUIS retraits), sans modifier les groupes eux-memes.
$AllowAppendToUntrackedDestination = $false # true : ajoute tout l'historique a un fichier existant sans etat
# Liens anonymes/organisation : nouvelles URL, meme portee/type, selon politiques cible.
# false : fichier marque PARTIEL si un lien est present, jamais ignore silencieusement.
$RecreateSharingLinks = $true
# Champs systeme excludes par conception (pas des champs metier).
$ExcludedMetadataFields = @('Attachments','ContentType','ContentTypeId','FileLeafRef','FileRef','FileDirRef',
    'FSObjType','GUID','UniqueId','ID','id','DocIcon','ItemChildCount','FolderChildCount',
    'Author','Editor','Created','Modified','AuthorLookupId','EditorLookupId','AppAuthor','AppEditor',
    'AppAuthorLookupId','AppEditorLookupId','ComplianceAssetId','LinkTitle','LinkTitleNoMenu',
    'LinkFilename','LinkFilenameNoMenu','Edit','SelectTitle','PermMask','MetaInfo','Order',
    'Created_x0020_By','Modified_x0020_By','owshiddenversion','File_x0020_Type',
    'SMTotalSize','SMTotalFileStreamSize','HTML_x0020_File_x0020_Type')
# ================= FIN CONFIGURATION =================

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$script:Graph = 'https://graph.microsoft.com/v1.0'
$script:Tokens = @{}
$script:UserMap = @{}
$script:PrincipalMap = @{}
$script:RunId = [guid]::NewGuid().ToString('N')
$script:Stats = @{Files=0; Folders=0; UploadedVersions=0; SkippedContent=0; Errors=0; Bytes=0L; Requests=0; Retries=0}
$script:Http = $null
$script:LogWriter = $null
$script:ReportWriter = $null
$script:AuditWriter = $null
$script:Lock = $null
$script:SourceContext = $null
$script:TargetContext = $null

function Write-Log {
    param([string]$Level, [string]$Message)
    $line = '{0:u} [{1}] {2}' -f [datetime]::UtcNow,$Level,$Message
    Write-Host $line
    if ($script:LogWriter) { $script:LogWriter.WriteLine($line); $script:LogWriter.Flush() }
}
function Write-CsvRow {
    param($Writer,[object[]]$Values)
    $parts = foreach ($v in $Values) { '"' + ([string]$v).Replace('"','""') + '"' }
    $Writer.WriteLine(($parts -join ';')); $Writer.Flush()
}
function Write-Issue {
    param([string]$Path,[string]$Phase,[string]$Message)
    $script:Stats.Errors++
    Write-Log 'ERROR' "$Path [$Phase] $Message"
    Write-CsvRow $script:ReportWriter @([datetime]::UtcNow.ToString('o'),$script:Pair,$Path,$Phase,$Message)
}
function Get-Key {
    param([string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
}
function Save-State {
    param([string]$Path,$State)
    $json = ConvertTo-Json -InputObject $State -Depth 40 -Compress
    $dir = [IO.Path]::GetDirectoryName($Path)
    if ($dir -and -not [IO.Directory]::Exists($dir)) { [void][IO.Directory]::CreateDirectory($dir) }
    $tmp = $Path + '.' + $script:RunId + '.tmp'
    $stream = [IO.File]::Open($tmp,[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::None)
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($json)
        $stream.Write($bytes,0,$bytes.Length); $stream.Flush($true)
    } finally { $stream.Dispose() }
    if ([IO.File]::Exists($Path)) {
        [IO.File]::Copy($tmp,$Path,$true)
        [IO.File]::Delete($tmp)
    }
    else {
        [IO.File]::Move($tmp,$Path)
    }
}
function Read-State {
    param([string]$Path)
    if ([IO.File]::Exists($Path)) {
        # Etat corrompu = erreur, jamais un nouveau full implicite.
        $state = ([IO.File]::ReadAllText($Path) | ConvertFrom-Json)
        if ($state -and -not $state.PSObject.Properties['DestinationContentStamp']) {
            $state | Add-Member -NotePropertyName DestinationContentStamp -NotePropertyValue $null
        }
        return $state
    }
    return $null
}
function Set-Cache {
    param([hashtable]$Cache,[string]$Key,$Value)
    if ($Cache.Count -ge $CacheLimit) { $Cache.Clear() }
    $Cache[$Key] = $Value
}
function Encode-Path {
    param([string]$Path)
    return ((($Path -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/')
}
function ConvertTo-B64Url {
    param([byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+','-').Replace('/','_')
}
function Get-Certificate {
    param([hashtable]$App)
    if ($App.certificate_thumbprint) {
        $thumb = $App.certificate_thumbprint.Replace(' ','')
        $cert = Get-Item "Cert:\CurrentUser\My\$thumb" -ErrorAction SilentlyContinue
        if (-not $cert) { $cert = Get-Item "Cert:\LocalMachine\My\$thumb" -ErrorAction SilentlyContinue }
    } elseif ($App.certificate_path) {
        $pwd = $App.certificate_password
        if (-not $pwd) { $pwd = Read-Host "Mot de passe PFX pour $($App.client_id) (Entree si aucun)" -AsSecureString }
        $cert = [Security.Cryptography.X509Certificates.X509Certificate2]::new(
            $App.certificate_path,$pwd,[Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet)
    } else {
        return $null
    }
    if (-not $cert) { return $null }
    if (-not $cert.HasPrivateKey) { throw 'Certificat introuvable ou sans cle privee.' }
    if ($cert.NotAfter.ToUniversalTime() -le [datetime]::UtcNow) { throw 'Certificat expire.' }
    return $cert
}
function Normalize-AppConfig {
    param([hashtable]$App)

    foreach ($k in @('tenant_id','client_id','client_secret','certificate_thumbprint','certificate_path')) {
        if ($App.ContainsKey($k) -and $null -ne $App[$k]) {
            $App[$k] = ([string]$App[$k]).Trim()
        }
    }
}
function Get-AccessToken {
    param($Context,[string]$Audience,[switch]$Refresh)
    $tenantId = ([string]$Context.App.tenant_id).Trim()
    $clientId = ([string]$Context.App.client_id).Trim()
    $audienceNorm = ([string]$Audience).Trim().TrimEnd('/')
    $key = "$tenantId|$clientId|$audienceNorm"
    if (-not $Refresh -and $script:Tokens.ContainsKey($key) -and $script:Tokens[$key].Expires -gt [datetime]::UtcNow.AddMinutes(5)) {
        return $script:Tokens[$key].Token
    }
    $parsed=[guid]::Empty
    if (-not [guid]::TryParse($tenantId,[ref]$parsed) -or -not [guid]::TryParse($clientId,[ref]$parsed)) {
        throw "Configuration OAuth invalide : tenant_id/client_id non GUID (tenant='$tenantId', client_id='$clientId')."
    }
    $endpoint = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'
    $form = [Collections.Generic.Dictionary[string,string]]::new()
    $form.Add('client_id',$clientId)
    $form.Add('scope',$audienceNorm+'/.default')
    $form.Add('grant_type','client_credentials')
    if ($Context.Cert) {
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $head = @{alg='RS256';typ='JWT';x5t=(ConvertTo-B64Url $Context.Cert.GetCertHash())} | ConvertTo-Json -Compress
        $claims = @{aud=$endpoint;iss=$clientId;sub=$clientId;jti=[guid]::NewGuid().ToString();nbf=$now-60;exp=$now+600} | ConvertTo-Json -Compress
        $unsigned = (ConvertTo-B64Url ([Text.Encoding]::UTF8.GetBytes($head))) + '.' + (ConvertTo-B64Url ([Text.Encoding]::UTF8.GetBytes($claims)))
        $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Context.Cert)
        try { $sig = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($unsigned),[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1) }
        finally { if ($rsa) { $rsa.Dispose() } }
        $form.Add('client_assertion_type','urn:ietf:params:oauth:client-assertion-type:jwt-bearer')
        $form.Add('client_assertion',$unsigned+'.'+(ConvertTo-B64Url $sig))
    } elseif ($Context.App.ContainsKey('client_secret') -and -not [string]::IsNullOrWhiteSpace([string]$Context.App.client_secret)) {
        $form.Add('client_secret',[string]$Context.App.client_secret)
    } else {
        throw "Aucune methode d'authentification validee pour $tenantId : fournir client_secret OU certificat." 
    }
    $content = [Net.Http.FormUrlEncodedContent]::new($form)
    $resp = $null
    try {
        $resp = $script:Http.PostAsync($endpoint,$content).GetAwaiter().GetResult()
        if (-not $resp.IsSuccessStatusCode) {
            $status = [int]$resp.StatusCode
            $errText = if ($resp.Content) { $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult() } else { '' }
            $aadCode = ''
            $aadDesc = ''
            try {
                $j = $errText | ConvertFrom-Json
                $aadCode = [string]$j.error
                $aadDesc = [string]$j.error_description
            } catch {}
            $detail = if ($aadCode) { "$aadCode $aadDesc".Trim() } else { $errText }
            if ($detail.Length -gt 300) { $detail = $detail.Substring(0,300) + '...' }
            $ex = [Exception]::new("Authentification refusee (HTTP $status), tenant=$tenantId, client_id=$clientId, audience=$audienceNorm, detail=$detail")
            $ex.Data['Status'] = $status
            $ex.Data['ApiCode'] = $aadCode
            throw $ex
        }
        $token = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult() | ConvertFrom-Json
        if (-not $token.access_token) { throw 'Reponse OAuth sans access_token.' }
        $script:Tokens[$key] = @{Token=$token.access_token;Expires=[datetime]::UtcNow.AddSeconds([int]$token.expires_in)}
        return $token.access_token
    } finally { $content.Dispose(); if ($resp) { $resp.Dispose() } }
}
function Get-RetrySeconds {
    param($Response,[int]$Attempt)
    if ($Response -and $Response.Headers.RetryAfter) {
        if ($Response.Headers.RetryAfter.Delta) { return [Math]::Max(1,[Math]::Ceiling($Response.Headers.RetryAfter.Delta.TotalSeconds)) }
        if ($Response.Headers.RetryAfter.Date) { return [Math]::Max(1,[Math]::Ceiling(($Response.Headers.RetryAfter.Date-[DateTimeOffset]::UtcNow).TotalSeconds)) }
    }
    return [Math]::Min(60,[Math]::Pow(2,$Attempt)) + (Get-Random -Minimum 0 -Maximum 3)
}
function Wait-Retry {
    param([double]$Seconds)
    $script:Stats.Retries++
    Write-Log 'WARN' "Limitation/retry HTTP : attente $Seconds s."
    while ($Seconds -gt 0) { $part=[Math]::Min(30,$Seconds); Start-Sleep -Seconds $part; $Seconds-=$part }
}
function Invoke-Api {
    param($Context,[string]$Uri,[string]$Method='GET',$Body,[hashtable]$Headers=@{},
        [scriptblock]$ContentFactory,[string]$OutFile,[switch]$Anonymous,[switch]$AllowNotFound,
        [switch]$NoWriteRetry)
    # HttpClient reutilise, messages ET reponses liberes a CHAQUE tentative.
    # ResponseHeadersRead evite la bufferisation integrale des telechargements.
    for ($attempt=1; $attempt -le $MaxRetries; $attempt++) {
        $req=$null; $resp=$null; $cts=$null
        try {
            $req=[Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::new($Method),$Uri)
            if ($Method -eq 'GET' -and $attempt -gt 1) {
                # Hint the server/proxy to use a fresh connection after a transient socket failure.
                $req.Headers.ConnectionClose = $true
            }
            if (-not $Anonymous) {
                $audience = if (([uri]$Uri).Host -eq 'graph.microsoft.com') { 'https://graph.microsoft.com' } else { $Context.HostUrl }
                if (([uri]$Uri).Host -ne 'graph.microsoft.com' -and ([uri]$Uri).Host -ne ([uri]$Context.HostUrl).Host) { throw 'URL API hors du contexte attendu.' }
                $req.Headers.Authorization=[Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer',(Get-AccessToken $Context $audience))
            }
            $req.Headers.Accept.Clear()
            if ($OutFile -or $Uri -match '/content($|\?)') {
                # Binary content endpoints are sensitive to content negotiation; do not send Accept here.
            } elseif (([uri]$Uri).Host -eq 'graph.microsoft.com') {
                # Graph rejects some legacy OData Accept variants on specific endpoints.
                [void]$req.Headers.TryAddWithoutValidation('Accept','application/json')
            } else {
                [void]$req.Headers.TryAddWithoutValidation('Accept','application/json;odata=nometadata')
            }
            foreach ($k in $Headers.Keys) { [void]$req.Headers.TryAddWithoutValidation($k,[string]$Headers[$k]) }
            if ($ContentFactory) { $req.Content = & $ContentFactory }
            elseif ($null -ne $Body) { $req.Content=[Net.Http.StringContent]::new((ConvertTo-Json -InputObject $Body -Depth 40 -Compress),[Text.Encoding]::UTF8,'application/json') }
            $cts=[Threading.CancellationTokenSource]::new([TimeSpan]::FromMinutes($RequestTimeoutMinutes))
            $script:Stats.Requests++
            $resp=$script:Http.SendAsync($req,[Net.Http.HttpCompletionOption]::ResponseHeadersRead,$cts.Token).GetAwaiter().GetResult()
            if (-not $resp) {
                throw 'Reponse HTTP nulle (transitoire) pendant appel API.'
            }
            $status=[int]$resp.StatusCode
            if ($status -eq 404 -and $AllowNotFound) { return $null }
            if ($resp.IsSuccessStatusCode) {
                if ($OutFile) {
                    $file=[IO.File]::Open($OutFile,[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::None)
                    try {
                        $stream=$resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
                        if (-not $stream) { throw 'Flux HTTP de telechargement indisponible.' }
                        try { $stream.CopyToAsync($file,81920,$cts.Token).GetAwaiter().GetResult() }
                        finally { $stream.Dispose() }
                    } finally { $file.Dispose() }
                    return
                }
                $text=$resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                if ($text) { return ($text | ConvertFrom-Json) }
                return $null
            }
            if ($status -eq 401 -and -not $Anonymous -and $attempt -lt $MaxRetries) {
                $null=Get-AccessToken $Context $audience -Refresh
                continue
            }
            # Ne pas journaliser les URL preauthentifiees ni les jetons.
            $bodyText=$resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $code=''
            $apiMessage=''
            try {
                $jsonErr = $bodyText | ConvertFrom-Json
                if ($jsonErr -and $jsonErr.error) {
                    if ($jsonErr.error.code) { $code = [string]$jsonErr.error.code }
                    if ($jsonErr.error.message) { $apiMessage = [string]$jsonErr.error.message }
                }
                if (-not $code -and $jsonErr.error) { $code = [string]$jsonErr.error }
                if (-not $apiMessage -and $jsonErr.error_description) { $apiMessage = [string]$jsonErr.error_description }
            } catch { }

            $retryable = ($status -eq 429 -or $status -in @(500,502,503,504))
            $isUserDriveLookup = ($Method -eq 'GET' -and $Uri -match '/users/.+/drive\?')
            $retryableBadRequest = ($status -eq 400 -and $code -eq 'BadRequest' -and $isUserDriveLookup)
            if (($retryable -or $retryableBadRequest) -and $attempt -lt $MaxRetries -and (-not $NoWriteRetry -or $status -eq 429)) {
                $delay=Get-RetrySeconds $resp $attempt
                $resp.Dispose(); $resp=$null; $req.Dispose(); $req=$null
                Wait-Retry $delay
                continue
            }

            $safeUri = ''
            try {
                $u = [uri]$Uri
                $safeUri = ($u.Scheme + '://' + $u.Host + $u.AbsolutePath)
            } catch {
                $safeUri = '[uri-invalide]'
            }

            $detail = if ($apiMessage) { $apiMessage } else { 'detail indisponible' }
            if ($detail.Length -gt 280) { $detail = $detail.Substring(0,280) + '...' }

            $hint = ''
            if ($status -eq 400 -and $code -eq 'BadRequest' -and $detail -match '(?i)mysite|personal site|onedrive') {
                $hint = ' Verifier que le OneDrive de cet utilisateur est provisionne dans le tenant source.'
            }

            $ex=[Exception]::new("HTTP $status ($code), methode $Method, endpoint=$safeUri, detail=$detail.$hint")
            $ex.Data['Status']=$status
            $ex.Data['ApiCode']=$code
            $ex.Data['ApiMessage']=$apiMessage
            $ex.Data['Endpoint']=$safeUri
            throw $ex
        } catch {
            # Un GET peut etre rejoue sans effet de bord. Une ecriture a issue inconnue ne l'est pas.
            if ($Method -eq 'GET' -and -not $_.Exception.Data.Contains('Status') -and $attempt -lt $MaxRetries) {
                Wait-Retry (Get-RetrySeconds $null $attempt); continue
            }
            throw
        } finally {
            if ($resp) { $resp.Dispose() }; if ($req) { $req.Dispose() }; if ($cts) { $cts.Dispose() }
        }
    }
    throw 'Nombre maximal de tentatives atteint.'
}
function Get-Paged {
    param($Context,[string]$Uri)
    while ($Uri) {
        $page=Invoke-Api $Context $Uri
        foreach ($entry in $page.value) { Write-Output $entry }
        $Uri=$page.'@odata.nextLink'
        if (-not $Uri) { $Uri=$page.'odata.nextLink' }
    }
}
function Get-ItemMeta {
    param($Context,[string]$Id)
    Invoke-Api $Context ($script:Graph+'/drives/'+$Context.DriveId+'/items/'+[uri]::EscapeDataString($Id)+'?$select=id,name,size,eTag,cTag,file,folder,package,remoteItem,parentReference,sharepointIds,createdDateTime,lastModifiedDateTime,fileSystemInfo')
}
function Get-TargetByPath {
    param([string]$Path)
    Invoke-Api $script:TargetContext ($script:Graph+'/drives/'+$script:TargetContext.DriveId+'/root:/'+(Encode-Path $Path)+'?$select=id,name,size,eTag,cTag,folder,file,sharepointIds,parentReference,fileSystemInfo') -AllowNotFound
}
function Get-ItemApi {
    param($Context,$Item)
    if (-not $Item.sharepointIds.listItemId) { throw "listItemId absent : $($Item.name)" }
    return $Context.ListApi+'/items('+[int]$Item.sharepointIds.listItemId+')'
}
function Get-VersionStamp {
    param($Version)
    return Get-Key ("$($Version.id)|$($Version.lastModifiedDateTime)|$($Version.size)")
}
function Get-ContentStamp {
    param($Item)
    if ($Item.cTag) { return [string]$Item.cTag }
    return [string]$Item.eTag
}
function New-Context {
    param([hashtable]$App,[string]$Upn,[string]$ProvidedUrl)
    $ctx=@{App=$App;Cert=(Get-Certificate $App);HostUrl='';Users=@{};Principals=@{};Fields=@{};Roles=@{}}
    $drive=Invoke-Api $ctx ($script:Graph+'/users/'+[uri]::EscapeDataString($Upn)+'/drive?$select=id,webUrl')
    if (-not $drive.id) { throw "OneDrive introuvable : $Upn" }
    $ctx.DriveId=[string]$drive.id
    $root=Invoke-Api $ctx ($script:Graph+'/drives/'+$drive.id+'/root?$select=id,sharepointIds,webUrl')
    $ctx.Root=$root
    $ctx.SiteUrl=[string]$root.sharepointIds.siteUrl
    if (-not $ctx.SiteUrl -or -not $root.sharepointIds.listId) { throw 'OneDrive Business avec sharepointIds requis.' }
    $ctx.SiteUrl=$ctx.SiteUrl.TrimEnd('/')
    $ctx.HostUrl=([uri]$ctx.SiteUrl).GetLeftPart([UriPartial]::Authority)
    # Le drive appartient exactement a l'UPN. Jamais de startsWith + premier resultat.
    $provided=([uri]$ProvidedUrl).GetLeftPart([UriPartial]::Path).TrimEnd('/')
    $layout=$provided.IndexOf('/_layouts/',[StringComparison]::OrdinalIgnoreCase)
    if ($layout -ge 0) { $provided=$provided.Substring(0,$layout) }
    if ($provided -ne $ctx.SiteUrl -and $provided -ne ([string]$drive.webUrl).TrimEnd('/')) {
        throw "L'URL CSV ne correspond pas au OneDrive resolu pour $Upn."
    }
    $ctx.SiteId=([uri]$ctx.SiteUrl).Host+','+$root.sharepointIds.siteId+','+$root.sharepointIds.webId
    $ctx.ListId=[string]$root.sharepointIds.listId
    $ctx.ListApi=$ctx.SiteUrl+"/_api/web/lists(guid'"+$ctx.ListId+"')"
    $ctx.List=Invoke-Api $ctx ($ctx.ListApi+'?$select=Id,EnableVersioning,EnableMinorVersions,ForceCheckout,MajorVersionLimit')
    foreach ($f in (Get-Paged $ctx ($ctx.ListApi+'/fields?$select=InternalName,TypeAsString,ReadOnlyField,Hidden,FromBaseType'))) { $ctx.Fields[[string]$f.InternalName]=$f }
    foreach ($r in (Get-Paged $ctx ($ctx.SiteUrl+'/_api/web/roledefinitions?$select=Id,Name,RoleTypeKind,BasePermissions'))) { $ctx.Roles[[string]$r.Id]=$r }
    return $ctx
}
function Load-Mappings {
    if (Test-Path -LiteralPath $UserMappingCsvPath) {
        Import-Csv -LiteralPath $UserMappingCsvPath -Delimiter ';' | ForEach-Object {
            if (-not $_.email_source -or -not $_.email_cible) { throw 'user_mapping.csv : email_source;email_cible requis.' }
            $script:UserMap[$_.email_source.Trim().ToLowerInvariant()]=$_.email_cible.Trim().ToLowerInvariant()
        }
    }
    if (Test-Path -LiteralPath $PrincipalMappingCsvPath) {
        Import-Csv -LiteralPath $PrincipalMappingCsvPath -Delimiter ';' | ForEach-Object {
            if (-not $_.login_source -or -not $_.login_cible) { throw 'principal_mapping.csv : login_source;login_cible requis.' }
            $script:PrincipalMap[$_.login_source.Trim()]=$_.login_cible.Trim()
        }
    }
}
function Resolve-Email {
    param([string]$SourceEmail)
    if (-not $SourceEmail) { throw 'Identite source absente : impossible de reconstituer auteur ou droits sans inventer.' }
    $email=$SourceEmail.Trim().ToLowerInvariant()
    if ($email -eq $script:SourceUpn.ToLowerInvariant()) { return $script:TargetUpn.ToLowerInvariant() }
    if ($script:UserMap.ContainsKey($email)) { return $script:UserMap[$email] }
    if ($RequireExplicitUserMapping) { throw "Mapping manquant pour $email dans user_mapping.csv." }
    return $email
}
function Get-SourcePerson {
    param([int]$Id)
    $ctx=$script:SourceContext
    if ($ctx.Users.ContainsKey([string]$Id)) { return $ctx.Users[[string]$Id] }
    $person=Invoke-Api $ctx ($ctx.SiteUrl+'/_api/web/siteusers/getbyid('+ $Id +')?$select=Id,LoginName,Email,PrincipalType,Title')
    Set-Cache $ctx.Users ([string]$Id) $person
    return $person
}
function Resolve-Principal {
    param($Member)
    $ctx=$script:TargetContext
    $key=[string]$Member.LoginName
    if ($ctx.Principals.ContainsKey($key)) { return $ctx.Principals[$key] }
    if ($script:PrincipalMap.ContainsKey($key)) { $login=$script:PrincipalMap[$key] }
    elseif ([int]$Member.PrincipalType -eq 1) {
        $email=[string]$Member.Email
        if (-not $email -and $key -like 'i:0#.f|membership|*') { $email=$key.Split('|')[-1] }
        $login='i:0#.f|membership|'+(Resolve-Email $email)
    } else { throw "Mapping de groupe/principal manquant : $key (type $($Member.PrincipalType))." }
    if ($login.StartsWith('spgroup:')) {
        $group=[uri]::EscapeDataString($login.Substring(8).Replace("'","''"))
        $person=Invoke-Api $ctx ($ctx.SiteUrl+"/_api/web/sitegroups/getbyname('"+$group+"')?`$select=Id,LoginName,Title")
    } else {
        $person=Invoke-Api $ctx ($ctx.SiteUrl+'/_api/web/ensureuser') -Method POST -Body @{logonName=$login}
    }
    if (-not $person.Id) { throw "Principal cible introuvable : $login" }
    Set-Cache $ctx.Principals $key $person
    return $person
}
function Get-PeopleFormValue {
    param($LookupIds)
    $claims=[Collections.Generic.List[object]]::new()
    foreach ($id in @($LookupIds)) {
        if ($null -eq $id -or [string]$id -eq '') { continue }
        $p=Resolve-Principal (Get-SourcePerson ([int]$id))
        $claims.Add(@{Key=[string]$p.LoginName})
    }
    return ConvertTo-Json -InputObject @($claims.ToArray()) -Compress
}
function Test-MetadataSchema {
    $script:BusinessFields=[Collections.Generic.List[string]]::new()
    if (-not $CopyItemMetadata -and -not $CopyHistoricalMetadata) { return }
    $supported=@('Text','Note','Number','Currency','Boolean','DateTime','Choice','MultiChoice','URL','User','UserMulti')
    foreach ($name in ($script:SourceContext.Fields.Keys | Sort-Object)) {
        $f=$script:SourceContext.Fields[$name]
        # OneDrive internal Add-to-OneDrive technical fields are not business metadata.
        if ($name -like 'A2OD*') { continue }
        if ($name -in $ExcludedMetadataFields -or $name.StartsWith('_') -or $f.ReadOnlyField) { continue }
        if ($f.FromBaseType -and $name -ne 'Title') { continue }
        if ($f.TypeAsString -notin $supported) { throw "Champ metier non pris en charge automatiquement : $name ($($f.TypeAsString)). Mapping specifique requis avant migration." }
        if (-not $script:TargetContext.Fields.ContainsKey($name)) { throw "Creer le champ cible '$name' ($($f.TypeAsString)) avant migration." }
        $d=$script:TargetContext.Fields[$name]
        if ($d.ReadOnlyField -or $d.TypeAsString -ne $f.TypeAsString) { throw "Schema incompatible pour '$name'." }
        $script:BusinessFields.Add($name)
    }
}
function Get-SourceFields {
    param($Item,[string]$VersionId='')
    if ($VersionId) {
        $uri=$script:Graph+'/sites/'+$script:SourceContext.SiteId+'/lists/'+$script:SourceContext.ListId+'/items/'+$Item.sharepointIds.listItemId+'/versions/'+[uri]::EscapeDataString($VersionId)+'?$expand=fields'
        $v=Invoke-Api $script:SourceContext $uri
        if (-not $v.fields) { throw "Metadonnees de version $VersionId indisponibles." }
        return $v.fields
    }
    return Invoke-Api $script:SourceContext ($script:Graph+'/drives/'+$script:SourceContext.DriveId+'/items/'+$Item.id+'/listItem/fields')
}
function Convert-ToSharePointDateString {
    param($RawValue,[string]$Format='yyyy-MM-ddTHH:mm:ssZ')

    if ($null -eq $RawValue -or [string]::IsNullOrWhiteSpace([string]$RawValue)) { return $null }

    $parsed = $null
    try { $parsed = [datetime]$RawValue } catch { return $null }
    $utc = $parsed.ToUniversalTime()

    $min = [datetime]::SpecifyKind([datetime]'1900-01-01T00:00:00',[DateTimeKind]::Utc)
    $max = [datetime]::SpecifyKind([datetime]'8900-12-31T23:59:59',[DateTimeKind]::Utc)
    if ($utc -lt $min -or $utc -gt $max) { return $null }

    return $utc.ToString($Format)
}
function New-MetadataPlan {
    param($Fields,[switch]$Historical,$FileSystemInfo,[switch]$IsFolder,[string]$ItemPath='')
    $values=[Collections.Generic.List[object]]::new()
    if (($Historical -and $CopyHistoricalMetadata) -or (-not $Historical -and $CopyItemMetadata)) {
        foreach ($name in $script:BusinessFields) {
            $type=$script:SourceContext.Fields[$name].TypeAsString
            if ($type -in @('User','UserMulti')) {
                $lookupName=$name+'LookupId'
                $value=Get-PeopleFormValue $Fields.$lookupName
            } else {
                $raw=$Fields.$name
                $value=''
                if ($null -ne $raw) {
                    switch ($type) {
                        'Boolean' { $value=if ([bool]$raw) { '1' } else { '0' } }
                        'DateTime' {
                            $dateVal = Convert-ToSharePointDateString $raw
                            if ($dateVal) { $value=$dateVal }
                            else { $value='' }
                        }
                        'MultiChoice' { $value=if (@($raw).Count) { ';#'+(@($raw) -join ';#')+';#' } else { '' } }
                        'URL' {
                            if ($raw -is [string]) { $value=$raw }
                            elseif ($raw.Url) { $value=[string]$raw.Url+', '+[string]$raw.Description }
                            else { throw "URL non interpretable : $name" }
                        }
                        default { $value=[Convert]::ToString($raw,[Globalization.CultureInfo]::InvariantCulture) }
                    }
                }
            }
            $values.Add([ordered]@{FieldName=$name;FieldValue=$value})
        }
    }
    $expected=@{}
    if ($RestoreItemAuthorEditor -and -not $IsFolder) {
        foreach ($field in @('Author','Editor')) {
            $id=$Fields.($field+'LookupId')
            if (-not $id) { throw "Identifiant $field absent des metadonnees source." }
            $person=Resolve-Principal (Get-SourcePerson ([int]$id))
            $values.Add([ordered]@{FieldName=$field;FieldValue=(ConvertTo-Json -InputObject @(@{Key=[string]$person.LoginName}) -Compress)})
            $expected[$field+'Id']=[int]$person.Id
        }
    }
    if ($PreserveFileTimestamps -and -not $IsFolder) {
        # Ne pas pousser Created/Modified via ValidateUpdateListItem : ces champs systeme sont trop souvent refuses.
        # Les horodatages sont preserves via fileSystemInfo quand SharePoint les accepte.
        foreach ($field in @('Created','Modified')) {
            if ($Fields.$field -or $FileSystemInfo) { continue }
            $targetName = if ($ItemPath) { $ItemPath } else { '[fichier]' }
            Write-Log 'WARN' "$targetName : date $field non exploitable, champ ignore."
        }
    }
    $fsInfo=$null
    if ($PreserveFileTimestamps -and $FileSystemInfo) {
        $fsInfo=[ordered]@{}
        foreach ($field in @('createdDateTime','lastModifiedDateTime')) {
            if ($FileSystemInfo.$field) {
                $date = Convert-ToSharePointDateString $FileSystemInfo.$field 'o'
                if ($date) { $fsInfo[$field]=$date }
            }
        }
    }
    return [pscustomobject][ordered]@{FileSystemInfo=$fsInfo;Values=@($values.ToArray());Expected=[pscustomobject]([ordered]@{AuthorId=$expected.AuthorId;EditorId=$expected.EditorId})}
}
function Normalize-MetadataPlan {
    param($Plan,[string]$ItemPath='')
    if (-not $Plan) { return $Plan }
    $values=[Collections.Generic.List[object]]::new()
    $removed=$false
    foreach ($entry in @($Plan.Values)) {
        if ($entry.FieldName -in @('Created','Modified')) { $removed=$true; continue }
        $values.Add($entry)
    }
    if ($removed) {
        $targetName = if ($ItemPath) { $ItemPath } else { '[item]' }
        Write-Log 'WARN' "$targetName : plan de metadonnees herite nettoye (Created/Modified ignores)."
    }
    $expected=@{}
    if ($Plan.Expected) {
        foreach ($prop in $Plan.Expected.PSObject.Properties) {
            if ($prop.Name -in @('Created','Modified')) { continue }
            $expected[$prop.Name]=$prop.Value
        }
    }
    return [pscustomobject][ordered]@{FileSystemInfo=$Plan.FileSystemInfo;Values=@($values.ToArray());Expected=[pscustomobject]$expected}
}
function Set-Metadata {
    param($TargetItem,$Plan)
    $Plan = Normalize-MetadataPlan $Plan
    if ($Plan.FileSystemInfo) {
        $body=@{fileSystemInfo=$Plan.FileSystemInfo}
        $result=Invoke-Api $script:TargetContext ($script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$TargetItem.id) -Method PATCH -Headers @{'If-Match'=$TargetItem.eTag} -Body $body
        foreach ($field in @('createdDateTime','lastModifiedDateTime')) {
            if ($Plan.FileSystemInfo.$field -and (-not $result.fileSystemInfo.$field -or [Math]::Abs((([datetime]$result.fileSystemInfo.$field)-([datetime]$Plan.FileSystemInfo.$field)).TotalSeconds) -gt 1)) { throw 'fileSystemInfo non restaure.' }
        }
    }
    if (@($Plan.Values).Count -eq 0) { return $null }
    $base=Get-ItemApi $script:TargetContext $TargetItem
    $res=Invoke-Api $script:TargetContext ($base+'/ValidateUpdateListItem()') -Method POST -Body @{formValues=@($Plan.Values);bNewDocumentUpdate=$true;checkInComment=''}
    if (-not $res -or -not $res.PSObject.Properties['value']) { throw 'ValidateUpdateListItem : reponse non verifiable.' }
    foreach ($v in $res.value) {
        if ($v.HasException -or $v.ErrorMessage) { throw "Champ $($v.FieldName) refuse : $($v.ErrorMessage)" }
    }
    # Ne plus confondre HTTP 200 et succes des champs. Verifier identites et dates.
    $actual=Invoke-Api $script:TargetContext ($base+'?$select=AuthorId,EditorId,Created,Modified,OData__UIVersionString')
    foreach ($p in $Plan.Expected.PSObject.Properties) {
        if ($null -eq $p.Value) { continue }
        if ([int]$actual.($p.Name) -ne [int]$p.Value) { throw "Identite $($p.Name) non restauree." }
    }
    return $actual
}
function Get-Acl {
    param($Context,[string]$Base)
    $head=Invoke-Api $Context ($Base+'?$select=HasUniqueRoleAssignments')
    $roles=@(Get-Paged $Context ($Base+'/roleassignments?$expand=Member,RoleDefinitionBindings'))
    return [pscustomobject]@{Unique=[bool]$head.HasUniqueRoleAssignments;Assignments=$roles}
}
function Get-RoleMapping {
    param($SourceRole)
    # Comparaison des masques : les libelles peuvent etre localises, les IDs inter-tenant ne sont pas portables.
    $src=$script:SourceContext.Roles[[string]$SourceRole.Id]
    if (-not $src) { throw "Role source introuvable : $($SourceRole.Id)" }
    foreach ($dst in $script:TargetContext.Roles.Values) {
        if ([string]$src.BasePermissions.High -eq [string]$dst.BasePermissions.High -and [string]$src.BasePermissions.Low -eq [string]$dst.BasePermissions.Low) { return [int]$dst.Id }
    }
    throw "Creer le role cible equivalent a '$($src.Name)' (masque $($src.BasePermissions.High):$($src.BasePermissions.Low))."
}
function Sync-Permissions {
    param($SourceItem,$TargetItem,[switch]$Root)
    if (-not $CopyPermissions) { return }
    $srcBase=if ($Root) { $script:SourceContext.ListApi } else { Get-ItemApi $script:SourceContext $SourceItem }
    $dstBase=if ($Root) { $script:TargetContext.ListApi } else { Get-ItemApi $script:TargetContext $TargetItem }
    $src=Get-Acl $script:SourceContext $srcBase
    $dst=Get-Acl $script:TargetContext $dstBase
    if (-not $src.Unique -and -not $Root) {
        if ($dst.Unique) {
            $null=Invoke-Api $script:TargetContext ($dstBase+'/resetroleinheritance()') -Method POST
            $check=Invoke-Api $script:TargetContext ($dstBase+'?$select=HasUniqueRoleAssignments')
            if ($check.HasUniqueRoleAssignments) { throw 'Heritage cible non retabli.' }
        }
        return
    }
    $desired=@{}
    $hasSharingPrincipal=$false
    # Resoudre TOUS les comptes et niveaux de droits AVANT de toucher aux ACL.
    foreach ($a in $src.Assignments) {
        if ([string]$a.Member.LoginName -match '(?i)SharingLinks\.') { $hasSharingPrincipal=$true; continue }
        if ([int]$a.Member.PrincipalType -eq 32 -or [string]$a.Member.LoginName -match '(?i)^i:0i\.t\|ms\.sp\.ext\|') {
            Write-Log 'WARN' "Principal externe non mappable ignore dans ACL source: $($a.Member.LoginName) (type $($a.Member.PrincipalType))."
            continue
        }
        $realRoles=@($a.RoleDefinitionBindings | Where-Object { [int]$_.RoleTypeKind -ne 1 }) # Limited Access : gere par SharePoint
        if (-not $realRoles.Count) { continue }
        $principal=Resolve-Principal $a.Member
        foreach ($r in $realRoles) {
            $role=Get-RoleMapping $r
            $desired[([string]$principal.Id+'|'+$role)]=@{Principal=[int]$principal.Id;Role=$role}
        }
    }
    if ($hasSharingPrincipal -and -not $RecreateSharingLinks) { throw 'Liens de partage source : activer RecreateSharingLinks ou traiter ces liens avant migration des ACL.' }
    if (-not $dst.Unique) {
        $null=Invoke-Api $script:TargetContext ($dstBase+'/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)') -Method POST
    }
    $current=@{}
    foreach ($a in $dst.Assignments) {
        if ([int]$a.Member.PrincipalType -eq 32 -or [string]$a.Member.LoginName -match '(?i)^i:0i\.t\|ms\.sp\.ext\|') { continue }
        foreach ($r in $a.RoleDefinitionBindings) {
            if ([int]$r.RoleTypeKind -eq 1) { continue }
            if ([string]$a.Member.LoginName -match '(?i)SharingLinks\.') { continue } # les liens sont geres separement
            $current[([string]$a.Member.Id+'|'+$r.Id)]=@{Principal=[int]$a.Member.Id;Role=[int]$r.Id}
        }
    }
    foreach ($key in $desired.Keys) {
        if (-not $current.ContainsKey($key)) {
            $p=$desired[$key]
            $null=Invoke-Api $script:TargetContext ($dstBase+'/roleassignments/addroleassignment(principalid='+$p.Principal+',roledefid='+$p.Role+')') -Method POST
        }
    }
    foreach ($key in $current.Keys) {
        if (-not $desired.ContainsKey($key)) {
            $p=$current[$key]
            $null=Invoke-Api $script:TargetContext ($dstBase+'/roleassignments/removeroleassignment(principalid='+$p.Principal+',roledefid='+$p.Role+')') -Method POST
        }
    }
    $verified=Get-Acl $script:TargetContext $dstBase
    $seen=@{}
    foreach ($a in $verified.Assignments) {
        if ([string]$a.Member.LoginName -match '(?i)SharingLinks\.') { continue }
        if ([int]$a.Member.PrincipalType -eq 32 -or [string]$a.Member.LoginName -match '(?i)^i:0i\.t\|ms\.sp\.ext\|') { continue }
        foreach ($r in $a.RoleDefinitionBindings) { if ([int]$r.RoleTypeKind -ne 1) { $seen[([string]$a.Member.Id+'|'+$r.Id)]=$true } }
    }
    foreach ($k in $desired.Keys) { if (-not $seen.ContainsKey($k)) { throw 'Verification ACL : droit cible manquant.' } }
    foreach ($k in $seen.Keys) { if (-not $desired.ContainsKey($k)) { throw 'Verification ACL : droit cible supplementaire.' } }
}
function Sync-Links {
    param($SourceItem,$TargetItem,$State,[string]$StatePath,[string]$Path)
    if (-not $CopyPermissions) { return }
    if ($Path -ne '/') {
        $security=Invoke-Api $script:SourceContext ((Get-ItemApi $script:SourceContext $SourceItem)+'?$select=HasUniqueRoleAssignments')
        if (-not $security.HasUniqueRoleAssignments) {
            # Graph Business n'expose pas toujours inheritedFrom : SharePoint fait autorite.
            # Sync-Permissions vient de retablir l'heritage cible si necessaire.
            $State.Links=@(); Save-State $StatePath $State
            return
        }
    }
    $srcUri=$script:Graph+'/drives/'+$script:SourceContext.DriveId+'/items/'+$SourceItem.id+'/permissions'
    $links=@(Get-Paged $script:SourceContext $srcUri | Where-Object { $_.link -and -not $_.inheritedFrom })
    $destinationLinks=@(Get-Paged $script:TargetContext ($script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$TargetItem.id+'/permissions') | Where-Object { $_.link -and -not $_.inheritedFrom })
    foreach ($link in $destinationLinks) {
        if (-not @($State.Links | Where-Object { $_.DestinationId -eq [string]$link.id }).Count) { throw 'Lien cible preexistant sans correspondance : ne pas annoncer des permissions identiques.' }
    }
    $kinds=@{}
    foreach ($link in $links) {
        $kind=[string]$link.link.scope+'|'+[string]$link.link.type
        if ($kinds.ContainsKey($kind)) { throw 'Plusieurs liens source de meme portee/type : Graph createLink ne garantit pas leur reproduction distincte.' }
        $kinds[$kind]=$true
    }
    if ($links.Count -and -not $RecreateSharingLinks) { throw 'Liens de partage non recrees : RecreateSharingLinks=false.' }
    if (-not $RecreateSharingLinks) { return }
    $present=@{}
    foreach ($p in $links) {
        $present[[string]$p.id]=$true
        if ($p.link.scope -notin @('anonymous','organization') -or $p.link.type -notin @('view','edit')) {
            Write-Log 'WARN' "Lien $($p.id) ignore: portee/type non pris en charge ($($p.link.scope)/$($p.link.type))."
            continue
        }
        if ($p.hasPassword -or $p.link.preventsDownload -or $p.link.type -eq 'blocksDownload') { throw 'Lien protege : ne pas le recreer avec des protections affaiblies.' }
        $expired=$p.expirationDateTime -and ([datetime]$p.expirationDateTime).ToUniversalTime() -le [datetime]::UtcNow
        $stamp=Get-Key ("$($p.link.type)|$($p.link.scope)|$($p.expirationDateTime)")
        $old=@($State.Links | Where-Object { $_.SourceId -eq [string]$p.id })
        if ($old.Count -gt 1) { throw 'Etat des liens incoherent.' }
        $destBase=$script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$TargetItem.id
        if ($old.Count -and ($old[0].Stamp -ne $stamp -or $expired)) {
            $null=Invoke-Api $script:TargetContext ($destBase+'/permissions/'+$old[0].DestinationId) -Method DELETE -AllowNotFound
            $State.Links=@($State.Links | Where-Object { $_.SourceId -ne [string]$p.id })
            Save-State $StatePath $State
            $old=@()
        }
        if ($expired) { continue }
        if ($old.Count) {
            $actual=Invoke-Api $script:TargetContext ($destBase+'/permissions/'+$old[0].DestinationId) -AllowNotFound
            if ($actual -and $actual.link.type -eq $p.link.type -and $actual.link.scope -eq $p.link.scope -and [string]$actual.expirationDateTime -eq [string]$p.expirationDateTime) { continue }
            throw 'Lien cible modifie ou supprime : intervention requise avant recreation.'
        }
        $body=@{type=[string]$p.link.type;scope=[string]$p.link.scope;retainInheritedPermissions=$true}
        if ($p.expirationDateTime) { $body.expirationDateTime=[string]$p.expirationDateTime }
        $new=Invoke-Api $script:TargetContext ($destBase+'/createLink') -Method POST -Body $body
        if (-not $new.id -or $new.link.scope -ne $p.link.scope -or $new.link.type -ne $p.link.type) { throw 'Lien recree non conforme.' }
        if ($p.expirationDateTime -and (-not $new.expirationDateTime -or [Math]::Abs((([datetime]$new.expirationDateTime)-([datetime]$p.expirationDateTime)).TotalSeconds) -gt 1)) {
            $null=Invoke-Api $script:TargetContext ($destBase+'/permissions/'+$new.id) -Method DELETE
            throw 'Expiration non respectee : lien cible retire.'
        }
        $State.Links=@($State.Links)+[pscustomobject]@{SourceId=[string]$p.id;DestinationId=[string]$new.id;Stamp=$stamp;Url=[string]$new.link.webUrl}
        Save-State $StatePath $State
        Write-CsvRow $script:AuditWriter @($script:Pair,$Path,'LINK',$p.id,$new.id,$p.link.webUrl,$new.link.webUrl)
    }
    foreach ($old in @($State.Links)) {
        if (-not $present.ContainsKey([string]$old.SourceId)) {
            $base=$script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$TargetItem.id+'/permissions/'+$old.DestinationId
            $null=Invoke-Api $script:TargetContext $base -Method DELETE -AllowNotFound
            $State.Links=@($State.Links | Where-Object { $_.SourceId -ne $old.SourceId })
            Save-State $StatePath $State
        }
    }
}
function Send-File {
    param([string]$LocalPath,[string]$Path,$Existing)
    $ctx=$script:TargetContext
    $total=([IO.FileInfo]$LocalPath).Length
    $base=$script:Graph+'/drives/'+$ctx.DriveId
    $head=@{}
    if ($Existing) { $head['If-Match']=[string]$Existing.eTag }
    if ($total -le $SmallFileThresholdMB*1048576) {
        $uri=if ($Existing) { $base+'/items/'+$Existing.id+'/content' } else { $base+'/root:/'+(Encode-Path $Path)+':/content?@microsoft.graph.conflictBehavior=fail' }
        $factory={
            $stream=[IO.File]::OpenRead($LocalPath)
            $content=[Net.Http.StreamContent]::new($stream,81920)
            $content.Headers.ContentType=[Net.Http.Headers.MediaTypeHeaderValue]::new('application/octet-stream')
            return $content
        }
        $result=Invoke-Api $ctx $uri -Method PUT -Headers $head -ContentFactory $factory -NoWriteRetry
        if (-not $result.id -or [long]$result.size -ne $total) { throw 'PUT termine sans confirmation de taille/ID.' }
        return $result
    }
    $create=if ($Existing) { $base+'/items/'+$Existing.id+'/createUploadSession' } else { $base+'/root:/'+(Encode-Path $Path)+':/createUploadSession' }
    $behavior=if ($Existing) { 'replace' } else { 'fail' }
    $session=Invoke-Api $ctx $create -Method POST -Headers $head -Body @{item=@{'@microsoft.graph.conflictBehavior'=$behavior}}
    if (-not $session.uploadUrl) { throw 'Session de transfert absente.' }
    $chunkBytes=[int]($ChunkSizeMB*1048576)
    $buffer=[byte[]]::new($chunkBytes)
    $fs=[IO.File]::OpenRead($LocalPath)
    $pos=0L; $failures=0
    try {
        while ($pos -lt $total) {
            $fs.Position=$pos
            $wanted=[int][Math]::Min($chunkBytes,$total-$pos)
            $read=0
            while ($read -lt $wanted) {
                $n=$fs.Read($buffer,$read,$wanted-$read)
                if ($n -le 0) { throw 'Fin de fichier temporaire prematuree.' }
                $read+=$n
            }
            $factory={
                $content=[Net.Http.ByteArrayContent]::new($buffer,0,$read)
                $content.Headers.ContentType=[Net.Http.Headers.MediaTypeHeaderValue]::new('application/octet-stream')
                [void]$content.Headers.TryAddWithoutValidation('Content-Range',('bytes {0}-{1}/{2}' -f $pos,($pos+$read-1),$total))
                return $content
            }
            try {
                $response=Invoke-Api $null $session.uploadUrl -Method PUT -ContentFactory $factory -Anonymous -NoWriteRetry
                if ($response.id) {
                    if ([long]$response.size -ne $total) { throw 'Taille finale incorrecte.' }
                    return $response
                }
                if (-not $response.nextExpectedRanges) { throw 'Session sans nextExpectedRanges : transfert non confirme.' }
                $next=[long](($response.nextExpectedRanges[0] -split '-')[0])
                if ($next -ne $pos+$read) { throw 'Progression inattendue de la session.' }
                $pos=$next; $failures=0
            } catch {
                $failures++
                if ($failures -ge $MaxRetries) { throw }
                # Requete de statut : si le serveur a recu le bloc, ne pas le renvoyer.
                $status=Invoke-Api $null $session.uploadUrl -Anonymous -AllowNotFound
                if (-not $status -or -not $status.nextExpectedRanges) { throw 'Etat upload inconnu apres erreur. Reprise manuelle requise pour eviter un doublon de version.' }
                $next=[long](($status.nextExpectedRanges[0] -split '-')[0])
                if ($next -lt 0 -or $next -ge $total -or ($next % 327680) -ne 0) { throw 'Offset de reprise serveur invalide.' }
                $pos=$next
                Wait-Retry (Get-RetrySeconds $null $failures)
            }
        }
        throw 'Upload incomplet : aucune confirmation finale 200/201 avec driveItem.'
    } finally { $fs.Dispose() }
}
function Copy-VersionContent {
    param($SourceItem,[string]$VersionId,[long]$ExpectedSize,[string]$Path,$Existing,[switch]$Current,$State,[string]$StatePath)
    $tmp=Join-Path $TempTransferDir ($script:RunId+'_'+[guid]::NewGuid().ToString('N')+'.bin')
    try {
        $base=$script:Graph+'/drives/'+$script:SourceContext.DriveId+'/items/'+$SourceItem.id
        $url=if ($Current) { $base+'/content' } else { $base+'/versions/'+[uri]::EscapeDataString($VersionId)+'/content' }
        Invoke-Api $script:SourceContext $url -OutFile $tmp
        if (([IO.FileInfo]$tmp).Length -ne $ExpectedSize) { throw 'Taille telechargee differente de la taille source attendue.' }
        if ($Current) {
            $fresh=Get-ItemMeta $script:SourceContext $SourceItem.id
            if ($fresh.eTag -ne $SourceItem.eTag) { throw 'Source modifiee pendant le telechargement. Relancer apres stabilisation.' }
        }
        $State.Pending.Phase='Uploading'; Save-State $StatePath $State
        $result=Send-File $tmp $Path $Existing
        $script:Stats.Bytes+=$ExpectedSize
        $script:Stats.UploadedVersions++
        return $result
    } finally { if ([IO.File]::Exists($tmp)) { [IO.File]::Delete($tmp) } }
}
function New-ItemState {
    param($Item,[string]$Path)
    return [pscustomobject]@{Schema=2;SourceId=[string]$Item.id;DestinationId='';Path=$Path;
    SourceCTag='';SourceETag='';DestinationETag='';DestinationContentStamp='';Versions=@();Pending=$null;Links=@();
        MetadataStamp='';LastMetadataRun='';Complete=$false;LastContentRun='';LastSeenRun=''}
}
function Get-StatePath {
    param([string]$SourceId)
    $key=Get-Key $SourceId
    $bucket=Join-Path $script:PairDir $key.Substring(0,2)
    [void][IO.Directory]::CreateDirectory($bucket)
    return Join-Path $bucket ($key+'.json')
}
function Assert-TargetStable {
    param($State,$Target)
    if ($State.DestinationId -and (-not $Target -or $Target.id -ne $State.DestinationId)) { throw 'Fichier cible supprime/remplace depuis le checkpoint. Ne pas reutiliser cet etat sur une autre destination.' }
    if (-not $Target) { return }
    $currentStamp = Get-ContentStamp $Target
    if ($State.DestinationContentStamp) {
        if ($currentStamp -and $currentStamp -ne $State.DestinationContentStamp) { throw 'Cible modifiee depuis le checkpoint : conflit a examiner, aucun ecrasement automatique.' }
    } elseif ($State.DestinationETag -and $Target.eTag -ne $State.DestinationETag) {
        throw 'Cible modifiee depuis le checkpoint : conflit a examiner, aucun ecrasement automatique.'
    }
}
function Complete-PendingVersion {
    param($State,[string]$StatePath,$Target)
    $pending=$State.Pending
    if (-not $pending) { return $Target }
    if ($pending.Phase -eq 'Prepared') { $State.Pending=$null; Save-State $StatePath $State; return $Target }
    if ($pending.Phase -eq 'Uploading') {
        if ($Target -and $pending.Size -and [long]$Target.size -eq [long]$pending.Size) {
            Write-Log 'WARN' "$($State.Path) : reprise d'un upload deja present en cible, promotion du checkpoint."
            $pending.Phase='Uploaded'
        } else {
            throw 'Transfert interrompu a issue inconnue. Verifier la derniere version cible avant de reparer le checkpoint (voir notice).'
        }
    }
    Assert-TargetStable $State $Target
    try { $applied=Set-Metadata $Target $pending.Plan }
    finally {
        $fresh=Get-ItemMeta $script:TargetContext $Target.id
        if ($Target.cTag -and $fresh.cTag -ne $Target.cTag) { throw 'Contenu cible modifie pendant la restauration de metadonnees.' }
        $State.DestinationETag=[string]$fresh.eTag
        $State.DestinationContentStamp=Get-ContentStamp $fresh
        Save-State $StatePath $State
    }
    $label=''
    if ($applied) { $label=[string]$applied.OData__UIVersionString }
    if (-not $label) {
        $listItem=Invoke-Api $script:TargetContext ((Get-ItemApi $script:TargetContext $fresh)+'?$select=OData__UIVersionString')
        $label=[string]$listItem.OData__UIVersionString
    }
    if ($CopyVersionHistory -and -not $label) { throw 'Numero de version cible non verifiable.' }
    # Une restauration de metadonnees ne doit pas ecraser la version historique precedente.
    foreach ($record in $State.Versions) {
        if ($CopyVersionHistory -and $record.TargetVersion -eq $label) { throw 'SharePoint a reutilise la meme version cible : historique non preserve. Ne pas poursuivre.' }
    }
    $State.Versions=@($State.Versions | Where-Object { $_.SourceVersion -ne $pending.SourceVersion }) + [pscustomobject]@{
        SourceVersion=$pending.SourceVersion;Key=$pending.Key;TargetVersion=$label;Size=$pending.Size;Modified=$pending.Modified;SourceContentStamp=$pending.SourceContentStamp
    }
    $State.DestinationETag=[string]$fresh.eTag
    if ($pending.IsCurrent) {
        $State.MetadataStamp=Get-Key (ConvertTo-Json -InputObject $pending.Plan -Depth 30 -Compress)
        $State.LastMetadataRun=$script:RunId
    }
    $State.Pending=$null
    Save-State $StatePath $State
    Write-CsvRow $script:AuditWriter @($script:Pair,$State.Path,'VERSION',$pending.SourceVersion,$label,$pending.Modified,$pending.Size)
    return $fresh
}
function Sync-File {
    param($Source,[string]$Path)
    $script:Stats.Files++
    $statePath=Get-StatePath $Source.id
    $state=Read-State $statePath
    if (-not $state) { $state=New-ItemState $Source $Path }
    if ($state.Schema -ne 2 -or $state.SourceId -ne $Source.id) { throw 'Format de checkpoint incompatible.' }
    $target=if ($state.DestinationId) { Get-ItemMeta $script:TargetContext $state.DestinationId } else { Get-TargetByPath $Path }
    if ($target -and $target.folder) { throw 'Conflit : dossier cible au lieu du fichier.' }
    if ($target -and $state.DestinationId -and -not $state.DestinationContentStamp) {
        $state.DestinationContentStamp=Get-ContentStamp $target
        Save-State $statePath $state
    }
    $stamp=Get-ContentStamp $Source
    if ($state.Complete -and $state.DestinationId -and $target -and $target.id -eq $state.DestinationId -and $state.SourceCTag -eq $stamp -and $state.SourceETag -eq $Source.eTag) {
        $script:Stats.SkippedContent++
        Write-Log 'INFO' "$Path : deja synchronise, reprise sans action."
        return
    }
    Assert-TargetStable $state $target
    if ($target -and -not $state.DestinationId) {
        if (-not $AllowAppendToUntrackedDestination) {
            Write-Log 'WARN' "$Path : fichier cible preexistant sans checkpoint v2, adoption temporaire pour reprise."
        }
        $state.DestinationId=[string]$target.id; $state.DestinationETag=[string]$target.eTag; $state.DestinationContentStamp=Get-ContentStamp $target
        Save-State $statePath $state
    }
    if ($state.Path -ne $Path -and $target) {
        $conflict=Get-TargetByPath $Path
        if ($conflict -and $conflict.id -ne $target.id) { throw 'Renommage/deplacement en conflit avec un autre element cible.' }
        $parentPath=($Path -split '/' | Select-Object -SkipLast 1) -join '/'
        $parent=if ($parentPath) { Get-TargetByPath $parentPath } else { $script:TargetContext.Root }
        if (-not $parent) { throw 'Parent cible absent pour le deplacement.' }
        $body=@{name=($Path -split '/')[-1];parentReference=@{id=$parent.id}}
        $target=Invoke-Api $script:TargetContext ($script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$target.id) -Method PATCH -Body $body -Headers @{'If-Match'=$target.eTag}
        $target=Get-ItemMeta $script:TargetContext $target.id
        $state.Path=$Path; $state.DestinationETag=[string]$target.eTag; Save-State $statePath $state
    }
    if ($state.Pending) { $target=Complete-PendingVersion $state $statePath $target }
    $wasComplete=[bool]$state.Complete
    $state.Complete=$false
    Save-State $statePath $state
    $contentChanged=($state.SourceCTag -ne $stamp -or -not $target)
    $inspectVersions=$CopyVersionHistory -and ($CopyMode -eq 'full' -or $contentChanged -or $state.SourceETag -ne $Source.eTag -or -not $wasComplete)
    $versions=@()
    if ($inspectVersions) {
        # L'API livre les versions de la plus recente a la plus ancienne : inverser, ne pas trier par date modifiable.
        $versions=@(Get-Paged $script:SourceContext ($script:Graph+'/drives/'+$script:SourceContext.DriveId+'/items/'+$Source.id+'/versions'))
        [array]::Reverse($versions)
        if (-not $versions.Count) { throw 'Historique demande mais aucune version source disponible.' }
        if ($target -and @($state.Versions).Count) {
            $destVersions=@{}
            Get-Paged $script:TargetContext ($script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$target.id+'/versions') | ForEach-Object { $destVersions[[string]$_.id]=$_ }
            foreach ($record in $state.Versions) {
                if (-not $destVersions.ContainsKey([string]$record.TargetVersion)) { throw 'Version cible du checkpoint disparue (retention ou suppression) : historique incomplet.' }
                if ([long]$destVersions[[string]$record.TargetVersion].size -ne [long]$record.Size) { throw 'Taille de version cible incompatible avec le checkpoint.' }
            }
        }
    } elseif (-not $CopyVersionHistory -and ($contentChanged -or $CopyMode -eq 'full')) {
        $versions=@([pscustomobject]@{id='current';lastModifiedDateTime=$Source.lastModifiedDateTime;size=$Source.size})
    }
    if ($CopyVersionHistory -and @($state.Versions | Where-Object { $_.SourceVersion -eq 'current' }).Count) { throw 'Etat issu de CopyVersionHistory=false : historique complet a reconstruire dans une destination distincte.' }
    $known=@{}
    foreach ($record in $state.Versions) { $known[[string]$record.SourceVersion]=$record }
    if ($CopyVersionHistory -and $versions.Count) {
        $missing=0
        foreach ($v in $versions) { if (-not $known.ContainsKey([string]$v.id) -or $known[[string]$v.id].Key -ne (Get-VersionStamp $v) -or ([string]$v.id -eq [string]$versions[-1].id -and $contentChanged -and $known[[string]$v.id].SourceContentStamp -ne $stamp)) { $missing++ } }
        $limit=[int]$script:TargetContext.List.MajorVersionLimit
        $existingCount=if ($target) { @(Get-Paged $script:TargetContext ($script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$target.id+'/versions')).Count } else { 0 }
        if ($limit -gt 0 -and ($existingCount+$missing) -gt $limit) { throw 'Limite de versions cible insuffisante : augmenter la retention avant le transfert.' }
    }
    $count=0
    foreach ($version in $versions) {
        $count++
        $key=if ($CopyVersionHistory) { Get-VersionStamp $version } else { Get-Key $stamp }
        $versionId=[string]$version.id
        $isCurrent=($count -eq $versions.Count)
        $matchesContent=(-not $isCurrent -or -not $contentChanged -or ($known.ContainsKey($versionId) -and $known[$versionId].SourceContentStamp -eq $stamp))
        if ($known.ContainsKey($versionId) -and $known[$versionId].Key -eq $key -and $matchesContent) { continue }
        if ($known.ContainsKey($versionId) -and $count -ne $versions.Count) { throw "Version historique $versionId modifiee depuis le checkpoint : reconstitution non automatique." }
        $isCurrent=($count -eq $versions.Count)
        $fields=if ($isCurrent -or -not $CopyHistoricalMetadata) { Get-SourceFields $Source } else { Get-SourceFields $Source $versionId }
        if (-not $isCurrent -and -not $CopyHistoricalMetadata -and ($RestoreItemAuthorEditor -or $PreserveFileTimestamps)) {
            # Meme sans champs metier historiques, les auteurs/dates exigent le snapshot de version.
            $fields=Get-SourceFields $Source $versionId
        }
        $sourceFs=if ($isCurrent) { $Source.fileSystemInfo } else { $null }
        $plan=New-MetadataPlan $fields -Historical:(!$isCurrent) -FileSystemInfo $sourceFs -ItemPath $Path
        $state.Complete=$false
        $state.Pending=[pscustomobject]@{Phase='Prepared';SourceVersion=$versionId;Key=$key;Size=[long]$version.size;Modified=[string]$version.lastModifiedDateTime;Plan=$plan;IsCurrent=$isCurrent;SourceContentStamp=$(if ($isCurrent) { $stamp } else { '' })}
        Save-State $statePath $state
        Write-Log 'INFO' "$Path : transfert version $versionId ($($version.size) octets)."
        $target=Copy-VersionContent $Source $versionId ([long]$version.size) $Path $target -Current:$isCurrent -State $state -StatePath $statePath
        $state.DestinationId=[string]$target.id; $state.DestinationETag=[string]$target.eTag; $state.DestinationContentStamp=Get-ContentStamp $target
        $state.Pending.Phase='Uploaded'; Save-State $statePath $state
        $target=Get-ItemMeta $script:TargetContext $target.id
        $target=Complete-PendingVersion $state $statePath $target
    }
    if (-not $target) { throw 'Fichier cible absent apres traitement.' }
    # Verifier la stabilite de l'ensemble source, pas seulement les octets de la derniere version.
    $freshSource=Get-ItemMeta $script:SourceContext $Source.id
    if ($freshSource.eTag -ne $Source.eTag) { throw 'Source modifiee pendant la migration : passe non validee.' }
    $state.SourceCTag=$stamp; $state.SourceETag=[string]$Source.eTag
    $state.LastContentRun=$script:RunId
    Save-State $statePath $state
    # Les metadonnees/ACL peuvent changer sans nouveau contenu.
    $plan=New-MetadataPlan (Get-SourceFields $Source) -FileSystemInfo $Source.fileSystemInfo -ItemPath $Path
    $metaStamp=Get-Key (ConvertTo-Json -InputObject $plan -Depth 30 -Compress)
    try {
        if ($state.MetadataStamp -ne $metaStamp -or ($CopyMode -eq 'full' -and $state.LastMetadataRun -ne $script:RunId)) {
            $null=Set-Metadata $target $plan
            $state.MetadataStamp=$metaStamp; $state.LastMetadataRun=$script:RunId
        }
        Sync-Permissions $Source $target
        Sync-Links $Source $target $state $statePath $Path
    } finally {
        $freshTarget=Get-ItemMeta $script:TargetContext $target.id
        if ($target.cTag -and $freshTarget.cTag -ne $target.cTag) { throw 'Contenu cible modifie pendant les operations de metadonnees/ACL.' }
        $target=$freshTarget
        $state.DestinationETag=[string]$target.eTag
        $state.DestinationContentStamp=Get-ContentStamp $target
        Save-State $statePath $state
    }
    $sourceEnd=Get-ItemMeta $script:SourceContext $Source.id
    if ($sourceEnd.eTag -ne $Source.eTag) { throw 'Source modifiee pendant les metadonnees/permissions : relancer.' }
    $state.DestinationETag=[string]$target.eTag; $state.DestinationContentStamp=Get-ContentStamp $target; $state.Complete=$true; $state.LastSeenRun=$script:RunId
    Save-State $statePath $state
    if (-not $contentChanged) { $script:Stats.SkippedContent++ }
    Write-Log 'OK' "$Path : contenu, metadonnees et permissions traites."
}
function Ensure-Folder {
    param($Source,[string]$Path,[string]$ParentTargetId)
    $statePath=Get-StatePath $Source.id
    $state=Read-State $statePath
    if (-not $state) { $state=New-ItemState $Source $Path }
    $target=if ($state.DestinationId) { Get-ItemMeta $script:TargetContext $state.DestinationId } else { Get-TargetByPath $Path }
    if ($target -and -not $target.folder) { throw 'Conflit : fichier cible au lieu du dossier.' }
    # Les eTags de dossier evoluent avec les enfants : ne pas les traiter comme des conflits de contenu.
    if ($state.DestinationId -and (-not $target -or $target.id -ne $state.DestinationId)) { throw 'Dossier cible du checkpoint absent.' }
    if (-not $target) {
        $uri=$script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$ParentTargetId+'/children'
        $target=Invoke-Api $script:TargetContext $uri -Method POST -Body @{name=$Source.name;folder=@{};'@microsoft.graph.conflictBehavior'='fail'}
    } elseif ($state.Path -ne $Path) {
        $conflict=Get-TargetByPath $Path
        if ($conflict -and $conflict.id -ne $target.id) { throw 'Conflit de deplacement de dossier.' }
        $uri=$script:Graph+'/drives/'+$script:TargetContext.DriveId+'/items/'+$target.id
        $target=Invoke-Api $script:TargetContext $uri -Method PATCH -Headers @{'If-Match'=$target.eTag} -Body @{name=$Source.name;parentReference=@{id=$ParentTargetId}}
    }
    $target=Get-ItemMeta $script:TargetContext $target.id
    $state.DestinationId=[string]$target.id; $state.Path=$Path; $state.Complete=$false
    Save-State $statePath $state
    Sync-Permissions $Source $target
    Sync-Links $Source $target $state $statePath $Path
    $script:Stats.Folders++
    return $target
}
function Invoke-Pair {
    param($Row)
    $script:SourceUpn=$Row.upn_source.Trim(); $script:TargetUpn=$Row.upn_cible.Trim()
    $script:Pair=$script:SourceUpn+' -> '+$script:TargetUpn
    Write-Log 'INFO' "$($script:Pair) : initialisation ($CopyMode)."
    try {
        try {
            $script:SourceContext=New-Context $app_source $script:SourceUpn $Row.onedrive_url_source.Trim()
        } catch {
            throw "Echec contexte SOURCE (tenant=$($app_source.tenant_id), upn=$($script:SourceUpn)) : $($_.Exception.Message)"
        }
        try {
            $script:TargetContext=New-Context $app_cible $script:TargetUpn $Row.onedrive_url_cible.Trim()
        } catch {
            throw "Echec contexte CIBLE (tenant=$($app_cible.tenant_id), upn=$($script:TargetUpn)) : $($_.Exception.Message)"
        }
        if ($script:SourceContext.DriveId -eq $script:TargetContext.DriveId) { throw 'Source et destination identiques.' }
        $pairKey=Get-Key ($app_source.tenant_id+'|'+$script:SourceContext.DriveId+'|'+$app_cible.tenant_id+'|'+$script:TargetContext.DriveId)
        $script:PairDir=Join-Path $StateDirectory $pairKey
        [void][IO.Directory]::CreateDirectory($script:PairDir)
        Test-MetadataSchema
        if ($CopyVersionHistory) {
            if (-not $script:TargetContext.List.EnableVersioning) { throw 'Activer la gestion des versions sur la bibliotheque cible.' }
            if ($script:TargetContext.List.ForceCheckout) { throw 'Extraction obligatoire cible incompatible avec ce moteur.' }
            if ($script:SourceContext.List.EnableMinorVersions -or $script:TargetContext.List.EnableMinorVersions) { throw 'Versions mineures/brouillons : migration native ou outil specialise requis pour une restitution fidele.' }
        }
        # Racine : droits de la bibliotheque. Ni administrateurs de collection, ni groupes eux-memes ne sont modifies.
        Sync-Permissions $script:SourceContext.Root $script:TargetContext.Root -Root
        $rootStatePath=Get-StatePath '__root__'
        $rootState=Read-State $rootStatePath
        if (-not $rootState) { $rootState=New-ItemState $script:SourceContext.Root '/' }
        Sync-Links $script:SourceContext.Root $script:TargetContext.Root $rootState $rootStatePath '/'
        Save-State $rootStatePath $rootState
        # File de dossiers SUR DISQUE : aucune liste globale de fichiers/dossiers en RAM.
        $queuePath=Join-Path $TempTransferDir ($script:RunId+'_'+$pairKey+'.queue.jsonl')
        $folderDir=Join-Path $TempTransferDir ($script:RunId+'_'+$pairKey+'_folders')
        [void][IO.Directory]::CreateDirectory($folderDir)
        $qWriteStream=[IO.File]::Open($queuePath,[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::ReadWrite)
        $qWriter=[IO.StreamWriter]::new($qWriteStream,[Text.UTF8Encoding]::new($false))
        $qReader=$null
        $maxDepth=0
        try {
            $qWriter.WriteLine((@{Id=$script:SourceContext.Root.id;TargetId=$script:TargetContext.Root.id;Path='';Depth=0} | ConvertTo-Json -Compress)); $qWriter.Flush()
            $qReader=[IO.StreamReader]::new([IO.File]::Open($queuePath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite),[Text.Encoding]::UTF8)
            while ($null -ne ($line=$qReader.ReadLine())) {
                $folder=$line | ConvertFrom-Json
                $url=$script:Graph+'/drives/'+$script:SourceContext.DriveId+'/items/'+$folder.Id+'/children?$top='+$PageSize+'&$select=id,name,size,eTag,cTag,file,folder,package,remoteItem,parentReference,sharepointIds,createdDateTime,lastModifiedDateTime,fileSystemInfo'
                while ($url) {
                    # Erreur de pagination : la paire est incomplete, pas de faux succes.
                    $page=Invoke-Api $script:SourceContext $url
                    foreach ($item in $page.value) {
                        $path=if ($folder.Path) { $folder.Path+'/'+$item.name } else { [string]$item.name }
                        try {
                            if ($item.remoteItem -or $item.package -or (-not $item.file -and -not $item.folder)) { throw 'Raccourci, package/OneNote ou type special : transfert classique non pris en charge.' }
                            if (-not $item.sharepointIds.listItemId) { $item=Get-ItemMeta $script:SourceContext $item.id }
                            if ($item.folder) {
                                $target=Ensure-Folder $item $path $folder.TargetId
                                $depth=[int]$folder.Depth+1; $maxDepth=[Math]::Max($maxDepth,$depth)
                                $descriptor=@{Id=$item.id;TargetId=$target.id;Path=$path;Depth=$depth}
                                $json=ConvertTo-Json -InputObject $descriptor -Compress
                                $qWriter.WriteLine($json); $qWriter.Flush()
                                # Une ligne compacte par dossier, jamais de contenu en memoire.
                                [IO.File]::AppendAllText((Join-Path $folderDir ($depth.ToString()+'.jsonl')),$json+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
                            } else { Sync-File $item $path }
                        } catch { Write-Issue $path 'item' $_.Exception.Message }
                    }
                    $url=$page.'@odata.nextLink'
                    $page=$null
                    $mem=[Math]::Round([Diagnostics.Process]::GetCurrentProcess().WorkingSet64/1048576)
                    Write-Log 'INFO' "Avancement : $($script:Stats.Files) fichiers, $($script:Stats.Folders) dossiers, RAM processus ${mem} Mio."
                }
            }
        } finally { if ($qReader) { $qReader.Dispose() }; $qWriter.Dispose() }
        # Restituer les metadonnees de dossier APRES copie des enfants, du plus profond a la racine.
        for ($depth=$maxDepth; $depth -ge 1; $depth--) {
            $path=Join-Path $folderDir ($depth.ToString()+'.jsonl')
            if (-not [IO.File]::Exists($path)) { continue }
            $folderReader=[IO.StreamReader]::new($path,[Text.Encoding]::UTF8)
            try { while ($null -ne ($line=$folderReader.ReadLine())) {
                $f=$line | ConvertFrom-Json
                try {
                    $source=Get-ItemMeta $script:SourceContext $f.Id
                    $target=Get-ItemMeta $script:TargetContext $f.TargetId
                    $plan=New-MetadataPlan (Get-SourceFields $source) -FileSystemInfo $source.fileSystemInfo -IsFolder -ItemPath $f.Path
                    $statePath=Get-StatePath $f.Id; $state=Read-State $statePath
                    $metaStamp=Get-Key (ConvertTo-Json -InputObject $plan -Depth 30 -Compress)
                    # Les changements d'enfants peuvent avoir modifie les dates cible meme si la source est inchangee.
                    $null=Set-Metadata $target $plan
                    $state.MetadataStamp=$metaStamp; $state.SourceETag=[string]$source.eTag; $state.Complete=$true; $state.LastSeenRun=$script:RunId
                    Save-State $statePath $state
                } catch { Write-Issue $f.Path 'folder_metadata' $_.Exception.Message }
            } } finally { $folderReader.Dispose() }
        }
        [IO.File]::Delete($queuePath)
        [IO.Directory]::Delete($folderDir,$true)
    } finally {
        foreach ($ctx in @($script:SourceContext,$script:TargetContext)) { if ($ctx -and $ctx.Cert) { $ctx.Cert.Dispose() } }
        $script:SourceContext=$null; $script:TargetContext=$null
        $script:Tokens.Clear()
    }
}
function Invoke-Main {
    $stopwatch=[Diagnostics.Stopwatch]::StartNew()
    try {
        Normalize-AppConfig $app_source
        Normalize-AppConfig $app_cible
        if ($CopyMode -notin @('full','incremental')) { throw 'CopyMode : full ou incremental uniquement.' }
        if ($ChunkSizeMB*1048576 -le 0 -or ($ChunkSizeMB*1048576)%327680 -ne 0 -or $ChunkSizeMB -ge 60) { throw 'ChunkSizeMB doit donner un multiple de 320 Kio, strictement inferieur a 60 Mio (10 convient).' }
        if ($SmallFileThresholdMB -lt 0 -or $SmallFileThresholdMB -gt 250) { throw 'SmallFileThresholdMB doit etre compris entre 0 et 250.' }
        if ($PageSize -lt 1 -or $PageSize -gt 200 -or $MaxRetries -lt 1 -or $CacheLimit -lt 1 -or $RequestTimeoutMinutes -lt 1) { throw 'Parametres de pagination, cache, timeout ou retry invalides.' }
        foreach ($app in @($app_source,$app_cible)) {
            $parsed=[guid]::Empty
            if (-not [guid]::TryParse($app.tenant_id,[ref]$parsed) -or -not [guid]::TryParse($app.client_id,[ref]$parsed)) { throw 'Renseigner tenant_id et client_id des deux applications.' }
        }
        foreach ($dir in @($StateDirectory,$LogDirectory,$TempTransferDir)) { [void][IO.Directory]::CreateDirectory($dir) }
        # Verrou exclusif : deux processus ne doivent pas partager les memes checkpoints.
        $script:Lock=[IO.File]::Open((Join-Path $StateDirectory 'migration.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $prefix=Join-Path $LogDirectory ('OneDrive_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'_'+$script:RunId.Substring(0,8))
        $utf8=[Text.UTF8Encoding]::new($true)
        $script:LogWriter=[IO.StreamWriter]::new($prefix+'.log',$false,$utf8)
        $script:ReportWriter=[IO.StreamWriter]::new($prefix+'_errors.csv',$false,$utf8)
        $script:AuditWriter=[IO.StreamWriter]::new($prefix+'_versions_links.csv',$false,$utf8)
        Write-CsvRow $script:ReportWriter @('utc','migration','chemin','phase','erreur')
        Write-CsvRow $script:AuditWriter @('migration','chemin','type','source_id','cible_id','source_date_ou_url','taille_ou_url_cible')
        $handler=[Net.Http.HttpClientHandler]::new()
        $script:Http=[Net.Http.HttpClient]::new($handler)
        $script:Http.Timeout=[TimeSpan]::FromMinutes($RequestTimeoutMinutes)
        Load-Mappings
        if (-not (Test-Path -LiteralPath $CsvPath)) { throw 'migrations.csv introuvable.' }
        $rows=0
        Import-Csv -LiteralPath $CsvPath -Delimiter ';' | ForEach-Object {
            $row=$_; $rows++
            foreach ($col in @('onedrive_url_source','onedrive_url_cible','upn_source','upn_cible')) {
                if (-not $row.$col) { throw "Ligne $rows : colonne $col absente/vide." }
            }
            try { Invoke-Pair $row } catch { Write-Issue ('ligne_'+$rows) 'migration' $_.Exception.Message }
        }
        if (-not $rows) { throw 'CSV de migrations vide.' }
        $stopwatch.Stop()
        $result=if ($script:Stats.Errors) { 'PARTIEL/ECHEC' } else { 'TERMINE' }
        Write-Log 'INFO' "$result : $($script:Stats.Files) fichiers examines, $($script:Stats.UploadedVersions) transferts, $($script:Stats.Errors) erreurs, $([Math]::Round($stopwatch.Elapsed.TotalMinutes,2)) min, $($script:Stats.Requests) requetes, $($script:Stats.Retries) retries."
        Write-Log 'INFO' "Rapports : $prefix"
        if ($script:Stats.Errors) { throw 'Migration incomplete. Corriger le rapport erreurs et relancer avec les memes etats.' }
    } finally {
        foreach ($resource in @($script:Http,$script:LogWriter,$script:ReportWriter,$script:AuditWriter,$script:Lock)) { if ($resource) { $resource.Dispose() } }
    }
}
# Le chargement par dot-sourcing ne lance jamais une migration : permet tests unitaires hors ligne.
if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
