# Script PowerShell - Copie SharePoint depuis PostgreSQL (fichiers + droits)
# Version: 1.0

# =========================
# PARAMETRES GRAPH
# =========================
$connexion_source = @{
        tenant_Id = "85a5a352-25d6-4894-a97d-221cd1712dd2"
        client_Id = "684f670e-6931-4120-b053-f8c9538744f6"
        client_Secret = "O5s8Q~Z~e_d1-dSQ.vt4uFTeR8c4Xssaa093Racz"
}

$connexion_cible = @{
    tenant_id     = "2eea08b8-1972-447b-ad43-d044d042500a"
    client_id     = "821109a4-6e0e-48e8-b477-8f9b70aa32a4"
    client_secret = "Yhy8Q~mBxTrYqx~ve7enqTQatjwhEZ5.34-q-a4j"
   # client_secret ="66A8Q~t2-Q_vZp54cDnf8TH~3kl0vp1UyHMZ8bVu"
}


$site_url_source = "https://ipdlab.sharepoint.com/sites/SERVICES_GENERAUX/"
$site_url_cible  = "https://infoprodigital365.sharepoint.com/sites/WH_GPA-test/"

$translation_csv = "E:\UKREiiF_FinalCopy_path.csv"
$mapping_csv     = "E:\UserMapping.csv"
$CsvLibraryPrefixToStrip = ""

$ForceOverwrite = $true
$ChunkSize      = 8MB

# Discovery run a utiliser. Laisser vide pour prendre le dernier run discovery successful.
$DiscoveryRunId = ""

# =========================
# PARAMETRES POSTGRES
# =========================
$pg = @{
    host = "VIT-V-LICD-PGS.info.local"
    port = 5432
    database = "mercury_dev"
    user = "svc_mercury_dev"
    password = "auRCQAgt1W6f_9eTlrJYUMvz"
    sslmode = "Disable"
    psqlPath = "psql"
}

# =========================
# ETAT GLOBAL
# =========================
$script:TokenCache = @{}
$script:CopyRunId = $null
$script:UrlTranslations = @{}
$script:UserMapping = @{}
$script:UserMappingByName = @{}
$script:ResolvedIdCache = @{}
$script:NpgsqlReady = $false
$script:NpgsqlVersion = "8.0.8"
$script:AssemblyResolverReady = $false
$script:AssemblyPathCache = @{}
$script:DbSchema = "mercury"
$script:DbLogEnabled = $true
$script:DbLogWriteInProgress = $false
$script:DbLogFailureWarned = $false

Add-Type -AssemblyName System.Net.Http
$script:HttpClient = [System.Net.Http.HttpClient]::new()
$script:HttpClient.Timeout = [TimeSpan]::FromMinutes(30)

# =========================
# LOG
# =========================
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Phase = "runtime",
        [string]$Path = "/",
        $Context = $null
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$ts][$Level] $Message"

    if (-not $script:DbLogEnabled) { return }
    if (-not $script:CopyRunId) { return }
    if (-not $script:NpgsqlReady) { return }
    if ($script:DbLogWriteInProgress) { return }

    try {
        $script:DbLogWriteInProgress = $true
        $ctxObj = [ordered]@{
            source = "copy"
            phase = $Phase
            path = $Path
        }
        if ($Context) {
            if ($Context -is [hashtable] -or $Context -is [System.Collections.IDictionary]) {
                foreach ($k in $Context.Keys) { $ctxObj[[string]$k] = $Context[$k] }
            } else {
                $ctxObj["detail"] = [string]$Context
            }
        }
        $ctx = ($ctxObj | ConvertTo-Json -Depth 8 -Compress)
        $sql = @"
INSERT INTO sp_logs (run_id, log_level, message, context)
VALUES (
    $(Sql-Literal $script:CopyRunId),
    $(Sql-Literal $Level),
    $(Sql-Literal $Message),
    $(Sql-Literal $ctx)::jsonb
);
"@
        Invoke-PgNonQuery -Sql $sql
    } catch {
        # Ne jamais casser le flux principal pour un echec de log DB.
        if (-not $script:DbLogFailureWarned) {
            $script:DbLogFailureWarned = $true
            Write-Host "[WARN] Echec ecriture log DB (sp_logs): $($_.Exception.Message)"
        }
    } finally {
        $script:DbLogWriteInProgress = $false
    }
}

# =========================
# SQL HELPERS
# =========================
function Sql-Literal([AllowNull()][string]$value) {
    if ($null -eq $value) { return "NULL" }
    return "'" + ($value -replace "'", "''") + "'"
}

function Get-LoaderExceptionText {
    param([System.Exception]$Ex)

    if ($Ex -and $Ex.LoaderExceptions) {
        $msgs = @($Ex.LoaderExceptions | Where-Object { $_ } | ForEach-Object { $_.Message })
        if ($msgs.Count -gt 0) { return ($msgs -join " | ") }
    }
    return $Ex.Message
}

function Register-NuGetAssemblyResolver {
    if ($script:AssemblyResolverReady) { return }
    $script:AssemblyResolverReady = $true
}

function Import-PackageDependency {
    param(
        [Parameter(Mandatory)][string]$PackagesRoot,
        [Parameter(Mandatory)][string]$PackagePrefix,
        [Parameter(Mandatory)][string]$DllName
    )

    if (-not (Test-Path -LiteralPath $PackagesRoot)) { return }

    $pkgDir = Get-ChildItem -LiteralPath $PackagesRoot -Directory -Filter ("{0}*" -f $PackagePrefix) -ErrorAction SilentlyContinue |
        Sort-Object -Property Name -Descending |
        Select-Object -First 1
    if (-not $pkgDir) { return }

    $baseDir = $pkgDir.FullName
    if (-not (Test-Path -LiteralPath (Join-Path -Path $baseDir -ChildPath "lib"))) {
        $major = (($script:NpgsqlVersion -split '\\.')[0])
        $verDir = Get-ChildItem -LiteralPath $baseDir -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d+\.\d+\.\d+' } |
            Sort-Object -Property @{Expression={
                if ($major -and $_.Name -match ("^{0}\\." -f [regex]::Escape($major))) { 1 } else { 0 }
            }}, @{Expression={$_.Name}} -Descending |
            Select-Object -First 1
        if ($verDir) { $baseDir = $verDir.FullName }
    }

    $tfms = @("net48","net472","net471","net47","net462","net461","net46","net451","net45","net40","netstandard2.0")
    foreach ($tfm in $tfms) {
        $dllPath = Join-Path -Path $baseDir -ChildPath ("lib\\{0}\\{1}" -f $tfm, $DllName)
        if (Test-Path -LiteralPath $dllPath) {
            try { [void][System.Reflection.Assembly]::LoadFrom($dllPath); return } catch {}
        }
    }
}

function Import-DependencyByPath {
    param([Parameter(Mandatory)][string]$DllPath)
    if (Test-Path -LiteralPath $DllPath) {
        try {
            [void][System.Reflection.Assembly]::LoadFrom($DllPath)
            Write-Log ("Dependance chargee: {0}" -f $DllPath) "DEBUG"
            return $true
        } catch {
            Write-Log ("Echec chargement dependance {0}: {1}" -f $DllPath, $_.Exception.Message) "WARN"
        }
    }
    return $false
}

function Import-NpgsqlFromRoot {
    param([Parameter(Mandatory)][string]$Root)

    if (-not (Test-Path -LiteralPath $Root)) { return $false }

    Register-NuGetAssemblyResolver
    $preferredTfms = @("net48","net472","net471","net47","net462","net461","net46","net451","net45","net40","netstandard2.0")

    foreach ($tfm in $preferredTfms) {
        $libDir = Join-Path -Path $Root -ChildPath ("lib\\{0}" -f $tfm)
        $npgsqlDll = Join-Path -Path $libDir -ChildPath "Npgsql.dll"
        if (-not (Test-Path -LiteralPath $npgsqlDll)) { continue }

        try {
            $packagesRoot = Split-Path -Parent $Root
            $leaf = Split-Path -Leaf $Root
            if ($leaf -match '^\d+\.\d+\.\d+') {
                $packagesRoot = Split-Path -Parent $packagesRoot
            }
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Extensions.Logging.Abstractions" -DllName "Microsoft.Extensions.Logging.Abstractions.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Bcl.AsyncInterfaces" -DllName "Microsoft.Bcl.AsyncInterfaces.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Extensions.DependencyInjection.Abstractions" -DllName "Microsoft.Extensions.DependencyInjection.Abstractions.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Extensions.Options" -DllName "Microsoft.Extensions.Options.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Extensions.Primitives" -DllName "Microsoft.Extensions.Primitives.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Diagnostics.DiagnosticSource" -DllName "System.Diagnostics.DiagnosticSource.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Text.Json" -DllName "System.Text.Json.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Text.Encodings.Web" -DllName "System.Text.Encodings.Web.dll"
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.threading.tasks.extensions\4.5.4\lib\net461\System.Threading.Tasks.Extensions.dll" | Out-Null
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Threading.Tasks.Extensions" -DllName "System.Threading.Tasks.Extensions.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Runtime.CompilerServices.Unsafe" -DllName "System.Runtime.CompilerServices.Unsafe.dll"
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\Desktop\Mercury script powershell\deps\System.Buffers.4.0.3.0.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.memory\4.5.5\lib\net461\System.Memory.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.runtime.compilerservices.unsafe\6.0.0\lib\netstandard2.0\System.Runtime.CompilerServices.Unsafe.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\Desktop\Mercury script powershell\deps\System.Runtime.CompilerServices.Unsafe.4.0.4.1.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.numerics.vectors\4.5.0\lib\net46\System.Numerics.Vectors.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.text.json\8.0.5\lib\netstandard2.0\System.Text.Json.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.text.encodings.web\8.0.0\lib\net462\System.Text.Encodings.Web.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.threading.channels\8.0.0\lib\net462\System.Threading.Channels.dll" | Out-Null
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Memory" -DllName "System.Memory.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Numerics.Vectors" -DllName "System.Numerics.Vectors.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.ValueTuple" -DllName "System.ValueTuple.dll"

            $depDlls = Get-ChildItem -LiteralPath $libDir -Filter "*.dll" -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ne "Npgsql.dll" }
            foreach ($dep in $depDlls) {
                try { [void][System.Reflection.Assembly]::LoadFrom($dep.FullName) } catch {}
            }

            [void][System.Reflection.Assembly]::LoadFrom($npgsqlDll)
            $null = [Npgsql.NpgsqlConnection]
            Write-Log ("Provider Npgsql charge (TFM {0}): {1}" -f $tfm, $npgsqlDll) "INFO"
            return $true
        } catch [System.Reflection.ReflectionTypeLoadException] {
            Write-Log ("Echec chargement Npgsql depuis {0}: {1}" -f $npgsqlDll, (Get-LoaderExceptionText -Ex $_.Exception)) "WARN"
        } catch {
            Write-Log ("Echec chargement Npgsql depuis {0}: {1}" -f $npgsqlDll, $_.Exception.Message) "WARN"
        }
    }

    return $false
}

function Initialize-Npgsql {
    if ($script:NpgsqlReady) { return }

    try {
        Add-Type -AssemblyName Npgsql -ErrorAction Stop
        $script:NpgsqlReady = $true
        Write-Log "Provider Npgsql charge depuis le systeme" "INFO"
        return
    } catch {}

    $preferredRoot = Join-Path -Path "$env:USERPROFILE\\.nuget\\packages\\npgsql" -ChildPath $script:NpgsqlVersion
    if (Import-NpgsqlFromRoot -Root $preferredRoot) {
        $script:NpgsqlReady = $true
        return
    }

    $candidatePatterns = @(
        "$env:USERPROFILE\\.nuget\\packages\\npgsql",
        "$env:LOCALAPPDATA\\PackageManagement\\NuGet\\Packages\\Npgsql.*",
        "$env:ProgramFiles\\PackageManagement\\NuGet\\Packages\\Npgsql.*"
    )

    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($pattern in $candidatePatterns) {
        $items = @(Get-ChildItem -Path $pattern -Directory -ErrorAction SilentlyContinue)
        foreach ($item in $items) { $roots.Add($item.FullName) | Out-Null }
    }

    foreach ($root in @($roots | Sort-Object -Descending -Unique)) {
        if ($root -ieq $preferredRoot) { continue }
        if (Import-NpgsqlFromRoot -Root $root) {
            $script:NpgsqlReady = $true
            return
        }
    }

    Write-Log "Npgsql introuvable localement, tentative d'installation NuGet..." "WARN"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    } catch {}

    Install-Package -Name Npgsql -RequiredVersion $script:NpgsqlVersion -ProviderName NuGet -Scope CurrentUser -Force -ErrorAction Stop | Out-Null

    if (-not (Import-NpgsqlFromRoot -Root $preferredRoot)) {
        throw "Npgsql installe mais non chargeable. Verifiez .NET Framework 4.7.2+ et les dependances runtime."
    }

    $script:NpgsqlReady = $true
    Write-Log "Provider Npgsql installe et charge avec succes" "SUCCESS"
}

function New-PgConnectionString {
    Initialize-Npgsql

    $builder = New-Object Npgsql.NpgsqlConnectionStringBuilder
    $builder.Host = [string]$pg.host
    $builder.Port = [int]$pg.port
    $builder.Database = [string]$pg.database
    $builder.Username = [string]$pg.user
    $builder.Password = [string]$pg.password

    $modeRaw = if ([string]::IsNullOrWhiteSpace([string]$pg.sslmode)) { "Prefer" } else { [string]$pg.sslmode }
    $modeNormalized = switch -Regex ($modeRaw.Trim().ToLowerInvariant()) {
        "^(disable|disabled|off|none)$" { "Disable"; break }
        "^(allow)$" { "Allow"; break }
        "^(prefer|preferred)$" { "Prefer"; break }
        "^(require|required|on|true)$" { "Require"; break }
        "^(verifyca|verify-ca)$" { "VerifyCA"; break }
        "^(verifyfull|verify-full)$" { "VerifyFull"; break }
        default { "Prefer" }
    }

    try {
        $builder.SslMode = [Npgsql.SslMode]::$modeNormalized
    } catch {
        Write-Log ("SslMode invalide '{0}', fallback sur Disable" -f $modeRaw) "WARN"
        $builder.SslMode = [Npgsql.SslMode]::Disable
    }

    if ($builder.SslMode -ne [Npgsql.SslMode]::Disable) {
        try { $builder.TrustServerCertificate = $true } catch {}
    }

    if ($script:DbSchema) {
        $builder.SearchPath = [string]$script:DbSchema
    }

    return $builder.ConnectionString
}

function Invoke-PgNonQuery {
    param([Parameter(Mandatory)][string]$Sql)

    Initialize-Npgsql
    $conn = New-Object Npgsql.NpgsqlConnection (New-PgConnectionString)
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        [void]$cmd.ExecuteNonQuery()
    } finally {
        $conn.Dispose()
    }
}

function Invoke-PgQuery {
    param([Parameter(Mandatory)][string]$Sql)

    Initialize-Npgsql
    $conn = New-Object Npgsql.NpgsqlConnection (New-PgConnectionString)
    $rows = New-Object System.Collections.Generic.List[object]
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $obj = [ordered]@{}
            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                $v = $reader.GetValue($i)
                if ($v -is [System.DBNull]) { $v = $null }
                $obj[$reader.GetName($i)] = $v
            }
            $rows.Add([pscustomobject]$obj) | Out-Null
        }
        $reader.Close()
    } finally {
        $conn.Dispose()
    }
    return $rows.ToArray()
}

function Invoke-PgScalar {
    param([Parameter(Mandatory)][string]$Sql)

    Initialize-Npgsql
    $conn = New-Object Npgsql.NpgsqlConnection (New-PgConnectionString)
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $v = $cmd.ExecuteScalar()
        if ($v -is [System.DBNull]) { return $null }
        return $v
    } finally {
        $conn.Dispose()
    }
}

function Get-DbWritableSchema {
    $sql = @"
SELECT n.nspname
FROM pg_catalog.pg_namespace n
WHERE pg_catalog.has_schema_privilege(current_user, n.nspname, 'CREATE')
  AND n.nspname NOT LIKE 'pg_%'
  AND n.nspname <> 'information_schema'
ORDER BY CASE WHEN n.nspname = 'public' THEN 0 ELSE 1 END, n.nspname
LIMIT 1;
"@

    $schema = Invoke-PgScalar -Sql $sql
    if ($schema) { return [string]$schema }
    return $null
}

function Get-DbSchemaPrivilegeReport {
    $sql = @"
SELECT n.nspname || ':' ||
       CASE WHEN pg_catalog.has_schema_privilege(current_user, n.nspname, 'USAGE') THEN 'U' ELSE '-' END ||
       CASE WHEN pg_catalog.has_schema_privilege(current_user, n.nspname, 'CREATE') THEN 'C' ELSE '-' END
FROM pg_catalog.pg_namespace n
WHERE n.nspname NOT LIKE 'pg_%'
  AND n.nspname <> 'information_schema'
ORDER BY CASE WHEN n.nspname = 'public' THEN 0 ELSE 1 END, n.nspname;
"@

    $rows = Invoke-PgQuery -Sql $sql
    return @($rows | ForEach-Object { $_.nspname })
}

function Get-DbExistingSchemaWithTables {
    $sql = @"
SELECT n.nspname
FROM pg_catalog.pg_namespace n
WHERE pg_catalog.has_schema_privilege(current_user, n.nspname, 'USAGE')
  AND n.nspname NOT LIKE 'pg_%'
  AND n.nspname <> 'information_schema'
  AND EXISTS (
      SELECT 1
    FROM pg_catalog.pg_class c
    JOIN pg_catalog.pg_namespace ns ON ns.oid = c.relnamespace
      WHERE ns.nspname = n.nspname
        AND c.relname = 'sp_runs'
        AND c.relkind = 'r'
  )
ORDER BY CASE WHEN n.nspname = 'public' THEN 0 ELSE 1 END, n.nspname
LIMIT 1;
"@

    $schema = Invoke-PgScalar -Sql $sql
    if ($schema) { return [string]$schema }
    return $null
}

function Ensure-DatabaseSchema {
    $schemaPath = Join-Path -Path $PSScriptRoot -ChildPath "create_tables_postgres.sql"
    if (-not (Test-Path -LiteralPath $schemaPath)) {
        throw "Fichier schema introuvable: $schemaPath"
    }

    $sql = Get-Content -LiteralPath $schemaPath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($sql)) {
        throw "Fichier schema vide: $schemaPath"
    }

    $targetSchema = $script:DbSchema
    $hasCreate = $false
    try {
        $hasCreate = [bool](Invoke-PgScalar -Sql ("SELECT pg_catalog.has_schema_privilege(current_user, {0}, 'CREATE');" -f (Sql-Literal $targetSchema)))
    } catch {}

    if (-not $hasCreate) {
        $existingSchema = Get-DbExistingSchemaWithTables
        if ($existingSchema) {
            $script:DbSchema = $existingSchema
            Write-Log ("Aucun CREATE sur '{0}'. Utilisation du schema existant '{1}'" -f $targetSchema, $existingSchema) "WARN"
            return
        }

        $writableSchema = Get-DbWritableSchema
        if ($writableSchema) {
            $script:DbSchema = $writableSchema
            $targetSchema = $writableSchema
            Write-Log ("Schema cible '{0}' non accessible. Bascule vers le schema writable '{1}'" -f $targetSchema, $writableSchema) "WARN"
        } else {
            $report = Get-DbSchemaPrivilegeReport
            $details = if ($report -and $report.Count -gt 0) { ($report -join ', ') } else { 'aucun schema visible' }
            throw "Aucun schema PostgreSQL writable disponible pour '$targetSchema'. Schemas/privileges visibles: $details. Demander au DBA: CREATE SCHEMA mercury; GRANT USAGE, CREATE ON SCHEMA mercury TO $($pg.user);"
        }
    }

    $escapedSchema = ($targetSchema -replace '"', '""')
    $sqlWithSearchPath = ("SET search_path TO ""{0}"";`n" -f $escapedSchema) + $sql

    try {
        Invoke-PgNonQuery -Sql $sqlWithSearchPath
        Write-Log ("Schema PostgreSQL verifie/applique dans le schema '{0}'" -f $targetSchema) "SUCCESS"
        return
    } catch {
        $existingSchema = Get-DbExistingSchemaWithTables
        if ($existingSchema) {
            $script:DbSchema = $existingSchema
            Write-Log ("Aucun CREATE sur '{0}'. Utilisation du schema existant '{1}'" -f $targetSchema, $existingSchema) "WARN"
            return
        }
        throw
    }
}

function Start-CopyRun {
    param([string]$DiscoveryRun)

    $newRunId = [guid]::NewGuid().ToString()

    $sql = @"
INSERT INTO sp_runs (run_id, run_type, status, source_site_url, target_site_url, source_tenant_id, target_tenant_id, notes)
VALUES ($(Sql-Literal $newRunId), 'copy', 'running', $(Sql-Literal $site_url_source), $(Sql-Literal $site_url_cible), $(Sql-Literal $connexion_source.tenant_id), $(Sql-Literal $connexion_cible.tenant_id), $(Sql-Literal ("discovery_run=" + $DiscoveryRun)))
RETURNING run_id;
"@
    $rows = Invoke-PgQuery -Sql $sql
    $script:CopyRunId = [string]$rows[0].run_id
    Write-Log "Run copy cree: $($script:CopyRunId)" "SUCCESS"
}

function Finish-CopyRun {
    param([string]$Status)

    $sql = @"
UPDATE sp_runs
SET finished_at = now(),
    status = $(Sql-Literal $Status),
    total_items = (SELECT COUNT(*) FROM sp_items WHERE run_id = $(Sql-Literal $script:CopyRunId)),
    total_permissions = (SELECT COUNT(*) FROM sp_permissions WHERE run_id = $(Sql-Literal $script:CopyRunId)),
    total_errors = (SELECT COUNT(*) FROM sp_errors WHERE run_id = $(Sql-Literal $script:CopyRunId))
WHERE run_id = $(Sql-Literal $script:CopyRunId);
"@
    Invoke-PgNonQuery -Sql $sql
}

function Db-InsertError {
    param([string]$Path, [string]$Phase, [string]$Message)
    if (-not $script:CopyRunId) { return }

    $sql = @"
INSERT INTO sp_errors (run_id, path, phase, error_message, error_level)
VALUES ($(Sql-Literal $script:CopyRunId), $(Sql-Literal $Path), $(Sql-Literal $Phase), $(Sql-Literal $Message), 'ERROR');
"@
    Invoke-PgNonQuery -Sql $sql
}

function Db-UpsertCopyItem {
    param(
        [string]$PathOriginal,
        [string]$PathTranslated,
        [string]$ItemType,
        [bool]$IsFolder,
        [string]$ItemId,
        [string]$Status
    )

    $sql = @"
INSERT INTO sp_items (
    run_id, site_url, drive_id, item_id, path_original, path_translated, item_type, is_folder,
    copy_status, copied_at, created_at, updated_at
)
VALUES (
    $(Sql-Literal $script:CopyRunId),
    $(Sql-Literal $site_url_cible),
    NULL,
    $(Sql-Literal $ItemId),
    $(Sql-Literal $PathOriginal),
    $(Sql-Literal $PathTranslated),
    $(Sql-Literal $ItemType),
    $([int]$IsFolder),
    $(Sql-Literal $Status),
    now(),
    now(),
    now()
)
ON CONFLICT (run_id, path_original)
DO UPDATE SET
    item_id = EXCLUDED.item_id,
    path_translated = EXCLUDED.path_translated,
    item_type = EXCLUDED.item_type,
    is_folder = EXCLUDED.is_folder,
    copy_status = EXCLUDED.copy_status,
    copied_at = now(),
    updated_at = now();
"@
    Invoke-PgNonQuery -Sql $sql
}

function Db-InsertPermissionResult {
    param(
        [string]$PathOriginal,
        [string]$PathTranslated,
        [string]$ItemType,
        [string]$GrantedTo,
        [string]$GrantedToId,
        [string]$TargetType,
        [string]$Role,
        [string]$Status,
        [string]$Phase,
        [string]$ErrorMessage,
        [string]$MappedDestination,
        [string]$MappedDestinationId
    )

    $sql = @"
INSERT INTO sp_permissions (
    run_id, site_url, drive_id, item_id,
    path_original, path_translated, item_type,
    granted_to, granted_to_id, target_type, role_name,
    mapped_destination, mapped_destination_id,
    permission_status, phase, error_message, created_at
)
VALUES (
    $(Sql-Literal $script:CopyRunId),
    $(Sql-Literal $site_url_cible),
    NULL,
    NULL,
    $(Sql-Literal $PathOriginal),
    $(Sql-Literal $PathTranslated),
    $(Sql-Literal $ItemType),
    $(Sql-Literal $GrantedTo),
    $(Sql-Literal $GrantedToId),
    $(Sql-Literal $TargetType),
    $(Sql-Literal $Role),
    $(Sql-Literal $MappedDestination),
    $(Sql-Literal $MappedDestinationId),
    $(Sql-Literal $Status),
    $(Sql-Literal $Phase),
    $(Sql-Literal $ErrorMessage),
    now()
);
"@
    Invoke-PgNonQuery -Sql $sql
}

function Get-DiscoveryRunId {
    if ($DiscoveryRunId) { return $DiscoveryRunId }

    $sql = @"
SELECT run_id
FROM sp_runs
WHERE run_type = 'discovery' AND status IN ('success','warning')
ORDER BY started_at DESC
LIMIT 1
"@
    $rows = Invoke-PgQuery -Sql $sql
    if (-not $rows -or -not $rows[0].run_id) {
        throw "Aucun run discovery valide trouve en base"
    }
    return [string]$rows[0].run_id
}

function Get-WorkRowsFromDb {
    param([string]$RunId)

    $sql = @"
SELECT
    p.path_original AS "Path",
    COALESCE(i.item_type, p.item_type, 'File') AS "ItemType",
    p.granted_to AS "GrantedTo",
    p.granted_to_id AS "grantedToID",
    p.target_type AS "TargetType",
    p.role_name AS "Role"
FROM sp_permissions p
LEFT JOIN sp_items i
    ON i.run_id = p.run_id
   AND i.path_original = p.path_original
WHERE p.run_id = $(Sql-Literal $RunId)
ORDER BY p.path_original;
"@
    $rows = @(Invoke-PgQuery -Sql $sql)
    Write-Log ("Permissions chargees depuis DB: {0}" -f $rows.Count) "INFO" "load" "/" @{ run_id = $RunId }
    return $rows
}

function Load-UserMapping {
    # Priorite DB
    $sql = @"
SELECT
    source_id AS "SourceId",
    source_display_name AS "SourceDisplayName",
    target_type AS "TargetType",
    destination_display_name AS "DestinationDisplayName"
FROM sp_user_mapping
WHERE is_active = true
  AND COALESCE(destination_display_name, '') <> '';
"@

    $rows = @(Invoke-PgQuery -Sql $sql)

    if ($rows.Count -eq 0 -and (Test-Path $mapping_csv)) {
        Write-Log "Mapping DB vide, fallback CSV" "WARN"
        $rows = @(Import-Csv -Path $mapping_csv -Encoding UTF8)
    }

    foreach ($row in $rows) {
        $srcId = if ($row.SourceId) { [string]$row.SourceId } else { "" }
        $srcName = if ($row.SourceDisplayName) { [string]$row.SourceDisplayName } else { "" }
        $dstName = if ($row.DestinationDisplayName) { [string]$row.DestinationDisplayName } else { "" }
        $tt = if ($row.TargetType) { [string]$row.TargetType } else { "" }

        if ([string]::IsNullOrWhiteSpace($dstName)) { continue }

        $entry = @{ DisplayName = $dstName; TargetType = $tt }
        if ($srcId) { $script:UserMapping[$srcId] = $entry }
        if ($srcName) { $script:UserMappingByName[$srcName] = $entry }
    }

    Write-Log ("Mappages charges: ID={0}, Name={1}" -f $script:UserMapping.Count, $script:UserMappingByName.Count) "SUCCESS"
}

# =========================
# GRAPH HELPERS
# =========================
function Normalize-RelPath([string]$rel) {
    if (-not $rel) { return "" }
    return (($rel.Trim() -replace '\\','/').TrimStart('/').TrimEnd('/ '))
}

function Strip-LibraryPrefix([string]$rel) {
    if ([string]::IsNullOrWhiteSpace($CsvLibraryPrefixToStrip)) { return $rel }
    $prefix = $CsvLibraryPrefixToStrip.Trim().TrimEnd('/') + "/"
    if ($rel.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $rel.Substring($prefix.Length)
    }
    if ($rel -ieq $CsvLibraryPrefixToStrip.Trim()) { return "" }
    return $rel
}

function UrlEncode-Segment([string]$seg) {
    if ([string]::IsNullOrWhiteSpace($seg)) { return "" }
    [System.Uri]::EscapeDataString($seg)
}

function Encode-RelForGraph([string]$rel) {
    $segs = ($rel -replace '\\','/') -split '/' | Where-Object { $_ -ne '' }
    return [string]::Join('/', ($segs | ForEach-Object { UrlEncode-Segment $_ }))
}

function Build-AuthHeader([string]$token) { @{ Authorization = "Bearer $token" } }

function Get-AccessTokenResource {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$ResourceAppIdUri
    )

    $cacheKey = "$TenantId|$ClientId|$ResourceAppIdUri"
    $now = Get-Date
    if ($script:TokenCache.ContainsKey($cacheKey)) {
        $e = $script:TokenCache[$cacheKey]
        if ($e.expires -gt $now.AddMinutes(5)) { return $e.token }
    }

    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
        scope         = "$ResourceAppIdUri/.default"
    }

    $resp = Invoke-RestMethod -Method POST -Uri $uri -Body $body -ContentType "application/x-www-form-urlencoded"
    $token = $resp.access_token
    $script:TokenCache[$cacheKey] = @{ token = $token; expires = (Get-Date).AddSeconds([int]$resp.expires_in) }
    return $token
}

function Get-AccessToken {
    param([hashtable]$Conn)
    return (Get-AccessTokenResource -TenantId $Conn.tenant_id -ClientId $Conn.client_id -ClientSecret $Conn.client_secret -ResourceAppIdUri "https://graph.microsoft.com")
}

function Invoke-GraphJson {
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Connexion,
        $Body = $null,
        [int]$MaxRetries = 5
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $tok = Get-AccessToken -Conn $Connexion
            $headers = Build-AuthHeader $tok

            if ($null -ne $Body) {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body ($Body | ConvertTo-Json -Depth 12) -ContentType "application/json"
            }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
        } catch {
            $st = $null
            try { $st = [int]$_.Exception.Response.StatusCode } catch {}
            $retry = ($st -in @(401,408,429)) -or ($st -ge 500 -and $st -lt 600)
            if ($retry -and $attempt -lt $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Min(20, [Math]::Pow(2, $attempt)))
                continue
            }
            throw
        }
    }
}

function Get-HostAndPathFromSiteUrl([string]$siteUrl) {
    $u = [Uri]$siteUrl
    $p = $u.AbsolutePath
    if ($p.Length -gt 1) { $p = $p.TrimEnd('/') }
    [pscustomobject]@{ host = $u.Host; path = $p }
}

function Get-SiteAndDrive {
    param([string]$SiteUrl, [hashtable]$Connexion)

    $tok = Get-AccessToken -Conn $Connexion
    $hp = Get-HostAndPathFromSiteUrl $SiteUrl

    $site = Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($hp.host):$($hp.path)?`$select=id,webUrl,displayName" -Headers (Build-AuthHeader $tok)
    $drives = Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,name,driveType" -Headers (Build-AuthHeader $tok)

    $drive = $drives.value | Where-Object { $_.driveType -eq "documentLibrary" } | Select-Object -First 1
    return [pscustomobject]@{ Site = $site; Drive = $drive }
}

function Item-Exists {
    param([string]$DriveId, [string]$Rel, [string]$AccessToken)

    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$(Encode-RelForGraph $Rel)"
    try {
        $it = Invoke-RestMethod -Method GET -Uri $uri -Headers (Build-AuthHeader $AccessToken)
        return [pscustomobject]@{ Exists = $true; IsFolder = ($null -ne $it.folder); Id = $it.id }
    } catch {
        return [pscustomobject]@{ Exists = $false; IsFolder = $false; Id = $null }
    }
}

function Ensure-FolderPath {
    param([string]$DriveId, [string]$FolderRel, [string]$AccessToken)

    $rel = Normalize-RelPath $FolderRel
    $segs = $rel -split '/' | Where-Object { $_ -ne '' }
    if ($segs.Count -eq 0) { return $null }

    $parentId = "root"
    $lastId = $null

    foreach ($seg in $segs) {
        $enc = UrlEncode-Segment $seg
        $checkUri = if ($parentId -eq "root") {
            "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$enc"
        } else {
            "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$parentId`:/$enc"
        }

        $child = $null
        try {
            $child = Invoke-RestMethod -Method GET -Uri $checkUri -Headers (Build-AuthHeader $AccessToken)
        } catch {
            $createUri = if ($parentId -eq "root") {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
            } else {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$parentId/children"
            }
            $body = @{ name = $seg; folder = @{}; "@microsoft.graph.conflictBehavior" = "replace" } | ConvertTo-Json -Depth 4
            $child = Invoke-RestMethod -Method POST -Uri $createUri -Headers (Build-AuthHeader $AccessToken) -Body $body -ContentType "application/json"
        }

        $parentId = $child.id
        $lastId = $child.id
    }

    return $lastId
}

function Download-FileContent {
    param([string]$DriveId, [string]$RelativePath, [hashtable]$Connexion)

    $tok = Get-AccessToken -Conn $Connexion
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$(Encode-RelForGraph $RelativePath)`:/content"

    $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $uri)
    $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $tok)

    $resp = $script:HttpClient.SendAsync($req).Result
    if (-not $resp.IsSuccessStatusCode) {
        throw "Download HTTP $([int]$resp.StatusCode)"
    }
    return $resp.Content.ReadAsByteArrayAsync().Result
}

function Upload-File {
    param([string]$DriveId, [string]$RelativePath, [byte[]]$Bytes, [hashtable]$Connexion)

    $rel = Normalize-RelPath $RelativePath
    $parent = ([System.IO.Path]::GetDirectoryName($rel) -replace '\\','/')
    $tok = Get-AccessToken -Conn $Connexion

    if ($parent -and $parent -ne ".") {
        Ensure-FolderPath -DriveId $DriveId -FolderRel $parent -AccessToken $tok | Out-Null
    }

    if ($Bytes.Length -le 4MB) {
        $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$(Encode-RelForGraph $rel)`:/content"
        $u = Invoke-RestMethod -Method PUT -Uri $uri -Headers (Build-AuthHeader $tok) -Body $Bytes -ContentType "application/octet-stream"
        return $u.id
    }

    $createUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$(Encode-RelForGraph $rel)`:/createUploadSession"
    $sess = Invoke-GraphJson -Method "POST" -Uri $createUri -Connexion $Connexion -Body @{ item = @{ "@microsoft.graph.conflictBehavior" = "replace" } }
    $uploadUrl = $sess.uploadUrl

    $total = $Bytes.Length
    $pos = 0
    $lastResp = $null

    while ($pos -lt $total) {
        $chunk = [Math]::Min($ChunkSize, $total - $pos)
        $from = $pos
        $to = $pos + $chunk - 1

        $content = [System.Net.Http.ByteArrayContent]::new($Bytes, $from, $chunk)
        $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")
        $null = $content.Headers.TryAddWithoutValidation("Content-Range", ("bytes {0}-{1}/{2}" -f $from, $to, $total))

        $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $uploadUrl)
        $req.Content = $content
        $resp = $script:HttpClient.SendAsync($req).Result

        if (-not $resp.IsSuccessStatusCode) {
            throw "Upload HTTP $([int]$resp.StatusCode)"
        }

        $txt = $resp.Content.ReadAsStringAsync().Result
        if ($txt) {
            try { $lastResp = $txt | ConvertFrom-Json } catch {}
        }

        $pos += $chunk
    }

    if ($lastResp -and $lastResp.id) { return $lastResp.id }
    $check = Item-Exists -DriveId $DriveId -Rel $rel -AccessToken $tok
    return $check.Id
}

function Load-UrlTranslations {
    if (-not (Test-Path $translation_csv)) {
        Write-Log "Fichier de traduction introuvable: $translation_csv" "WARN"
        return
    }

    $csv = Import-Csv -Path $translation_csv -Encoding UTF8
    foreach ($r in $csv) {
        if ($r.Source -and $r.Destination) {
            $script:UrlTranslations[(Normalize-RelPath $r.Source)] = (Normalize-RelPath $r.Destination)
        }
    }
    Write-Log ("Traductions chargees: {0}" -f $script:UrlTranslations.Count) "SUCCESS"
}

function Apply-UrlTranslations {
    param([string]$Path)

    $p = Normalize-RelPath $Path
    foreach ($k in $script:UrlTranslations.Keys) {
        if ($p.Contains($k)) {
            $p = $p.Replace($k, $script:UrlTranslations[$k])
        }
    }
    return $p
}

function Resolve-DestinationObjectId {
    param([string]$DisplayName, [string]$TargetType)

    $cacheKey = "$TargetType|$DisplayName"
    if ($script:ResolvedIdCache.ContainsKey($cacheKey)) {
        Write-Log ("Cache hit resolution destination: {0}" -f $DisplayName) "DEBUG" "permissions" "/"
        return $script:ResolvedIdCache[$cacheKey]
    }

    $safe = $DisplayName.Replace("'","''")
    $isGroup = $TargetType.Trim().ToLowerInvariant() -in @("group","securitygroup","m365group","microsoft365group","aadgroup")

    $uri = if ($isGroup) {
        "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'&`$select=id&`$top=1"
    } else {
        "https://graph.microsoft.com/v1.0/users?`$filter=displayName eq '$safe'&`$select=id&`$top=1"
    }

    try {
        $r = Invoke-GraphJson -Method "GET" -Uri $uri -Connexion $connexion_cible
        $obj = $r.value | Select-Object -First 1
        $id = if ($obj) { $obj.id } else { $null }
        $script:ResolvedIdCache[$cacheKey] = $id
        if ($id) {
            Write-Log ("Resolution destination OK: {0} -> {1}" -f $DisplayName, $id) "DEBUG" "permissions" "/"
        } else {
            Write-Log ("Resolution destination introuvable: {0}" -f $DisplayName) "WARN" "permissions" "/"
        }
        return $id
    } catch {
        $script:ResolvedIdCache[$cacheKey] = $null
        return $null
    }
}

function Apply-Item-Permissions {
    param([string]$PathOriginal, [string]$PathTranslated, [string]$ItemType, [string]$DstDriveId, [string]$ItemId, [object[]]$PermRows)

    Write-Log ("Application des permissions: item={0}, count={1}" -f $PathOriginal, $PermRows.Count) "INFO" "permissions" $PathOriginal
    $applied = 0
    $skipped = 0
    $failed = 0

    foreach ($perm in $PermRows) {
        $srcId = if ($perm.grantedToID) { [string]$perm.grantedToID } else { "" }
        $srcName = if ($perm.GrantedTo) { [string]$perm.GrantedTo } else { "" }
        $srcRole = if ($perm.Role) { [string]$perm.Role } else { "read" }
        $targetType = if ($perm.TargetType) { [string]$perm.TargetType } else { "User" }

        $entry = $null
        if ($srcId -and $script:UserMapping.ContainsKey($srcId)) {
            $entry = $script:UserMapping[$srcId]
        } elseif ($srcName -and $script:UserMappingByName.ContainsKey($srcName)) {
            $entry = $script:UserMappingByName[$srcName]
        }

        if (-not $entry) {
            Db-InsertPermissionResult -PathOriginal $PathOriginal -PathTranslated $PathTranslated -ItemType $ItemType -GrantedTo $srcName -GrantedToId $srcId -TargetType $targetType -Role $srcRole -Status "skipped" -Phase "permissions" -ErrorMessage "Mapping manquant" -MappedDestination $null -MappedDestinationId $null
            $skipped++
            continue
        }

        $dstId = Resolve-DestinationObjectId -DisplayName $entry.DisplayName -TargetType $entry.TargetType
        if (-not $dstId) {
            Db-InsertPermissionResult -PathOriginal $PathOriginal -PathTranslated $PathTranslated -ItemType $ItemType -GrantedTo $srcName -GrantedToId $srcId -TargetType $targetType -Role $srcRole -Status "failed" -Phase "permissions" -ErrorMessage "Resolution destination impossible" -MappedDestination $entry.DisplayName -MappedDestinationId $null
            $failed++
            continue
        }

        $role = switch ($srcRole.ToLowerInvariant()) {
            "owner" { "write" }
            "write" { "write" }
            default { "read" }
        }

        $inviteUri = "https://graph.microsoft.com/v1.0/drives/$DstDriveId/items/$ItemId/invite"
        $body = @{
            requireSignIn = $true
            sendInvitation = $false
            roles = @($role)
            recipients = @(@{ objectId = $dstId })
        }

        try {
            $null = Invoke-GraphJson -Method "POST" -Uri $inviteUri -Connexion $connexion_cible -Body $body
            Db-InsertPermissionResult -PathOriginal $PathOriginal -PathTranslated $PathTranslated -ItemType $ItemType -GrantedTo $srcName -GrantedToId $srcId -TargetType $targetType -Role $role -Status "applied" -Phase "permissions" -ErrorMessage $null -MappedDestination $entry.DisplayName -MappedDestinationId $dstId
            $applied++
        } catch {
            $msg = $_.Exception.Message
            Db-InsertPermissionResult -PathOriginal $PathOriginal -PathTranslated $PathTranslated -ItemType $ItemType -GrantedTo $srcName -GrantedToId $srcId -TargetType $targetType -Role $role -Status "failed" -Phase "permissions" -ErrorMessage $msg -MappedDestination $entry.DisplayName -MappedDestinationId $dstId
            Db-InsertError -Path $PathOriginal -Phase "permissions" -Message $msg
            $failed++
        }
    }

    Write-Log ("Permissions terminees: applied={0}, skipped={1}, failed={2}" -f $applied, $skipped, $failed) "INFO" "permissions" $PathOriginal
}

function Copy-One {
    param([string]$Rel, [string]$ItemType, [string]$SrcDriveId, [string]$DstDriveId, [string]$SrcToken, [string]$DstToken)

    $originalRel = Normalize-RelPath (Strip-LibraryPrefix (Normalize-RelPath $Rel))
    $translatedRel = Normalize-RelPath (Apply-UrlTranslations $originalRel)
    Write-Log ("Preparation copie: src={0}, dst={1}, type={2}" -f $originalRel, $translatedRel, $ItemType) "DEBUG" "copy" $originalRel

    if ([string]::IsNullOrWhiteSpace($originalRel)) {
        return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel; IsFolder = $false; ItemType = $ItemType }
    }

    $dstExisting = Item-Exists -DriveId $DstDriveId -Rel $translatedRel -AccessToken $DstToken
    if ($dstExisting.Exists -and -not $ForceOverwrite) {
        Write-Log ("Element deja present, overwrite desactive: {0}" -f $translatedRel) "INFO" "copy" $translatedRel
        return @{ Ok = $true; ItemId = $dstExisting.Id; TranslatedRel = $translatedRel; IsFolder = ($ItemType -eq "Folder"); ItemType = $ItemType }
    }

    $srcExisting = Item-Exists -DriveId $SrcDriveId -Rel $originalRel -AccessToken $SrcToken
    if (-not $srcExisting.Exists) {
        Db-InsertError -Path $originalRel -Phase "copy" -Message "Source introuvable"
        Write-Log ("Source introuvable: {0}" -f $originalRel) "WARN" "copy" $originalRel
        return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel; IsFolder = $false; ItemType = $ItemType }
    }

    if ($srcExisting.IsFolder -or $ItemType -eq "Folder") {
        Write-Log ("Creation dossier destination: {0}" -f $translatedRel) "DEBUG" "copy" $translatedRel
        $folderId = Ensure-FolderPath -DriveId $DstDriveId -FolderRel $translatedRel -AccessToken $DstToken
        return @{ Ok = $true; ItemId = $folderId; TranslatedRel = $translatedRel; IsFolder = $true; ItemType = "Folder" }
    }

    $bytes = Download-FileContent -DriveId $SrcDriveId -RelativePath $originalRel -Connexion $connexion_source
    $newItemId = Upload-File -DriveId $DstDriveId -RelativePath $translatedRel -Bytes $bytes -Connexion $connexion_cible

    if (-not $newItemId) {
        Db-InsertError -Path $originalRel -Phase "copy" -Message "Upload termine sans item id"
        Write-Log ("Upload sans item id: {0}" -f $translatedRel) "WARN" "copy" $translatedRel
        return @{ Ok = $false; ItemId = $null; TranslatedRel = $translatedRel; IsFolder = $false; ItemType = "File" }
    }

    Write-Log ("Upload termine: {0}" -f $translatedRel) "DEBUG" "copy" $translatedRel

    return @{ Ok = $true; ItemId = $newItemId; TranslatedRel = $translatedRel; IsFolder = $false; ItemType = "File" }
}

# =========================
# MAIN
# =========================
try {
    Initialize-Npgsql
    Ensure-DatabaseSchema

    $discoveryRun = Get-DiscoveryRunId
    Write-Log "Run discovery utilise: $discoveryRun"

    Start-CopyRun -DiscoveryRun $discoveryRun

    Load-UrlTranslations
    Load-UserMapping

    $rows = Get-WorkRowsFromDb -RunId $discoveryRun
    if (-not $rows -or $rows.Count -eq 0) {
        throw "Aucune ligne a copier dans sp_permissions pour run $discoveryRun"
    }

    $grouped = $rows | Group-Object -Property Path
    Write-Log ("Chemins a traiter: {0}" -f $grouped.Count)

    $srcCtx = Get-SiteAndDrive -SiteUrl $site_url_source -Connexion $connexion_source
    $dstCtx = Get-SiteAndDrive -SiteUrl $site_url_cible -Connexion $connexion_cible

    if (-not $srcCtx.Drive -or -not $dstCtx.Drive) {
        throw "Drive source ou destination introuvable"
    }

    $srcToken = Get-AccessToken -Conn $connexion_source
    $dstToken = Get-AccessToken -Conn $connexion_cible

    foreach ($group in $grouped) {
        $path = [string]$group.Name
        $permRows = $group.Group
        $itemType = [string]$permRows[0].ItemType

        Write-Log ("Traitement: {0}" -f $path)

        try {
            $copy = Copy-One -Rel $path -ItemType $itemType -SrcDriveId $srcCtx.Drive.id -DstDriveId $dstCtx.Drive.id -SrcToken $srcToken -DstToken $dstToken

            if ($copy.Ok) {
                Db-UpsertCopyItem -PathOriginal $path -PathTranslated $copy.TranslatedRel -ItemType $copy.ItemType -IsFolder $copy.IsFolder -ItemId $copy.ItemId -Status "copied"
                Apply-Item-Permissions -PathOriginal $path -PathTranslated $copy.TranslatedRel -ItemType $copy.ItemType -DstDriveId $dstCtx.Drive.id -ItemId $copy.ItemId -PermRows $permRows
                Write-Log ("Copie OK: {0}" -f $copy.TranslatedRel) "SUCCESS"
            } else {
                Db-UpsertCopyItem -PathOriginal $path -PathTranslated $copy.TranslatedRel -ItemType $itemType -IsFolder ($itemType -eq "Folder") -ItemId $null -Status "failed"
                Write-Log ("Copie KO: {0}" -f $path) "WARN"
            }
        } catch {
            $msg = $_.Exception.Message
            Db-UpsertCopyItem -PathOriginal $path -PathTranslated $null -ItemType $itemType -IsFolder ($itemType -eq "Folder") -ItemId $null -Status "failed"
            Db-InsertError -Path $path -Phase "copy" -Message $msg
            Write-Log ("Erreur: {0}" -f $msg) "ERROR"
        }
    }

    Finish-CopyRun -Status "success"
    Write-Log "Copie terminee" "SUCCESS"
}
catch {
    $msg = $_.Exception.Message
    Write-Log ("Erreur fatale: {0}" -f $msg) "ERROR"
    if ($script:CopyRunId) {
        try {
            Db-InsertError -Path "/" -Phase "fatal" -Message $msg
            Finish-CopyRun -Status "failed"
        } catch {}
    }
    exit 1
}
finally {
    if ($script:HttpClient) { $script:HttpClient.Dispose() }
}
