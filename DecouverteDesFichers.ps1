# Script PowerShell - Decouverte SharePoint vers PostgreSQL
# Version: 1.0

try {
    if ($Host -and $Host.Name -and $Host.Name -notmatch 'ISE') {
        if (-not [Console]::IsInputRedirected) {
            [Console]::InputEncoding = [System.Text.Encoding]::UTF8
        }
        if (-not [Console]::IsOutputRedirected) {
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        }
    }
} catch {}
try {
    if ($Host -and $Host.UI) {
        $OutputEncoding = [System.Text.Encoding]::UTF8
    }
} catch {}

# =========================
# CONFIG GRAPH
# =========================
$tenantId = "85a5a352-25d6-4894-a97d-221cd1712dd2"
$clientId = "684f670e-6931-4120-b053-f8c9538744f6"
$clientSecret = "O5s8Q~Z~e_d1-dSQ.vt4uFTeR8c4Xssaa093Racz"

$siteUrl = "https://graph.microsoft.com/v1.0/sites/ipdlab.sharepoint.com/:/sites/SERVICES_GENERAUX"
$documentLibraryName = "Documents"

# Compatibilite: export CSV facultatif
$enableCsvExport = $true
$exportPath  = "E:\SharePoint_Permissions_Export_src9999.csv"
$mappingPath = "E:\UserMapping.csv"

# =========================
# CONFIG POSTGRES
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
# ETAT
# =========================
$script:accessToken = $null
$script:tokenExpiry = $null
$script:runId = $null
$script:NpgsqlReady = $false
$script:NpgsqlVersion = "8.0.8"
$script:AssemblyResolverReady = $false
$script:AssemblyPathCache = @{}
$script:PinnedResolverReady = $false
$script:DbSchema = "mercury"
$script:allPaths = [System.Collections.Generic.HashSet[string]]::new()
$script:permRowsForCsv = New-Object System.Collections.Generic.List[object]
$script:DbLogEnabled = $true
$script:DbLogWriteInProgress = $false
$script:DbLogFailureWarned = $false

# =========================
# UTILS
# =========================
function Write-Log {
    param(
        [string]$message,
        [string]$level = "INFO",
        [string]$phase = "runtime",
        [string]$path = "/",
        $context = $null
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp][$level] $message"

    if (-not $script:DbLogEnabled) { return }
    if (-not $script:runId) { return }
    if (-not $script:NpgsqlReady) { return }
    if ($script:DbLogWriteInProgress) { return }

    try {
        $script:DbLogWriteInProgress = $true
        $ctxObj = [ordered]@{
            source = "discovery"
            phase = $phase
            path = $path
        }
        if ($context) {
            if ($context -is [hashtable] -or $context -is [System.Collections.IDictionary]) {
                foreach ($k in $context.Keys) { $ctxObj[[string]$k] = $context[$k] }
            } else {
                $ctxObj["detail"] = [string]$context
            }
        }
        $ctx = ($ctxObj | ConvertTo-Json -Depth 8 -Compress)
        $sql = @"
INSERT INTO sp_logs (run_id, log_level, message, context)
VALUES (
    $(Sql-Literal $script:runId),
    $(Sql-Literal $level),
    $(Sql-Literal $message),
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

function Sql-Literal([AllowNull()][string]$value) {
    if ($null -eq $value) { return "NULL" }
    return "'" + ($value -replace "'", "''") + "'"
}

function Get-LoaderExceptionText {
    param([System.Exception]$Ex)

    if ($Ex -and $Ex.LoaderExceptions) {
        $msgs = @($Ex.LoaderExceptions | Where-Object { $_ } | ForEach-Object { $_.Message })
        if ($msgs.Count -gt 0) {
            return ($msgs -join " | ")
        }
    }
    return $Ex.Message
}

function Register-NuGetAssemblyResolver {
    if ($script:AssemblyResolverReady) { return }
    $script:AssemblyResolverReady = $true
}

function Register-PinnedAssemblyResolver {
    if ($script:PinnedResolverReady) { return }

    $handler = [System.ResolveEventHandler]{
        param($sender, $args)

        try {
            $asm = New-Object System.Reflection.AssemblyName($args.Name)
            $name = $asm.Name
            $ver = if ($asm.Version) { $asm.Version.ToString() } else { "" }

            $map = @{}
            $map["System.Runtime.CompilerServices.Unsafe|6.0.0.0"] = "C:\Users\ffaradji-op\.nuget\packages\system.runtime.compilerservices.unsafe\6.0.0\lib\netstandard2.0\System.Runtime.CompilerServices.Unsafe.dll"
            $map["System.Runtime.CompilerServices.Unsafe|4.0.4.1"] = "C:\Users\ffaradji-op\Desktop\Mercury script powershell\deps\System.Runtime.CompilerServices.Unsafe.4.0.4.1.dll"
            $map["System.Buffers|4.0.2.0"] = "C:\Users\ffaradji-op\Desktop\Mercury script powershell\deps\System.Buffers.4.0.2.0.dll"
            $map["System.Buffers|4.0.3.0"] = "C:\Users\ffaradji-op\Desktop\Mercury script powershell\deps\System.Buffers.4.0.3.0.dll"

            $key = "{0}|{1}" -f $name, $ver
            if ($map.ContainsKey($key)) {
                $dll = $map[$key]
                if (Test-Path -LiteralPath $dll) {
                    return [System.Reflection.Assembly]::LoadFrom($dll)
                }
            }
        } catch {
            return $null
        }

        return $null
    }

    [System.AppDomain]::CurrentDomain.add_AssemblyResolve($handler)
    $script:PinnedResolverReady = $true
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
    Register-PinnedAssemblyResolver

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

            # Chargement explicite prioritaire pour PS5 + Npgsql 8.
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\microsoft.bcl.asyncinterfaces\8.0.0\lib\net462\Microsoft.Bcl.AsyncInterfaces.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\microsoft.extensions.logging.abstractions\8.0.0\lib\net462\Microsoft.Extensions.Logging.Abstractions.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\microsoft.extensions.primitives\8.0.0\lib\net462\Microsoft.Extensions.Primitives.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.threading.tasks.extensions\4.5.4\lib\net461\System.Threading.Tasks.Extensions.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.memory\4.5.5\lib\net461\System.Memory.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\Desktop\Mercury script powershell\deps\System.Buffers.4.0.3.0.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.runtime.compilerservices.unsafe\6.0.0\lib\netstandard2.0\System.Runtime.CompilerServices.Unsafe.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\Desktop\Mercury script powershell\deps\System.Runtime.CompilerServices.Unsafe.4.0.4.1.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.numerics.vectors\4.5.0\lib\net46\System.Numerics.Vectors.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.text.json\8.0.5\lib\netstandard2.0\System.Text.Json.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.text.encodings.web\8.0.0\lib\net462\System.Text.Encodings.Web.dll" | Out-Null
            Import-DependencyByPath -DllPath "C:\Users\ffaradji-op\.nuget\packages\system.threading.channels\8.0.0\lib\net462\System.Threading.Channels.dll" | Out-Null

            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Extensions.Logging.Abstractions" -DllName "Microsoft.Extensions.Logging.Abstractions.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Bcl.AsyncInterfaces" -DllName "Microsoft.Bcl.AsyncInterfaces.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Extensions.DependencyInjection.Abstractions" -DllName "Microsoft.Extensions.DependencyInjection.Abstractions.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Extensions.Options" -DllName "Microsoft.Extensions.Options.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "Microsoft.Extensions.Primitives" -DllName "Microsoft.Extensions.Primitives.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Diagnostics.DiagnosticSource" -DllName "System.Diagnostics.DiagnosticSource.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Text.Json" -DllName "System.Text.Json.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Text.Encodings.Web" -DllName "System.Text.Encodings.Web.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Threading.Tasks.Extensions" -DllName "System.Threading.Tasks.Extensions.dll"
            Import-PackageDependency -PackagesRoot $packagesRoot -PackagePrefix "System.Runtime.CompilerServices.Unsafe" -DllName "System.Runtime.CompilerServices.Unsafe.dll"
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

    $candidateRoots = @(
        "$env:USERPROFILE\\.nuget\\packages\\npgsql",
        "$env:LOCALAPPDATA\\PackageManagement\\NuGet\\Packages\\Npgsql.*",
        "$env:ProgramFiles\\PackageManagement\\NuGet\\Packages\\Npgsql.*"
    )

    $expandedRoots = New-Object System.Collections.Generic.List[string]
    foreach ($pattern in $candidateRoots) {
        $items = @(Get-ChildItem -Path $pattern -Directory -ErrorAction SilentlyContinue)
        foreach ($it in $items) { $expandedRoots.Add($it.FullName) | Out-Null }
    }

    $expandedRoots = @($expandedRoots | Sort-Object -Descending -Unique)
    foreach ($root in $expandedRoots) {
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

    $installedRoot = Join-Path -Path "$env:USERPROFILE\\.nuget\\packages\\npgsql" -ChildPath $script:NpgsqlVersion
    if (-not (Import-NpgsqlFromRoot -Root $installedRoot)) {
        throw "Npgsql installe mais non chargeable. Verifiez .NET Framework 4.7.2+ et les dependances runtime."
    }

    $script:NpgsqlReady = $true
    Write-Log "Provider Npgsql installe et charge avec succes" "SUCCESS"
}

function New-PgConnectionString {
    $sslMode = if ([string]::IsNullOrWhiteSpace($pg.sslmode)) { "Prefer" } else { $pg.sslmode }
    $conn = "Host=$($pg.host);Port=$($pg.port);Database=$($pg.database);Username=$($pg.user);Password=$($pg.password);SSL Mode=$sslMode;Trust Server Certificate=true;"
    if ($script:DbSchema) {
        $conn += ";SearchPath=$($script:DbSchema)"
    }
    return $conn
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

function Start-DiscoveryRun {
    $newRunId = [guid]::NewGuid().ToString()
    $sql = @"
INSERT INTO sp_runs (run_id, run_type, status, source_site_url, source_tenant_id, source_library)
VALUES ($(Sql-Literal $newRunId), 'discovery', 'running', $(Sql-Literal $siteUrl), $(Sql-Literal $tenantId), $(Sql-Literal $documentLibraryName))
RETURNING run_id;
"@
    $rows = Invoke-PgQuery -Sql $sql
    if (-not $rows -or -not $rows[0].run_id) {
        throw "Impossible de creer le run en base"
    }
    $script:runId = [string]$rows[0].run_id
    Write-Log "Run discovery cree: $($script:runId)" "SUCCESS"
}

function Finish-DiscoveryRun {
    param([string]$status)

    $sql = @"
UPDATE sp_runs
SET finished_at = now(),
    status = $(Sql-Literal $status),
    total_items = (SELECT COUNT(*) FROM sp_items WHERE run_id = $(Sql-Literal $script:runId)),
    total_permissions = (SELECT COUNT(*) FROM sp_permissions WHERE run_id = $(Sql-Literal $script:runId)),
    total_errors = (SELECT COUNT(*) FROM sp_errors WHERE run_id = $(Sql-Literal $script:runId))
WHERE run_id = $(Sql-Literal $script:runId);
"@
    Invoke-PgNonQuery -Sql $sql
}

function Db-InsertError {
    param([string]$path, [string]$phase, [string]$message)

    if (-not $script:runId) { return }
    $sql = @"
INSERT INTO sp_errors (run_id, path, phase, error_message, error_level)
VALUES (
    $(Sql-Literal $script:runId),
    $(Sql-Literal $path),
    $(Sql-Literal $phase),
    $(Sql-Literal $message),
    'ERROR'
);
"@
    Invoke-PgNonQuery -Sql $sql
}

function Db-InsertRuntimeLog {
    param([string]$level, [string]$message)

    if (-not $script:runId) { return }
    if (-not $script:NpgsqlReady) { return }
    try {
        $ctx = ('{"source":"discovery","phase":"runtime","path":"/"}')
        $sql = @"
INSERT INTO sp_logs (run_id, log_level, message, context)
VALUES (
    $(Sql-Literal $script:runId),
    $(Sql-Literal $level),
    $(Sql-Literal $message),
    $(Sql-Literal $ctx)::jsonb
);
"@
        Invoke-PgNonQuery -Sql $sql
    } catch {
        Write-Host "[WARN] Echec insertion log runtime DB: $($_.Exception.Message)"
    }
}

function Db-UpsertItem {
    param(
        [string]$site,
        [string]$driveId,
        [string]$itemId,
        [string]$path,
        [string]$itemType,
        [bool]$isFolder,
        [long]$sizeBytes
    )

    $sql = @"
INSERT INTO sp_items (
    run_id, site_url, drive_id, item_id, path_original, item_type, is_folder, size_bytes,
    copy_status, created_at, updated_at
)
VALUES (
    $(Sql-Literal $script:runId),
    $(Sql-Literal $site),
    $(Sql-Literal $driveId),
    $(Sql-Literal $itemId),
    $(Sql-Literal $path),
    $(Sql-Literal $itemType),
    $(if ($isFolder) { 'TRUE' } else { 'FALSE' }),
    $sizeBytes,
    'pending',
    now(),
    now()
)
ON CONFLICT (run_id, path_original)
DO UPDATE SET
    item_id = EXCLUDED.item_id,
    item_type = EXCLUDED.item_type,
    is_folder = EXCLUDED.is_folder,
    size_bytes = EXCLUDED.size_bytes,
    updated_at = now();
"@
    Invoke-PgNonQuery -Sql $sql
}

function Db-InsertPermission {
    param(
        [string]$site,
        [string]$driveId,
        [string]$itemId,
        [string]$path,
        [string]$itemType,
        [string]$grantedTo,
        [string]$grantedToId,
        [string]$targetType,
        [string]$roleName
    )

    $sql = @"
INSERT INTO sp_permissions (
    run_id, site_url, drive_id, item_id,
    path_original, item_type,
    granted_to, granted_to_id, target_type, role_name,
    permission_status, phase, created_at
)
VALUES (
    $(Sql-Literal $script:runId),
    $(Sql-Literal $site),
    $(Sql-Literal $driveId),
    $(Sql-Literal $itemId),
    $(Sql-Literal $path),
    $(Sql-Literal $itemType),
    $(Sql-Literal $grantedTo),
    $(Sql-Literal $grantedToId),
    $(Sql-Literal $targetType),
    $(Sql-Literal $roleName),
    'detected',
    'discovery',
    now()
);
"@
    Invoke-PgNonQuery -Sql $sql

    if ($grantedTo -and $grantedTo -ne '<unknown>' -and $grantedToId -ne 'Unknown') {
        $mapSql = @"
INSERT INTO sp_user_mapping (
    source_id,
    source_display_name,
    target_type,
    destination_display_name,
    destination_object_id,
    is_active,
    created_at,
    updated_at
)
VALUES (
    NULLIF($(Sql-Literal $grantedToId), ''),
    $(Sql-Literal $grantedTo),
    $(Sql-Literal $targetType),
    '',
    NULL,
    true,
    now(),
    now()
)
ON CONFLICT (source_display_name)
DO UPDATE SET
    source_id = COALESCE(EXCLUDED.source_id, sp_user_mapping.source_id),
    target_type = EXCLUDED.target_type,
    is_active = true,
    updated_at = now();
"@
        try { Invoke-PgNonQuery -Sql $mapSql } catch {}
    }
}

function Sync-UserMappingFromPermissions {
    Write-Log "Synchronisation sp_user_mapping depuis sp_permissions" "INFO" "mapping" "/"
    $sql = @"
WITH dedup AS (
    SELECT
        granted_to,
        MAX(NULLIF(granted_to_id, '')) AS source_id,
        MAX(target_type) AS target_type
    FROM sp_permissions
    WHERE run_id = $(Sql-Literal $script:runId)
      AND COALESCE(granted_to, '') <> ''
      AND granted_to <> '<unknown>'
      AND COALESCE(granted_to_id, '') <> 'Unknown'
    GROUP BY granted_to
)
INSERT INTO sp_user_mapping (
    source_id,
    source_display_name,
    target_type,
    destination_display_name,
    destination_object_id,
    is_active,
    created_at,
    updated_at
)
SELECT
    source_id,
    granted_to,
    target_type,
    '',
    NULL,
    true,
    now(),
    now()
FROM dedup
ON CONFLICT (source_display_name)
DO UPDATE SET
    source_id = COALESCE(EXCLUDED.source_id, sp_user_mapping.source_id),
    target_type = EXCLUDED.target_type,
    is_active = true,
    updated_at = now();
"@

    Invoke-PgNonQuery -Sql $sql
    Write-Log "Synchronisation sp_user_mapping terminee" "SUCCESS" "mapping" "/"
}

function Export-MappingFromDb {
    if (-not $mappingPath) { return }

    Write-Log ("Generation export mapping CSV: {0}" -f $mappingPath) "INFO" "mapping" "/"

    $sql = @"
SELECT DISTINCT
    granted_to_id AS "SourceId",
    granted_to AS "SourceDisplayName",
    target_type AS "TargetType",
    ''::text AS "DestinationDisplayName"
FROM sp_permissions
WHERE run_id = $(Sql-Literal $script:runId)
  AND COALESCE(granted_to_id, '') <> ''
  AND granted_to_id <> 'Unknown'
  AND granted_to <> '<unknown>'
ORDER BY granted_to;
"@

    $rows = Invoke-PgQuery -Sql $sql
    if ($rows.Count -eq 0) {
        Write-Log "Aucune entree mapping a exporter" "WARN"
        return
    }

    $rows | Export-Csv -Path $mappingPath -NoTypeInformation -Encoding UTF8
    Write-Log "Mapping CSV genere: $mappingPath" "SUCCESS"
}

# =========================
# AUTH / GRAPH
# =========================
function Get-AccessToken {
    param(
        [string]$clientId,
        [string]$clientSecret,
        [string]$tenantId
    )

    if ($script:accessToken -and $script:tokenExpiry -and (Get-Date).AddMinutes(5) -lt $script:tokenExpiry) {
        return $script:accessToken
    }

    $tokenBody = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $clientId
        client_secret = $clientSecret
    }

    $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $tokenBody
    $script:accessToken = $tokenResponse.access_token
    $script:tokenExpiry = (Get-Date).AddSeconds($tokenResponse.expires_in)
    return $script:accessToken
}

function Invoke-With-Retry {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        $Body = $null,
        [int]$MaxRetries = 5
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $token = Get-AccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId
            $headers = @{ Authorization = "Bearer $token" }

            if ($Method -eq "GET") {
                return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers
            }

            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body ($Body | ConvertTo-Json -Depth 10) -ContentType "application/json"
        } catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

            if ($statusCode -eq 401) {
                $script:accessToken = $null
                $script:tokenExpiry = $null
            }

            $retryable = ($statusCode -in @(401,408,429)) -or ($statusCode -ge 500 -and $statusCode -lt 600)
            if ($retryable -and $attempt -lt $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Min(20, [Math]::Pow(2, $attempt)))
                continue
            }
            throw
        }
    }
}

# =========================
# SCAN
# =========================
function Get-Item-Permissions {
    param(
        [string]$driveId,
        [string]$itemId,
        [string]$path,
        [string]$type
    )

    $permUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/permissions"
    $permCount = 0
    $roleCount = 0

    do {
        $permResponse = Invoke-With-Retry -Uri $permUrl
        if (-not $permResponse) { return }

        foreach ($perm in $permResponse.value) {
            $permCount++
            $grantedTo = "<unknown>"
            $targetType = "Unknown"
            $grantedToID = "Unknown"

            if ($perm.grantedToV2 -and $perm.grantedToV2.user) {
                $grantedTo = $perm.grantedToV2.user.displayName
                $targetType = "User"
                $grantedToID = $perm.grantedToV2.user.id
            } elseif ($perm.grantedToV2 -and $perm.grantedToV2.group) {
                $grantedTo = $perm.grantedToV2.group.displayName
                $targetType = "Group"
                $grantedToID = $perm.grantedToV2.group.id
            } elseif ($perm.grantedTo -and $perm.grantedTo.user) {
                $grantedTo = $perm.grantedTo.user.displayName
                $targetType = "User"
                $grantedToID = $perm.grantedTo.user.id
            }

            foreach ($role in $perm.roles) {
                $roleCount++
                Db-InsertPermission -site $siteUrl -driveId $driveId -itemId $itemId -path $path -itemType $type -grantedTo $grantedTo -grantedToId $grantedToID -targetType $targetType -roleName $role

                if ($enableCsvExport) {
                    $row = [pscustomobject]@{
                        Path        = $path
                        ItemType    = $type
                        GrantedTo   = $grantedTo
                        grantedToID = $grantedToID
                        TargetType  = $targetType
                        Role        = $role
                    }
                    $script:permRowsForCsv.Add($row) | Out-Null
                }
            }
        }

        $permUrl = $permResponse.'@odata.nextLink'
    } while ($permUrl)

    Write-Log ("Permissions detectees: item={0}, perms={1}, roles={2}" -f $path, $permCount, $roleCount) "DEBUG" "permissions" $path @{ item_type = $type; item_id = $itemId }
}

function Get-Children-With-Paging {
    param(
        [string]$driveId,
        [string]$itemId,
        [string]$path = "/"
    )

    $url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/children"
    Write-Log ("Lecture enfants: parent={0}" -f $path) "DEBUG" "scan" $path @{ parent_item_id = $itemId }

    do {
        $response = Invoke-With-Retry -Uri $url
        if (-not $response) { return }
        Write-Log ("Page enfants recue: parent={0}, count={1}" -f $path, @($response.value).Count) "DEBUG" "scan" $path

        foreach ($item in $response.value) {
            $itemPath = "$path$($item.name)"
            if ($item.folder) {
                $itemPath = "$itemPath/"
                [void]$script:allPaths.Add($itemPath)
                Db-UpsertItem -site $siteUrl -driveId $driveId -itemId $item.id -path $itemPath -itemType "Folder" -isFolder $true -sizeBytes 0
                Get-Item-Permissions -driveId $driveId -itemId $item.id -path $itemPath -type "Folder"
                Get-Children-With-Paging -driveId $driveId -itemId $item.id -path $itemPath
            } else {
                [void]$script:allPaths.Add($itemPath)
                $size = 0
                try { $size = [long]$item.size } catch { $size = 0 }
                Db-UpsertItem -site $siteUrl -driveId $driveId -itemId $item.id -path $itemPath -itemType "File" -isFolder $false -sizeBytes $size
                Get-Item-Permissions -driveId $driveId -itemId $item.id -path $itemPath -type "File"
            }
        }

        $url = $response.'@odata.nextLink'
    } while ($url)
}

# =========================
# MAIN
# =========================
Write-Log "Demarrage scan SharePoint -> PostgreSQL"

try {
    Initialize-Npgsql
    Ensure-DatabaseSchema
    Start-DiscoveryRun

    Write-Log "Connexion au site SharePoint"
    $site = Invoke-With-Retry -Uri $siteUrl
    if (-not $site) { throw "Connexion au site impossible" }

    $siteId = $site.id
    $drives = Invoke-With-Retry -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
    if (-not $drives) { throw "Recuperation des drives impossible" }

    $drive = $drives.value | Where-Object { $_.name -eq $documentLibraryName } | Select-Object -First 1
    if (-not $drive) { throw "Drive '$documentLibraryName' non trouve" }

    Write-Log "Drive trouve: $($drive.name)"

    if ($enableCsvExport -and (Test-Path $exportPath)) {
        try {
            Remove-Item -LiteralPath $exportPath -Force -ErrorAction Stop
        } catch {
            Write-Log ("Impossible de supprimer l'ancien export CSV: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    Get-Children-With-Paging -driveId $drive.id -itemId "root" -path "/"

    if ($enableCsvExport) {
        $script:permRowsForCsv | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
        Write-Log "Export CSV permissions: $exportPath" "SUCCESS"
    }

    Sync-UserMappingFromPermissions
    Write-Log "Table sp_user_mapping synchronisee depuis sp_permissions" "SUCCESS"

    Export-MappingFromDb
    Finish-DiscoveryRun -status "success"

    Write-Log "Scan termine: $($script:allPaths.Count) items" "SUCCESS"
}
catch {
    $msg = $_.Exception.Message
    $inner = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { "" }
    $where = $_.ScriptStackTrace
    Write-Log "Erreur critique: $msg" "ERROR"
    if ($inner) { Write-Log "Inner: $inner" "ERROR" }
    if ($where) { Write-Log "Stack: $where" "ERROR" }
    try { Db-InsertError -path "/" -phase "fatal" -message ("$msg | $inner | $where") } catch {}
    try { Sync-UserMappingFromPermissions } catch {}
    if ($script:runId) {
        try { Finish-DiscoveryRun -status "failed" } catch {}
    }
    exit 1
}
