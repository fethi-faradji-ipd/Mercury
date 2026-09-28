# ============================================================
#  Get-PowerBILicenseReport.ps1
#  Rapport licences Power BI Pro : date d'affectation + dernière activité
#  Authentification via Service Principal (application Azure AD)
# ============================================================
#
#  PRÉREQUIS
#  ---------
#  Modules à installer (une seule fois) :
#    Install-Module Microsoft.Graph      -Scope CurrentUser -Force
#    Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser -Force
#
#  CONFIGURATION DE L'APPLICATION AZURE AD
#  ----------------------------------------
#  1. Portail Azure > Microsoft Entra ID > Inscriptions d'applications
#     → Nouvelle inscription (ex. "PowerBI-License-Report")
#
#  2. Permissions API requises (type : Application, pas Déléguées) :
#     Microsoft Graph :
#       - User.Read.All
#       - AuditLog.Read.All
#       - Directory.Read.All
#     → Accorder le consentement administrateur
#
#  3. Power BI Admin Portal (app.powerbi.com > Admin > Paramètres locataire) :
#     → Activer "Autoriser les principaux de service à utiliser les API
#       d'administration Power BI en lecture seule"
#     → Ajouter le groupe de sécurité contenant votre Service Principal
#       (ou activer pour toute l'organisation)
#
#  4. Créer un secret client :
#     Inscriptions d'applications > votre app > Certificats et secrets
#     → Nouveau secret client → copier la valeur
#
#  PARAMÈTRES
#  ----------
#  Renseigner TenantId, ClientId et ClientSecret ci-dessous,
#  ou les passer en ligne de commande.
#
# ============================================================

    $TenantId= "2eea08b8-1972-447b-ad43-d044d042500a"
    $ClientId="62a03bb6-b52a-4b2d-a3a7-82e542026446"
    $ClientSecret="6Xp8Q~PEE3L.1qRNR9EsSdV9.Adxi5mMFy_8TbgC" 

    # --- Options du rapport ---
       $ActivityDaysBack     = 30   # max 30 j (limite API Power BI)
      $LicenseAuditDaysBack = 90   # max 90 j (audit Entra sur E3/E5)
   $OutputPath           = "$env:USERPROFILE\Downloads\PowerBI_Licenses_Report.csv"

# ── SKU IDs Power BI ──────────────────────────────────────
$SkuPowerBIPro = "f8cdef31-a31e-4b4a-93e4-5f571e91255a"   # Power BI Pro
$SkuPowerBIPPU = "9c0dab89-a30c-4117-86e7-97bda240acd2"   # Premium Per User

# ─────────────────────────────────────────────────────────
#  ÉTAPE 0 – Connexions via Service Principal (token OAuth2)
# ─────────────────────────────────────────────────────────

# --- 0a. Token Microsoft Graph ---
Write-Host "`n=== Connexion Microsoft Graph (Service Principal) ===" -ForegroundColor Cyan
try {
    $graphTokenResponse = Invoke-RestMethod `
        -Method POST `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
        }

    $graphToken = ConvertTo-SecureString $graphTokenResponse.access_token -AsPlainText -Force
    Connect-MgGraph -AccessToken $graphToken -NoWelcome
    Write-Host "  → Connecté à Microsoft Graph." -ForegroundColor Green
}
catch {
    Write-Error "Échec de la connexion Microsoft Graph : $_"
    exit 1
}

# --- 0b. Token Power BI ---
Write-Host "=== Connexion Power BI Service (Service Principal) ===" -ForegroundColor Cyan
try {
    $SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $SpCredential = New-Object System.Management.Automation.PSCredential($ClientId, $SecureSecret)

    Connect-PowerBIServiceAccount `
        -ServicePrincipal `
        -Credential $SpCredential `
        -TenantId $TenantId
    Write-Host "  → Connecté au service Power BI." -ForegroundColor Green
}
catch {
    Write-Error "Échec de la connexion Power BI : $_"
    exit 1
}

# ─────────────────────────────────────────────────────────
#  ÉTAPE 1 – Utilisateurs avec licence Power BI Pro / PPU
# ─────────────────────────────────────────────────────────
Write-Host "`n[1/3] Récupération des utilisateurs licenciés..." -ForegroundColor Yellow

$allUsers = Get-MgUser -All `
    -Property "Id,DisplayName,UserPrincipalName,AssignedLicenses,AccountEnabled"

$licensedUsers = $allUsers | Where-Object {
    $_.AssignedLicenses.SkuId -contains "-" -or
    $_.AssignedLicenses.SkuId -contains "-"
}

Write-Host "  → $($licensedUsers.Count) utilisateur(s) avec licence Power BI Pro / PPU trouvé(s)."

if ($licensedUsers.Count -eq 0) {
    Write-Warning "Aucun utilisateur trouvé. Vérifiez les permissions et les SKU IDs."
    exit
}

# ─────────────────────────────────────────────────────────
#  ÉTAPE 2 – Date d'affectation via l'audit log Entra ID
# ─────────────────────────────────────────────────────────
Write-Host "`n[2/3] Récupération des dates d'affectation (audit log Entra, J-$LicenseAuditDaysBack)..." -ForegroundColor Yellow

$auditFrom = (Get-Date).AddDays(-$LicenseAuditDaysBack).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# Table : UPN → date d'affectation la plus récente
$assignmentDates = @{}

try {
    $auditEntries = Get-MgAuditLogDirectoryAudit -All `
        -Filter "activityDisplayName eq 'Change user license' and activityDateTime ge $auditFrom and result eq 'success'"

    foreach ($entry in $auditEntries) {
        foreach ($target in $entry.TargetResources) {
            $upn = $target.UserPrincipalName
            if (-not $upn) { continue }

            # Vérifier que la modification concerne bien un SKU Power BI
            $newLicenses = ($target.ModifiedProperties |
                Where-Object { $_.DisplayName -eq "AssignedLicense" }).NewValue

            if ($newLicenses -match $SkuPowerBIPro -or $newLicenses -match $SkuPowerBIPPU) {
                $eventDate = $entry.ActivityDateTime
                # Conserver la date la plus récente pour cet utilisateur
                if (-not $assignmentDates.ContainsKey($upn) -or $eventDate -gt $assignmentDates[$upn]) {
                    $assignmentDates[$upn] = $eventDate
                }
            }
        }
    }

    Write-Host "  → $($assignmentDates.Count) date(s) d'affectation retrouvée(s) dans l'audit log."
}
catch {
    Write-Warning "  Impossible de lire l'audit log Entra : $_"
    Write-Warning "  Vérifiez que la permission AuditLog.Read.All (Application) est accordée et que le consentement admin a été donné."
}

# ─────────────────────────────────────────────────────────
#  ÉTAPE 3 – Dernière activité via l'API Power BI Admin
# ─────────────────────────────────────────────────────────
Write-Host "`n[3/3] Récupération de l'activité Power BI (J-$ActivityDaysBack à aujourd'hui)..." -ForegroundColor Yellow

# Table : UPN → dernière date d'activité
$lastActivityDates = @{}

$startDay = (Get-Date).AddDays(-[Math]::Min($ActivityDaysBack, 30)).Date   # max 30 j
$today    = (Get-Date).Date

for ($day = $startDay; $day -le $today; $day = $day.AddDays(1)) {
    $dayStr = $day.ToString("yyyy-MM-dd")
    Write-Host "  Analyse du $dayStr..." -ForegroundColor Gray

    try {
        $eventsRaw = Get-PowerBIActivityEvent `
            -StartDateTime "$($dayStr)T00:00:00.000Z" `
            -EndDateTime   "$($dayStr)T23:59:59.999Z"

        # L'API renvoie une chaîne JSON
        if ($eventsRaw) {
            $events = $eventsRaw | ConvertFrom-Json -ErrorAction SilentlyContinue
            if (-not $events) { continue }

            foreach ($ev in $events) {
                # L'identifiant utilisateur peut être dans UserId ou UserPrincipalName
                $upn = if ($ev.UserPrincipalName) { $ev.UserPrincipalName } else { $ev.UserId }
                if (-not $upn) { continue }

                $evTime = [datetime]$ev.CreationTime
                if (-not $lastActivityDates.ContainsKey($upn) -or $evTime -gt $lastActivityDates[$upn]) {
                    $lastActivityDates[$upn] = $evTime
                }
            }
        }
    }
    catch {
        Write-Warning "  Erreur pour $dayStr : $_"
        Write-Warning "  Vérifiez que l'option 'Autoriser les principaux de service à utiliser les API d'administration Power BI' est activée dans le portail Admin Power BI."
    }
}

Write-Host "  → $($lastActivityDates.Count) utilisateur(s) actif(s) détecté(s) sur la période."

# ─────────────────────────────────────────────────────────
#  ÉTAPE 4 – Construction et export du rapport
# ─────────────────────────────────────────────────────────
Write-Host "`nConstruction du rapport..." -ForegroundColor Yellow

$report = foreach ($user in $licensedUsers) {
    $upn = $user.UserPrincipalName

    $licenseType = if ($user.AssignedLicenses.SkuId -contains $SkuPowerBIPPU) {
        "Premium Per User"
    } else {
        "Power BI Pro"
    }

    $assignDate = if ($assignmentDates.ContainsKey($upn)) {
        $assignmentDates[$upn].ToString("dd/MM/yyyy HH:mm")
    } else {
        "Non trouvée (> $LicenseAuditDaysBack j ou affectation par groupe)"
    }

    $lastActivity = if ($lastActivityDates.ContainsKey($upn)) {
        $lastActivityDates[$upn].ToString("dd/MM/yyyy HH:mm")
    } else {
        "Aucune activité sur J-$ActivityDaysBack"
    }

    $inactive = if ($lastActivityDates.ContainsKey($upn)) { "Non" } else { "Oui" }

    [PSCustomObject]@{
        Nom                     = $user.DisplayName
        UserPrincipalName       = $upn
        CompteActif             = if ($user.AccountEnabled) { "Oui" } else { "Non" }
        TypeLicence             = $licenseType
        DateAffectationLicence  = $assignDate
        DerniereActivitePowerBI = $lastActivity
        UtilisateurInactif      = $inactive
    }
}

$report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

# ─────────────────────────────────────────────────────────
#  RÉSUMÉ
# ─────────────────────────────────────────────────────────
$totalUsers    = $report.Count
$inactiveCount = ($report | Where-Object { $_.UtilisateurInactif -eq "Oui" }).Count
$activeCount   = $totalUsers - $inactiveCount

Write-Host "`n=============================================" -ForegroundColor Green
Write-Host "  Rapport généré : $OutputPath"                   -ForegroundColor Green
Write-Host "  Total utilisateurs   : $totalUsers"             -ForegroundColor Green
Write-Host "  Actifs  (J-$ActivityDaysBack)  : $activeCount"  -ForegroundColor Green
Write-Host "  Inactifs (J-$ActivityDaysBack) : $inactiveCount" -ForegroundColor Yellow
Write-Host "=============================================" -ForegroundColor Green

$report | Format-Table -AutoSize
