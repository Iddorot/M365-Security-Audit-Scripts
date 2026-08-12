<#
.SYNOPSIS
    SharePoint Online - Site Inventory Report

.DESCRIPTION
    Collects governance, sharing, permissions and link metadata
    for every site in the tenant and exports to a formatted Excel file.
    All headers are in English.
    ClientId is read automatically from config.json via SPO.Helpers.psm1.

.COLUMNS
    Site Name | Site ID | Site URL | Site Template | Primary Admin
    Primary Admin Email | External Sharing | Site Privacy | Site Sensitivity
    Number of users having access | Guest user permissions
    External participant permissions | Entra group count | File count
    Items with unique permissions count | PeopleInYourOrg link count
    Anyone link count | EEEU permission count | Everyone permission count
    Report date | Notes | Status

.USAGE
    .\Get-SiteInvetory.ps1.ps1
    .\Get-SiteInvetory.ps1.ps1 -ConfigPath "C:\MyPath\config.json"
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$PSScriptRoot\config.json"
)

#region ─── SETTINGS ─────────────────────────────────────────────────────────
$AuthMethod             = "Interactive"   # Interactive | Certificate | CertificateFile
$OutputFolder           = "$env:USERPROFILE\Desktop\SPO-Reports"
$ExcludeSystemSites     = $true
$ExcludeOneDrive        = $true
$ExcludeTemplates       = @("SRCHCEN#0", "SPSMSITEHOST#0", "REDIRECTSITE#0", "APP#0")
$ExcludeUrlPatterns     = @("-my\.sharepoint\.com")
$IncludeSiteUrlFilter   = ""             # e.g. "*marketing*" — leave empty for all sites
$MaxListsForUniquePerms = 10             # Max doc libs checked for unique perms per site
$PageSize               = 500            # Items per page when querying list items
$DelayBetweenSitesMs    = 0             # ms delay between sites (increase to reduce throttling)
$ReportDateFormat       = "dd/MM/yyyy"

$ConditionalFormats = @(
    New-ConditionalText "Anyone"      -BackgroundColor LightCoral  -ConditionalTextColor DarkRed
    New-ConditionalText "Existing"    -BackgroundColor LightYellow -ConditionalTextColor DarkOrange
    New-ConditionalText "Only People" -BackgroundColor LightGreen  -ConditionalTextColor DarkGreen
)
#endregion

#region ─── BOOTSTRAP ────────────────────────────────────────────────────────
Import-Module "$PSScriptRoot\SPO.Helpers.psm1" -Force

Assert-SPOModules -Modules @{
    "PnP.PowerShell" = "2.4.0"
    "ImportExcel"    = "7.8.0"
}
#region ─── GET ALL SITES ────────────────────────────────────────────────────
Write-SPOLog "Fetching all site collections..." -Level Section

$getSiteParams = @{ Detailed = $true }
if ($ExcludeOneDrive)             { $getSiteParams["IncludeOneDriveSites"] = $false }
if ($IncludeSiteUrlFilter -ne "") { $getSiteParams["Filter"] = "Url -like '$IncludeSiteUrlFilter'" }

$cfg  = Get-SPOConfig -ConfigPath $ConfigPath
$conn = Connect-SPOAdmin -Config $cfg -AuthMethod $AuthMethod

$allSites = Get-PnPTenantSite @getSiteParams -Connection $conn | Where-Object {
    if (-not $ExcludeSystemSites) { return $true }
    $templateOk = -not ($ExcludeTemplates  | Where-Object { $_.Template -like "*$_*" })
    $urlOk      = -not ($ExcludeUrlPatterns | Where-Object { $_.Url      -match $_   })
    $templateOk -and $urlOk
}

Write-SPOLog "Found $($allSites.Count) sites to process." -Level Success
#endregion

#region ─── HELPERS (local) ──────────────────────────────────────────────────

function Get-SharingLabel([string]$cap) {
    switch ($cap) {
        "ExternalUserAndGuestSharing"         { return "Anyone"                      }
        "ExternalUserSharingOnly"             { return "New and Existing Guests"     }
        "ExistingExternalUserSharingOnly"     { return "Existing Guests Only"        }
        "Disabled"                            { return "Only People in Organization" }
        default                               { return $cap                          }
    }
}

function Get-PrivacyLabel($groupId) {
    if ($null -eq $groupId -or $groupId -eq [guid]::Empty) { return "No M365 Group" }
    return "M365 Group Site"
}

function Get-SensitivityLabel($label) {
    if ($label -and $label -ne [guid]::Empty) { return $label.ToString() }
    return "None"
}

#endregion

#region ─── MAIN LOOP ────────────────────────────────────────────────────────
Write-SPOLog "Starting site processing..." -Level Section

$report     = [System.Collections.Generic.List[PSCustomObject]]::new()
$reportDate = Get-Date -Format $ReportDateFormat
$counter    = 0
$errors     = 0

foreach ($site in $allSites) {
    $counter++
    $pct = [math]::Round(($counter / $allSites.Count) * 100, 1)

    Write-Progress -Activity "🔄 Processing Sites ($pct%)" `
                   -Status "$counter of $($allSites.Count) — $($site.Url)" `
                   -PercentComplete $pct

    Write-SPOLog "[$counter/$($allSites.Count)] $($site.Url)" -Level Debug

    #── Static properties from tenant-level call ──────────────────────────
    $siteName         = $site.Title
    $siteId           = $site.Id
    $siteUrl          = $site.Url
    $template         = $site.Template
    $owner            = $site.Owner
    $ownerEmail       = $site.Owner
    $sharing          = Get-SharingLabel  $site.SharingCapability
    $privacy          = Get-PrivacyLabel  $site.GroupId
    $sensitivity      = Get-SensitivityLabel $site.SensitivityLabel

    #── Per-site metric defaults ───────────────────────────────────────────
    $userCount           = "N/A"
    $guestPerms          = "N/A"
    $externalParticipant = "N/A"
    $entraGroupCount     = "N/A"
    $fileCount           = "N/A"
    $uniquePermsCount    = "N/A"
    $peopleInOrgLinks    = "Requires Graph API"
    $anyoneLinks         = "Requires Graph API"
    $eeeuCount           = "N/A"
    $everyoneCount       = "N/A"


    if ( $conn ) {
        try {
            #── Users ─────────────────────────────────────────────────────
            $allUsers = Get-PnPUser -Connection  $conn  `
                                    -WithRightsAssigned `
                                    -ErrorAction SilentlyContinue

            $userCount = ($allUsers | Where-Object {
                $_.PrincipalType -eq "User"
            }).Count

            #── Guest users ────────────────────────────────────────────────
            $guestPerms = ($allUsers | Where-Object {
                $_.LoginName -match "#ext#" -or
                $_.LoginName -match "urn:spo:guest"
            }).Count

            #── External participants ──────────────────────────────────────
            $externalParticipant = ($allUsers | Where-Object {
                $_.LoginName -match "urn:spo:anon" -or
                $_.IsEmailAuthenticationGuestUser -eq $true
            }).Count

            #── Entra group count ──────────────────────────────────────────
            $siteGroups      = Get-PnPSiteGroup -Connection $conn -ErrorAction SilentlyContinue
            $entraGroupCount = ($siteGroups | Where-Object {
                $_.LoginName -match "c:0t\.c\|tenant\|" -or
                $_.LoginName -match "c:0o\.c\|federateddirectoryclaimprovider\|"
            }).Count

            #── File count (all non-hidden document libraries) ─────────────
            $docLibs   = Get-PnPList -Connection $conn |
                         Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false }
            $fileCount = ($docLibs | Measure-Object -Property ItemCount -Sum).Sum
            if (-not $fileCount) { $fileCount = 0 }

            #── Items with unique permissions ──────────────────────────────
            $uniquePermsCount = 0
            foreach ($list in $docLibs) {
                try {
                    $uniquePermsCount += (
                        Get-PnPListItem -List $list `
                                        -Connection $conn `
                                        -PageSize $PageSize `
                                        -ErrorAction SilentlyContinue |
                        Where-Object { $_.HasUniqueRoleAssignments -eq $true }
                    ).Count
                } catch {}
            }

            #── EEEU / Everyone permission counts ─────────────────────────
            $eeeuCount     = 0
            $everyoneCount = 0
            foreach ($list in ($docLibs | Select-Object -First $MaxListsForUniquePerms)) {
                try {
                    $ras = Get-PnPProperty -ClientObject $list `
                                           -Property RoleAssignments `
                                           -Connection $conn
                    foreach ($ra in $ras) {
                        Get-PnPProperty -ClientObject $ra -Property Member -Connection $conn | Out-Null
                        if ($ra.Member.LoginName -match "spo-grid-all-users") { $eeeuCount++     }
                        if ($ra.Member.LoginName -match "c:0\(\.s\|true")     { $everyoneCount++ }
                    }
                } catch {}
            }

            # PeopleInYourOrg / Anyone link counts require Microsoft Graph API
            # Endpoint: GET /v1.0/sites/{siteId}/permissions
            # Uncomment and implement below once Graph access is available:
            #
            # $graphSiteId      = (Get-PnPSite -Connection $conn).Id
            # $permissions      = Invoke-PnPGraphMethod -Url "sites/$graphSiteId/permissions" -Connection $conn
            # $peopleInOrgLinks = ($permissions.value | Where-Object { $_.link.scope -eq "organization" }).Count
            # $anyoneLinks      = ($permissions.value | Where-Object { $_.link.scope -eq "anonymous"    }).Count

        } catch {
            Write-SPOLog "Error processing $siteUrl : $_" -Level Warning
            $errors++
        } finally {
            Disconnect-SPOSite -Connection $conn
        }
    } else {
        Write-SPOLog "Skipped (could not connect): $siteUrl" -Level Warning
        $errors++
    }

    if ($DelayBetweenSitesMs -gt 0) { Start-Sleep -Milliseconds $DelayBetweenSitesMs }

    #── Build report row ──────────────────────────────────────────────────
    $report.Add([PSCustomObject]@{
        "Site Name"                          = $siteName
        "Site ID"                            = $siteId
        "Site URL"                           = $siteUrl
        "Site Template"                      = $template
        "Primary Admin"                      = $owner
        "Primary Admin Email"                = $ownerEmail
        "External Sharing"                   = $sharing
        "Site Privacy"                       = $privacy
        "Site Sensitivity"                   = $sensitivity
        "Number of users having access"      = $userCount
        "Guest user permissions"             = $guestPerms
        "External participant permissions"   = $externalParticipant
        "Entra group count"                  = $entraGroupCount
        "File count"                         = $fileCount
        "Items with unique permissions count"= $uniquePermsCount
        "PeopleInYourOrg link count"         = $peopleInOrgLinks
        "Anyone link count"                  = $anyoneLinks
        "EEEU permission count"              = $eeeuCount
        "Everyone permission count"          = $everyoneCount
        "Report date"                        = $reportDate
        "Notes"                              = ""
        "Status"                             = ""
    })
}

Write-Progress -Activity "Processing Sites" -Completed
Write-SPOLog "Processing complete. Sites: $($allSites.Count) | Errors: $errors" -Level Success
#endregion

#region ─── EXPORT & SUMMARY ─────────────────────────────────────────────────
$outputPath = Export-SPOExcel -Data $report `
                              -ReportName    "SPO_SiteInventory" `
                              -WorksheetName "Site Inventory" `
                              -TableName     "SiteInventory" `
                              -OutputFolder  $OutputFolder `
                              -ConditionalFormats $ConditionalFormats

Write-SPOSummary -TenantName $cfg.TenantName `
                 -ReportPath $outputPath `
                 -Stats ([ordered]@{
                     "Sites processed" = $allSites.Count
                     "Errors"          = $errors
                 })

Disconnect-SPOAdmin
#endregion