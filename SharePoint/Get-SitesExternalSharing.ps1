<#
.SYNOPSIS
    SharePoint Online - External Sharing Report

.DESCRIPTION
    Queries all site collections and reports their external sharing capability.
    Exports results to a formatted Excel file.
    ClientId is read automatically from config.json via SPO.Helpers.psm1.

.USAGE
    .\SPO_ExternalSharing.ps1
    .\SPO_ExternalSharing.ps1 -ConfigPath "C:\MyPath\config.json"
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "$PSScriptRoot\config.json"
)

#region ─── SETTINGS ─────────────────────────────────────────────────────────
$AuthMethod   = "Interactive"   # Interactive | Certificate | CertificateFile
$OutputFolder = "$env:USERPROFILE\Desktop\SPO-Reports"

$ConditionalFormats = @(
    New-ConditionalText "ExternalUserAndGuestSharing" -BackgroundColor LightCoral  -ConditionalTextColor DarkRed
    New-ConditionalText "ExternalUserSharingOnly"     -BackgroundColor LightYellow -ConditionalTextColor DarkOrange
    New-ConditionalText "Disabled"                    -BackgroundColor LightGreen  -ConditionalTextColor DarkGreen
)
#endregion

#region ─── BOOTSTRAP ────────────────────────────────────────────────────────
Import-Module "$PSScriptRoot\SPO.Helpers.psm1" -Force

Assert-SPOModules -Modules @{
    "PnP.PowerShell" = "2.4.0"
    "ImportExcel"    = "7.8.0"
}

# ClientId is read from config.json automatically — no need to pass it here
$cfg  = Get-SPOConfig -ConfigPath $ConfigPath
$conn = Connect-SPOAdmin -Config $cfg -AuthMethod $AuthMethod
#endregion

#region ─── GET SITES ────────────────────────────────────────────────────────
Write-SPOLog "Fetching all site collections..." -Level Section

$sites = Get-PnPTenantSite -Connection $conn

Write-SPOLog "Found $($sites.Count) sites. Re-querying each for accurate sharing values..." -Level Success
#endregion

#region ─── MAIN LOOP ────────────────────────────────────────────────────────
$report  = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter = 0
$errors  = 0

foreach ($s in $sites) {
    $counter++
    $pct = [math]::Round(($counter / $sites.Count) * 100, 1)

    Write-Progress -Activity "🔄 Querying Sites ($pct%)" `
                   -Status "$counter of $($sites.Count) — $($s.Url)" `
                   -PercentComplete $pct

    Write-SPOLog "[$counter/$($sites.Count)] $($s.Url)" -Level Debug

    try {
        # Re-query each site individually for accurate SharingCapability
        $site = Get-PnPTenantSite -Identity $s.Url -Connection $conn

        $report.Add([PSCustomObject]@{
            "Title"              = $site.Title
            "URL"                = $site.Url
            "Template"           = $site.Template
            "Sharing Capability" = $site.SharingCapability
            "Storage Usage (MB)" = $site.StorageUsageCurrent
            "Report Date"        = Get-Date -Format "dd/MM/yyyy"
        })

    } catch {
        Write-SPOLog "Error querying $($s.Url) : $_" -Level Warning
        $errors++
    }
}

Write-Progress -Activity "Querying Sites" -Completed
Write-SPOLog "Processing complete." -Level Success
#endregion

#region ─── DISPLAY IN CONSOLE ───────────────────────────────────────────────
Write-SPOLog "Results sorted by Sharing Capability:" -Level Section

$report | Sort-Object "Sharing Capability" | Format-Table -AutoSize
#endregion

#region ─── EXPORT & SUMMARY ─────────────────────────────────────────────────
$outputPath = Export-SPOExcel -Data $report `
                              -ReportName    "SPO_ExternalSharing" `
                              -WorksheetName "External Sharing" `
                              -TableName     "ExternalSharing" `
                              -OutputFolder  $OutputFolder `
                              -ConditionalFormats $ConditionalFormats

Write-SPOSummary -TenantName $cfg.TenantName `
                 -ReportPath $outputPath `
                 -Stats ([ordered]@{
                     "Sites processed" = $sites.Count
                     "Errors"          = $errors
                 })

Disconnect-SPOAdmin
#endregion