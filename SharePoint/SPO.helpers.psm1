<#
.SYNOPSIS
    SPO Shared Helper Module

.DESCRIPTION
    Provides reusable functions for all SPO scripts in this folder:
      - Config loading & validation  (adminUrl, tenantName, pnpClientId)
      - Module version checks
      - SharePoint connection management
      - Excel export
      - Logging

.USAGE
    At the top of any SPO script:
        Import-Module "$PSScriptRoot\SPO.Helpers.psm1" -Force
#>

#region ─── MODULE SETTINGS (shared defaults) ────────────────────────────────

$script:ModuleDefaults = @{
    PnPMinVersion    = "2.4.0"
    ExcelMinVersion  = "7.8.0"
    AuthMethod       = "Interactive"   # Interactive | Certificate | CertificateFile
    DateFormat       = "yyyyMMdd_HHmm"
    ReportDateFormat = "dd/MM/yyyy"
    OutputFolder     = "$env:USERPROFILE\Desktop\SPO-Reports"
    AutoOpenFile     = $true
    PageSize         = 500
}

#endregion

#region ─── LOGGING ──────────────────────────────────────────────────────────

<#
.SYNOPSIS
    Writes a formatted, color-coded log message to the console.

.PARAMETER Message
    The message to display.

.PARAMETER Level
    Log level: Info | Success | Warning | Error | Section | Debug

.EXAMPLE
    Write-SPOLog "Connected to tenant" -Level Success
    Write-SPOLog "Site not found"      -Level Warning
#>
function Write-SPOLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("Info","Success","Warning","Error","Section","Debug")]
        [string]$Level = "Info"
    )

    $timestamp = Get-Date -Format "HH:mm:ss"
    $prefix    = "[$timestamp]"

    switch ($Level) {
        "Info"    { Write-Host "$prefix ℹ️  $Message" -ForegroundColor Cyan   }
        "Success" { Write-Host "$prefix ✅ $Message"  -ForegroundColor Green  }
        "Warning" { Write-Host "$prefix ⚠️  $Message" -ForegroundColor Yellow }
        "Error"   { Write-Host "$prefix ❌ $Message"  -ForegroundColor Red    }
        "Section" {
            Write-Host ""
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
            Write-Host "  $Message"                                   -ForegroundColor Cyan
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
        }
        "Debug"   {
            if ($DebugPreference -ne "SilentlyContinue") {
                Write-Host "$prefix 🔍 $Message" -ForegroundColor Gray
            }
        }
    }
}

#endregion

#region ─── CONFIG ───────────────────────────────────────────────────────────

<#
.SYNOPSIS
    Loads and validates config.json.
    Returns adminUrl, tenantName and pnpClientId.

.PARAMETER ConfigPath
    Path to config.json. Defaults to config.json in the same folder as the module.

.OUTPUTS
    PSCustomObject with:
        .AdminUrl   - SharePoint Admin Center URL
        .TenantName - Tenant name (without .onmicrosoft.com)
        .ClientId   - App registration Client ID (pnpClientId)

.EXAMPLE
    $cfg = Get-SPOConfig
    $cfg = Get-SPOConfig -ConfigPath "C:\Scripts\config.json"
#>
function Get-SPOConfig {
    [CmdletBinding()]
    param(
        [string]$ConfigPath = "$PSScriptRoot\config.json"
    )

    Write-SPOLog "Loading config from: $ConfigPath" -Level Info

    # ── File exists? ──────────────────────────────────────────────────────
    if (-not (Test-Path $ConfigPath)) {
        Write-SPOLog "config.json not found at: $ConfigPath" -Level Error
        Write-SPOLog "Create a config.json file with the following content:" -Level Info
        Write-SPOLog (@"
  {
    "adminUrl"   : "https://YOUR-TENANT-admin.sharepoint.com",
    "tenantName" : "YOUR-TENANT",
    "pnpClientId": "YOUR-CLIENT-ID"
  }
"@) -Level Info
        throw "Missing config.json"
    }

    # ── Valid JSON? ───────────────────────────────────────────────────────
    try {
        $raw = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
        $cfg = $raw | ConvertFrom-Json
    } catch {
        Write-SPOLog "Failed to parse config.json: $_" -Level Error
        throw "Invalid config.json"
    }

    # ── Required keys: present and not placeholders? ──────────────────────
    $requiredKeys = @{
        "adminUrl"    = "YOUR-TENANT|YOUR-CLIENT-ID"
        "tenantName"  = "YOUR-TENANT|YOUR-CLIENT-ID"
        "pnpClientId" = "YOUR-TENANT|YOUR-CLIENT-ID"
    }

    foreach ($key in $requiredKeys.Keys) {
        $val = $cfg.$key
        if ([string]::IsNullOrWhiteSpace($val)) {
            Write-SPOLog "config.json is missing required key: '$key'" -Level Error
            throw "Missing key '$key' in config.json"
        }
        if ($val -match $requiredKeys[$key]) {
            Write-SPOLog "config.json still has a placeholder value for '$key'. Please update it." -Level Error
            throw "Placeholder not replaced for '$key' in config.json"
        }
    }

    Write-SPOLog "Tenant   : $($cfg.tenantName)"  -Level Success
    Write-SPOLog "URL      : $($cfg.adminUrl)"     -Level Success
    Write-SPOLog "ClientId : $($cfg.pnpClientId)"  -Level Success

    return [PSCustomObject]@{
        AdminUrl   = $cfg.adminUrl
        TenantName = $cfg.tenantName
        ClientId   = $cfg.pnpClientId
    }
}

#endregion

#region ─── MODULE VALIDATION ────────────────────────────────────────────────

<#
.SYNOPSIS
    Checks that required PowerShell modules are installed and meet minimum versions.

.PARAMETER Modules
    Hashtable of @{ ModuleName = "MinVersion" }

.EXAMPLE
    Assert-SPOModules -Modules @{
        "PnP.PowerShell" = "2.4.0"
        "ImportExcel"    = "7.8.0"
    }
#>
function Assert-SPOModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Modules
    )

    Write-SPOLog "Checking required modules..." -Level Info
    $allOk = $true

    foreach ($moduleName in $Modules.Keys) {
        $minVersion = $Modules[$moduleName]
        $installed  = Get-Module -ListAvailable -Name $moduleName |
                      Sort-Object Version -Descending |
                      Select-Object -First 1

        if (-not $installed) {
            Write-SPOLog "Module '$moduleName' is not installed. Run: Install-Module $moduleName -Scope CurrentUser" -Level Error
            $allOk = $false
        } elseif ([version]$installed.Version -lt [version]$minVersion) {
            Write-SPOLog "Module '$moduleName' v$($installed.Version) is below minimum v$minVersion." -Level Warning
        } else {
            Write-SPOLog "$moduleName v$($installed.Version)" -Level Success
        }
    }

    if (-not $allOk) { throw "One or more required modules are missing." }
}

#endregion

#region ─── CONNECTION ───────────────────────────────────────────────────────

<#
.SYNOPSIS
    Connects to the SharePoint Online Admin Center.

.DESCRIPTION
    ClientId is automatically taken from the Config object (loaded from config.json).
    No need to pass ClientId manually in any calling script.

.PARAMETER Config
    PSCustomObject returned by Get-SPOConfig.

.PARAMETER AuthMethod
    Authentication method: Interactive | Certificate | CertificateFile
    Defaults to "Interactive".

.PARAMETER Thumbprint
    Certificate thumbprint. Required only for Certificate auth.

.PARAMETER CertificatePath
    Path to .pfx file. Required only for CertificateFile auth.

.OUTPUTS
    PnP connection object. Pass to all subsequent PnP cmdlets via -Connection.

.EXAMPLE
    $conn = Connect-SPOAdmin -Config $cfg
    $conn = Connect-SPOAdmin -Config $cfg -AuthMethod Certificate -Thumbprint "ABC123"
#>
function Connect-SPOAdmin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [ValidateSet("Interactive","Certificate","CertificateFile")]
        [string]$AuthMethod = $script:ModuleDefaults.AuthMethod,

        [string]$Thumbprint      = "",
        [string]$CertificatePath = "",
        [System.Security.SecureString]$CertificatePassword = $null
    )

    Write-SPOLog "Connecting to $($Config.AdminUrl) using [$AuthMethod]..." -Level Info

    try {
        $conn = switch ($AuthMethod) {
            "Interactive" {
                Connect-PnPOnline -Url      $Config.AdminUrl `
                                  -ClientId $Config.ClientId `
                                  -Interactive `
                                  -ReturnConnection
            }
            "Certificate" {
                Connect-PnPOnline -Url        $Config.AdminUrl `
                                  -ClientId   $Config.ClientId `
                                  -Thumbprint $Thumbprint `
                                  -Tenant     "$($Config.TenantName).onmicrosoft.com" `
                                  -ReturnConnection
            }
            "CertificateFile" {
                if ($CertificatePassword) {
                    Connect-PnPOnline -Url                 $Config.AdminUrl `
                                      -ClientId            $Config.ClientId `
                                      -CertificatePath     $CertificatePath `
                                      -CertificatePassword $CertificatePassword `
                                      -Tenant              "$($Config.TenantName).onmicrosoft.com" `
                                      -ReturnConnection
                } else {
                    Connect-PnPOnline -Url             $Config.AdminUrl `
                                      -ClientId        $Config.ClientId `
                                      -CertificatePath $CertificatePath `
                                      -Tenant          "$($Config.TenantName).onmicrosoft.com" `
                                      -ReturnConnection
                }
            }
        }

        Write-SPOLog "Connected to tenant: $($Config.TenantName)" -Level Success
        return $conn

    } catch {
        Write-SPOLog "Connection failed: $_" -Level Error
        throw
    }
}

<#
.SYNOPSIS
    Connects to an individual SharePoint site collection.

.DESCRIPTION
    ClientId is automatically taken from the Config object (loaded from config.json).
    No need to pass ClientId manually in any calling script.

.PARAMETER SiteUrl
    Full URL of the target site collection.

.PARAMETER Config
    PSCustomObject returned by Get-SPOConfig.

.PARAMETER AuthMethod
    Authentication method: Interactive | Certificate | CertificateFile
    Defaults to "Interactive".

.PARAMETER Thumbprint
    Certificate thumbprint. Required only for Certificate auth.

.PARAMETER CertificatePath
    Path to .pfx file. Required only for CertificateFile auth.

.OUTPUTS
    PnP connection object, or $null if the connection failed.

.EXAMPLE
    $siteConn = Connect-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/HR" -Config $cfg
#>
function Connect-SPOSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [ValidateSet("Interactive","Certificate","CertificateFile")]
        [string]$AuthMethod = $script:ModuleDefaults.AuthMethod,

        [string]$Thumbprint      = "",
        [string]$CertificatePath = "",
        [System.Security.SecureString]$CertificatePassword = $null
    )

    try {
        $conn = switch ($AuthMethod) {
            "Interactive" {
                Connect-PnPOnline -Url      $SiteUrl `
                                  -ClientId $Config.ClientId `
                                  -Interactive `
                                  -ReturnConnection
            }
            "Certificate" {
                Connect-PnPOnline -Url        $SiteUrl `
                                  -ClientId   $Config.ClientId `
                                  -Thumbprint $Thumbprint `
                                  -Tenant     "$($Config.TenantName).onmicrosoft.com" `
                                  -ReturnConnection
            }
            "CertificateFile" {
                if ($CertificatePassword) {
                    Connect-PnPOnline -Url                 $SiteUrl `
                                      -ClientId            $Config.ClientId `
                                      -CertificatePath     $CertificatePath `
                                      -CertificatePassword $CertificatePassword `
                                      -Tenant              "$($Config.TenantName).onmicrosoft.com" `
                                      -ReturnConnection
                } else {
                    Connect-PnPOnline -Url             $SiteUrl `
                                      -ClientId        $Config.ClientId `
                                      -CertificatePath $CertificatePath `
                                      -Tenant          "$($Config.TenantName).onmicrosoft.com" `
                                      -ReturnConnection
                }
            }
        }
        return $conn

    } catch {
        Write-SPOLog "Failed to connect to site: $SiteUrl — $_" -Level Warning
        return $null
    }
}

<#
.SYNOPSIS
    Safely disconnects a per-site PnP connection without throwing errors.

.PARAMETER Connection
    The PnP connection object returned by Connect-SPOSite.

.EXAMPLE
    Disconnect-SPOSite -Connection $siteConn
#>
function Disconnect-SPOSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Connection
    )
    try {
        Disconnect-PnPOnline -Connection $Connection -ErrorAction SilentlyContinue
    } catch {}
}

<#
.SYNOPSIS
    Disconnects the main Admin Center PnP session.

.EXAMPLE
    Disconnect-SPOAdmin
#>
function Disconnect-SPOAdmin {
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-SPOLog "Disconnected from SharePoint Online." -Level Info
    } catch {}
}

#endregion

#region ─── EXCEL EXPORT ─────────────────────────────────────────────────────

<#
.SYNOPSIS
    Exports a data collection to a formatted Excel (.xlsx) file.

.PARAMETER Data
    List of PSCustomObjects to export.

.PARAMETER ReportName
    Used to build the filename: {ReportName}_{timestamp}.xlsx

.PARAMETER WorksheetName
    Excel worksheet tab name. Default: "Report".

.PARAMETER TableName
    Excel table name (no spaces). Default: "ReportTable".

.PARAMETER TableStyle
    Excel table style. Default: "Medium9".

.PARAMETER OutputFolder
    Destination folder. Created automatically if it does not exist.
    Default: Desktop\SPO-Reports.

.PARAMETER AutoOpen
    Opens the file in Excel after export. Default: $true.

.PARAMETER ConditionalFormats
    Array of New-ConditionalText objects for conditional formatting.

.OUTPUTS
    Full path of the saved .xlsx file.

.EXAMPLE
    $path = Export-SPOExcel -Data $report -ReportName "SiteInventory"

    $path = Export-SPOExcel -Data $report `
                            -ReportName         "SiteInventory" `
                            -WorksheetName      "Sites" `
                            -OutputFolder       "C:\Reports" `
                            -AutoOpen           $false `
                            -ConditionalFormats @(
                                New-ConditionalText "Anyone" -BackgroundColor LightCoral
                            )
#>
function Export-SPOExcel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[PSCustomObject]]$Data,

        [Parameter(Mandatory)]
        [string]$ReportName,

        [string]$WorksheetName      = "Report",
        [string]$TableName          = "ReportTable",
        [string]$TableStyle         = "Medium9",
        [string]$OutputFolder       = $script:ModuleDefaults.OutputFolder,
        [bool]  $AutoOpen           = $script:ModuleDefaults.AutoOpenFile,
        [array] $ConditionalFormats = @()
    )

    Write-SPOLog "Exporting to Excel..." -Level Info

    # Ensure output folder exists
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        Write-SPOLog "Created output folder: $OutputFolder" -Level Info
    }

    $timestamp  = Get-Date -Format $script:ModuleDefaults.DateFormat
    $outputPath = Join-Path $OutputFolder "$ReportName`_$timestamp.xlsx"

    $excelParams = @{
        Path          = $outputPath
        WorksheetName = $WorksheetName
        TableName     = $TableName
        TableStyle    = $TableStyle
        AutoSize      = $true
        FreezeTopRow  = $true
        BoldTopRow    = $true
        AutoFilter    = $true
    }

    if ($ConditionalFormats.Count -gt 0) {
        $excelParams["ConditionalText"] = $ConditionalFormats
    }

    try {
        $Data | Export-Excel @excelParams
        Write-SPOLog "Report saved: $outputPath" -Level Success
    } catch {
        Write-SPOLog "Failed to export Excel: $_" -Level Error
        throw
    }

    if ($AutoOpen) { Start-Process $outputPath }

    return $outputPath
}

#endregion

#region ─── SUMMARY ──────────────────────────────────────────────────────────

<#
.SYNOPSIS
    Prints a formatted summary block at the end of any SPO script.

.PARAMETER TenantName
    Tenant name shown in the summary.

.PARAMETER ReportPath
    Path to the generated report file.

.PARAMETER Stats
    Ordered hashtable of label → value rows to display.

.EXAMPLE
    Write-SPOSummary -TenantName "contoso" `
                     -ReportPath $outputPath `
                     -Stats ([ordered]@{
                         "Sites processed" = $allSites.Count
                         "Errors"          = $errors
                     })
#>
function Write-SPOSummary {
    [CmdletBinding()]
    param(
        [string]$TenantName,
        [string]$ReportPath,
        [System.Collections.Specialized.OrderedDictionary]$Stats = [ordered]@{}
    )

    Write-Host ""
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
    Write-Host "  📊 REPORT SUMMARY"                         -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
    Write-Host "  Tenant : $TenantName"
    Write-Host "  Date   : $(Get-Date -Format 'dd/MM/yyyy HH:mm')"

    foreach ($key in $Stats.Keys) {
        Write-Host "  $key : $($Stats[$key])"
    }

    if ($ReportPath) {
        Write-Host "  Report : $ReportPath" -ForegroundColor Yellow
    }

    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkCyan
}

#endregion

#region ─── EXPORTS ──────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    "Write-SPOLog"
    "Get-SPOConfig"
    "Assert-SPOModules"
    "Connect-SPOAdmin"
    "Connect-SPOSite"
    "Disconnect-SPOSite"
    "Disconnect-SPOAdmin"
    "Export-SPOExcel"
    "Write-SPOSummary"
)

#endregion