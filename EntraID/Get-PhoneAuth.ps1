# ============================================================
# Get-PhoneAuth.ps1
# Exports authentication methods for all users in the org to CSV
# ============================================================

# --- 1. Connect to Microsoft Graph ---
Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All" -NoWelcome

# --- 2. Settings ---
$OutputFile = "C:\Temp\AuthMethods_$(Get-Date -f yyyyMMdd_HHmm).csv"

# Authentication Method GUIDs
$MobilePhoneId     = "3179e48a-750b-4051-897c-87b9720928f7"
$OfficePhoneId     = "e37fc753-ff3b-4958-9484-eaa9425c82bc"
$AlternateMobileId = "b6332ec1-7057-4abe-9331-3d72feddfe41"

# --- 3. Fetch all users ---
Write-Host "Fetching users..." -ForegroundColor Cyan
$Users = Get-MgUser -All -Property "Id","DisplayName","UserPrincipalName","Mail"

$Report  = [System.Collections.Generic.List[Object]]::new()
$Counter = 0
$Total   = $Users.Count

# --- 4. Loop through each user ---
foreach ($User in $Users) {
    $Counter++
    $Percent = [math]::Round(($Counter / $Total) * 100)
    Write-Progress -Activity "Checking authentication methods" `
                   -Status "[$Counter/$Total] $($User.UserPrincipalName) — $Percent%" `
                   -PercentComplete $Percent

    # --- Fetch all phone methods in a single API call ---
    try {
        $AllPhoneMethods = Get-MgUserAuthenticationPhoneMethod -UserId $User.Id -ErrorAction Stop
    } catch {
        continue  # If the call fails entirely — skip the user
    }

    # --- Filter by GUID ---
    $Mobile    = $AllPhoneMethods | Where-Object { $_.Id -eq $MobilePhoneId }
    $Office    = $AllPhoneMethods | Where-Object { $_.Id -eq $OfficePhoneId }
    $AltMobile = $AllPhoneMethods | Where-Object { $_.Id -eq $AlternateMobileId }

    # If no phone methods found — skip the user entirely
    if (-not $Mobile -and -not $Office -and -not $AltMobile) {
        continue
    }

    # --- Build the report row only if at least one method exists ---
    $Row = [PSCustomObject]@{
        "Display Name"                = $User.DisplayName
        "Email"                       = if ($User.Mail) { $User.Mail } else { $User.UserPrincipalName }
        "Mobile Phone (3179e48a)"     = if ($Mobile)    { "Registered" } else { "" }
        "Mobile Phone Number"         = if ($Mobile)    { $Mobile.PhoneNumber }    else { "" }
        "Office Phone (e37fc753)"     = if ($Office)    { "Registered" } else { "" }
        "Office Phone Number"         = if ($Office)    { $Office.PhoneNumber }    else { "" }
        "Alternate Mobile (b6332ec1)" = if ($AltMobile) { "Registered" } else { "" }
        "Alternate Mobile Number"     = if ($AltMobile) { $AltMobile.PhoneNumber } else { "" }
        "User ID"                     = $User.Id
    }

    $Report.Add($Row)
}

Write-Progress -Activity "Checking authentication methods" -Completed

# --- 5. Export to CSV ---
if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" | Out-Null }
$Report | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Host "`n✅ Report saved to: $OutputFile" -ForegroundColor Green
Write-Host "Total users included in report: $($Report.Count) out of $Total" -ForegroundColor Yellow