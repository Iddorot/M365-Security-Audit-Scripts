# Setup Guide

This guide walks through everything you need before running the scripts in this
repository, including the SharePoint Online (SPO) audit scripts.

## 1. Prerequisites

- PowerShell 7+
- Appropriate Microsoft 365 admin permissions (SharePoint Admin / Global Reader
  as a minimum, depending on the script)
- Required PowerShell modules, per script family:
  - **SharePoint (`SharePoint/`)**: `PnP.PowerShell` (>= 2.4.0), `ImportExcel` (>= 7.8.0)
  - **Entra ID (`EntraID/`)**: `Microsoft.Graph` (`Microsoft.Graph.Authentication`,
    `Microsoft.Graph.Users`, `Microsoft.Graph.Identity.SignIns`)

Install the modules you need with:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

## 2. SharePoint Online audit setup

The SPO scripts (`SharePoint/Get-SiteInvetory.ps1`,
`SharePoint/Get-SitesExternalSharing.ps1`) connect using the **PnP PowerShell**
module, which requires an **Azure AD (Entra ID) app registration** — the PnP
Management Shell pattern — before you can connect to your tenant.

### 2.1 Register the PnP app in your tenant

Before running any SPO audit script, you must register (or grant admin consent
to) a PnP Azure AD application in your tenant:

1. Make sure `PnP.PowerShell` is installed:
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   ```
2. Run the built-in app registration cmdlet as a **Global Administrator**:
   ```powershell
   Register-PnPManagementShellAccess
   ```
   This registers the multi-tenant "PnP Management Shell" Azure AD application
   in your tenant and prompts you to sign in and consent.
3. Alternatively, register your **own** Azure AD app (recommended for
   production/least-privilege use):
   ```powershell
   Register-PnPAzureADApp -ApplicationName "SPO-Audit-Scripts" `
                           -Tenant "<tenantName>.onmicrosoft.com" `
                           -Interactive
   ```
   This creates a dedicated app registration with the SharePoint permissions
   needed for these scripts and outputs a **Client ID** (`AppId`).
4. In the [Azure Portal](https://portal.azure.com) → **App registrations**,
   confirm the app exists and that the required **SharePoint** API
   permissions (e.g. `Sites.FullControl.All`, delegated) have been
   **granted admin consent**.
5. Note the app's **Client ID (Application ID)** — you will need it for
   `config.json` in the next step.

> Without this app registration, `Connect-PnPOnline` will fail to
> authenticate against your tenant.

### 2.2 Configure `config.json`

The SPO scripts read their connection settings from a `config.json` file
located in the `SharePoint/` folder (next to the scripts), loaded automatically
by `SPO.Helpers.psm1`.

1. Copy the example file:
   ```powershell
   Copy-Item .\SharePoint\config.example.json .\SharePoint\config.json
   ```
2. Edit `SharePoint/config.json` and replace the placeholders:

   ```json
   {
     "adminUrl": "https://<tenantName>-admin.sharepoint.com",
     "tenantName": "<tenantName>",
     "pnpClientId": "<pnpClientID>"
   }
   ```

   | Key           | Description                                                                 |
   |---------------|-------------------------------------------------------------------------------|
   | `adminUrl`    | Your SharePoint Admin Center URL, e.g. `https://contoso-admin.sharepoint.com` |
   | `tenantName`  | Your tenant name only, e.g. `contoso` (without `.onmicrosoft.com`)            |
   | `pnpClientId` | The **Client ID (Application ID)** of the PnP app registered in step 2.1      |

3. `config.json` is git-ignored so your tenant details and client ID are never
   committed. Do not commit real values — only `config.example.json` should be
   tracked in the repository.

### 2.3 Run a script

```powershell
cd SharePoint
.\Get-SiteInvetory.ps1
# or
.\Get-SiteInvetory.ps1 -ConfigPath "C:\MyPath\config.json"
```

On first run you'll be prompted to interactively sign in and consent (unless
you configure certificate-based auth — see `Connect-SPOAdmin` in
`SPO.helpers.psm1` for `Certificate` / `CertificateFile` auth methods).

Reports are exported as `.xlsx` files to `%USERPROFILE%\Desktop\SPO-Reports`
by default.

---

## 3. Entra ID audit setup

The Entra ID scripts (`EntraID/`) connect via `Connect-MgGraph` and only
require the Microsoft Graph PowerShell SDK and the delegated scopes requested
by each script (e.g. `User.Read.All`, `UserAuthenticationMethod.Read.All`).
No app registration is required for interactive/delegated use — sign in with
an account that holds the necessary admin role when prompted.

---

## 4. Conditional Access audit setup

The `ConditionalAccess/` scripts also connect via Microsoft Graph
(`Connect-MgGraph`) and require an account with permissions to read
Conditional Access policies (e.g. `Policy.Read.All`).
