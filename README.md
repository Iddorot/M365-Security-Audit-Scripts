# 🔐 M365-Security-Audit-Scripts

A growing collection of PowerShell scripts for auditing **Microsoft 365 security**,
**identity**, and **access configurations** across SharePoint, Entra ID, and Conditional Access.

---

## 📁 Repository Structure

```
M365-Security-Audit-Scripts/
├── SharePoint/             # SharePoint permission and access reports
│   ├── Get-SiteInvetory.ps1          # Full site inventory (sharing, permissions, links)
│   ├── Get-SitesExternalSharing.ps1  # External sharing capability report
│   ├── SPO.helpers.psm1              # Shared config/connection/export helpers
│   └── config.example.json           # Template for config.json (see Docs/setup-guide.md)
├── EntraID/                # Entra ID Reports
│   ├── Get-PhoneAuth.ps1              # Exports phone-based auth methods per user
│   └── Privileged_Roles_Assignment.ps1 # 🚧 planned — not yet implemented
├── ConditionalAccess/      # Conditional Access audits
│   └── CA_Policies_Audit.ps1          # 🚧 planned — not yet implemented
└── Docs/                   # Setup guides and documentation
    └── setup-guide.md                 # Full setup instructions, incl. PnP app registration
```

---

## 🚀 Getting Started

### Prerequisites
- PowerShell 7+
- Microsoft Graph PowerShell SDK (for EntraID / ConditionalAccess scripts)
- PnP PowerShell + ImportExcel modules (for SharePoint scripts)
- Appropriate Microsoft 365 admin permissions

### Setup

The SharePoint audit scripts require a **PnP app registration** in your tenant
and a local `config.json` before they can run.

👉 **See [Docs/setup-guide.md](Docs/setup-guide.md) for full step-by-step setup
instructions**, including how to register the PnP app and configure
`config.json`.

---

## ✅ Current Status

| Area               | Script                              | Status         |
|---------------------|--------------------------------------|----------------|
| SharePoint          | `Get-SiteInvetory.ps1`              | ✅ Working      |
| SharePoint          | `Get-SitesExternalSharing.ps1`      | ✅ Working      |
| Entra ID            | `Get-PhoneAuth.ps1`                 | ✅ Working      |
| Entra ID            | `Privileged_Roles_Assignment.ps1`   | 🚧 Not implemented (empty) |
| Conditional Access  | `CA_Policies_Audit.ps1`             | 🚧 Not implemented (empty) |

---

## ⚠️ Disclaimer

These scripts are provided **as-is** for auditing and informational purposes.
Always test in a **non-production environment** before running in production.
Ensure you have the appropriate permissions and approvals before executing any script.

---

## 📌 Roadmap

- [ ] Implement Entra ID privileged roles assignment audit
- [ ] Implement Conditional Access policies audit
- [ ] Add Exchange Online audits
- [ ] Add Defender for Cloud Apps audits
- [ ] Add output export to CSV / HTML reports
- [ ] Add scheduling guidance

---

## 🤝 Contributing

Contributions, ideas, and improvements are welcome!
Feel free to open an **Issue** or submit a **Pull Request**.

---

*🤖 Generated with the assistance of [Claude](https://claude.ai) by Anthropic*
