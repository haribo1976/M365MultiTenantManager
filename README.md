# M365 Multi-Tenant Manager

A PowerShell + WPF application for managing multiple Microsoft 365 tenants from a single interface.

Based on the architecture pattern of [Micke-K/IntuneManagement](https://github.com/Micke-K/IntuneManagement).

## Features

- **Multi-Tenant Management**: Register and switch between multiple M365 tenants
- **User & Group Inventory**: View users and groups across all tenants
- **License Management**: Aggregate license view with comparison
- **Security Management**: Conditional Access policy export/import/compare
- **Secure Score**: View and compare Secure Scores across tenants
- **Bulk Export**: Export tenant configurations to JSON for backup
- **Batch Mode**: Silent execution for DevOps automation

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or PowerShell 7.x (recommended)
- MSAL.PS module for authentication

## Quick Start

### 1. Install Prerequisites

```powershell
# Install MSAL.PS module
Install-Module MSAL.PS -Scope CurrentUser
```

### 2. Launch the Application

```powershell
# Option 1: Double-click Start.cmd

# Option 2: Run from PowerShell
.\Start-M365Manager.ps1

# Option 3: PowerShell 7 (recommended)
pwsh.exe -File .\Start-M365Manager.ps1
```

### 3. Add Your First Tenant

1. Click the **+** button next to the tenant selector
2. Enter your tenant ID or domain (e.g., `contoso.onmicrosoft.com`)
3. Authenticate with Global Admin credentials
4. Grant consent for the required permissions

## Authentication

The application supports multiple authentication methods:

### Interactive (Default)
```powershell
Connect-M365Tenant -TenantId "contoso.onmicrosoft.com"
```

### Device Code (Headless)
```powershell
Connect-M365Tenant -TenantId "tenant-guid" -UseDeviceCode
```

### App Credentials (Automation)
```powershell
Connect-M365Tenant -TenantId "tenant-guid" -AppId "app-guid" -Secret "client-secret"
```

### Certificate (Production)
```powershell
Connect-M365Tenant -TenantId "tenant-guid" -AppId "app-guid" -CertThumbprint "thumbprint"
```

## Command Reference

### Tenant Management

```powershell
# List registered tenants
Get-M365RegisteredTenants

# Register a new tenant
Register-M365Tenant -TenantId "contoso.com" -FriendlyName "Contoso Production" -Tags @("production", "uk")

# Switch tenant context
Switch-M365Tenant -TenantId "tenant-guid"

# Get tenant information
Get-M365TenantInfo
Get-M365TenantHealth
Get-M365TenantDomains
```

### User Management

```powershell
# Get users
Get-M365Users
Get-M365Users -All -GuestsOnly
Get-M365User -UserId "user@contoso.com"

# Get groups
Get-M365Groups
Get-M365Groups -Type Security
Get-M365GroupMembers -GroupId "group-guid"

# Export
Export-M365Users -Path ".\users.csv" -Format CSV
Export-M365Groups -Path ".\groups.json" -Format JSON -IncludeMembers
```

### License Management

```powershell
# Get licenses
Get-M365Licenses
Get-M365LicenseSummary -AllTenants

# Compare licenses
Compare-M365Licenses -SourceTenantId "guid1" -TargetTenantId "guid2"

# Export report
Export-M365LicenseReport -Path ".\licenses.csv" -AllTenants
```

### Security Management

```powershell
# Conditional Access
Get-M365ConditionalAccessPolicies
Export-M365ConditionalAccessPolicy -PolicyId "All" -Path ".\CA-Policies"
Import-M365ConditionalAccessPolicy -Path ".\policy.json" -State "disabled"
Compare-M365ConditionalAccessPolicies -SourceTenantId "guid1" -TargetTenantId "guid2"

# Secure Score
Get-M365SecureScore
Get-M365SecureScoreControlProfiles
Compare-M365SecureScores
```

### Bulk Operations

```powershell
# Export all data from current tenant
Export-M365TenantData -Scope All -OutputPath ".\Export"

# Export from all registered tenants
Export-M365AllTenants -Scope @("Users", "Licenses", "ConditionalAccess")

# Create migration table
New-M365MigrationTable -SourceTenantId "guid1" -TargetTenantId "guid2"
```

## Batch/Silent Mode

For automation and DevOps integration:

```powershell
.\Start-M365Manager.ps1 -Silent `
    -SilentBatchFile ".\Config\BatchJob.json" `
    -TenantId "tenant-guid" `
    -AppId "app-guid" `
    -Secret "client-secret"
```

### Azure DevOps Pipeline Example

```yaml
- task: PowerShell@2
  displayName: 'Export M365 Configuration'
  inputs:
    targetType: 'filePath'
    filePath: './M365Manager/Start-M365Manager.ps1'
    arguments: >
      -Silent
      -SilentBatchFile '$(Build.SourcesDirectory)/batch-export.json'
      -TenantId '$(TenantId)'
      -AppId '$(AppId)'
      -Secret '$(AppSecret)'
```

## Configuration

### Settings (Config/Settings.json)

| Setting | Description | Default |
|---------|-------------|---------|
| `DefaultAppId` | Microsoft Graph PowerShell app ID | `14d82eec-...` |
| `CustomAppId` | Your custom app registration | `null` |
| `GraphVersion` | API version (v1.0 or beta) | `v1.0` |
| `LogLevel` | Debug, Information, Warning, Error | `Information` |
| `ExportPath` | Default export directory | `.\Export` |
| `PageSize` | Items per page in API calls | `100` |

### Custom App Registration

For production use, register your own multi-tenant app:

1. Go to Azure Portal > App Registrations
2. Create new registration
3. Set "Supported account types" to "Accounts in any organizational directory"
4. Add required API permissions (see below)
5. Create client secret or certificate
6. Update `CustomAppId` in Settings.json

### Required Permissions

**Delegated (for interactive use):**
- User.Read.All
- Group.Read.All
- Directory.Read.All
- Policy.Read.All
- SecurityEvents.Read.All
- AuditLog.Read.All

**Application (for automation):**
- Directory.Read.All
- Policy.Read.All
- SecurityEvents.Read.All

## Project Structure

```
M365MultiTenantManager/
├── Start-M365Manager.ps1      # Entry point
├── Start.cmd                  # Windows launcher
├── Core/
│   ├── Logging.psm1          # CMTrace-compatible logging
│   ├── Settings.psm1         # Configuration management
│   ├── Authentication.psm1   # MSAL token handling
│   ├── GraphAPI.psm1         # Graph API wrapper
│   ├── Core.psm1             # Core functions
│   └── UI/
│       ├── MainWindow.xaml   # WPF interface
│       └── MainWindow.xaml.ps1
├── Extensions/
│   ├── TenantManagement/
│   ├── UserManagement/
│   ├── LicenseManagement/
│   ├── SecurityManagement/
│   └── Documentation/
├── Config/
│   ├── Settings.json
│   ├── Tenants.json
│   └── BatchJob-Example.json
├── Export/                    # Export output
└── Logs/                      # Log files
```

## Logging

Logs are written in CMTrace format to the `Logs` folder. View with:
- CMTrace (SCCM toolkit)
- OneTrace (Configuration Manager)
- Any text editor

## Troubleshooting

### Authentication Issues

1. Ensure MSAL.PS module is installed: `Get-Module MSAL.PS -ListAvailable`
2. Clear token cache: `Disconnect-M365Tenant -All`
3. Check if consent was granted in target tenant

### Graph API Errors

1. Check log files for detailed error messages
2. Verify permissions are granted in the tenant
3. Some endpoints require beta API (handled automatically)

### UI Not Loading

1. Run `Start-WithConsole.cmd` to see errors
2. Verify WPF assemblies are available
3. Try PowerShell 7: `Start-PS7.cmd`

## Contributing

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - See LICENSE file

## Acknowledgments

- [Micke-K/IntuneManagement](https://github.com/Micke-K/IntuneManagement) for the architecture pattern
- Microsoft Graph API documentation
- MSAL.PS module maintainers
