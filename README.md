# contractor-expiration-alert

PowerShell script that sends a daily HTML digest email listing contractor AD accounts expiring within a configurable window. Email is delivered via Microsoft Graph (no SMTP relay required). Intended to run as a scheduled task.

## Requirements

- PowerShell 5.1 or later
- ActiveDirectory module (RSAT: Active Directory Domain Services)
- Domain-joined machine with network access to a domain controller
- Entra ID app registration with Mail.Send application permission (see Prerequisites)

## Prerequisites

Before first use, register an application in Entra ID:

1. Register an application in Entra ID (any name, no redirect URI needed)
2. API permissions -> Add -> Microsoft Graph -> Application permissions -> Mail.Send
3. Grant admin consent for your organization
4. Certificates & secrets -> New client secret -> copy the value (shown once)
5. Note the Tenant ID and Client ID from the app Overview page

Run the script with `-Setup` to enter these values interactively. They are stored encrypted in `config\<username>.xml` and reused on subsequent runs.

## Quick Start

```powershell
# First-time setup (run interactively)
.\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it@contoso.com' -Setup

# Normal daily run
.\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it@contoso.com'

# Dry-run with diagnostics
.\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it@contoso.com' -ReportOnly -Diagnostics
```

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-GroupName` | string | mandatory | AD security group containing contractor accounts |
| `-To` | string[] | mandatory | Primary recipient email addresses |
| `-Cc` | string[] | -- | CC recipient email addresses |
| `-WarnDays` | int | 14 | Warn for accounts expiring within this many days |
| `-Setup` | switch | -- | Run the first-time setup wizard (or re-run to rotate the client secret) |
| `-ReportOnly` | switch | -- | Log intended email without sending |
| `-Diagnostics` | switch | -- | Print member counts and filter statistics |
| `-LogHistory` | int | 30 | Days of transcript logs to retain |
| `-SkipUpdateCheck` | switch | -- | Skip GitHub version check at startup |

## Email Output

The digest contains up to two sections:

1. **Expiring Within N Days** -- table of enabled accounts whose `AccountExpirationDate` falls within the warning window, sorted ascending by days remaining. Rows with 3 or fewer days remaining are highlighted red.
2. **No Expiration Date Configured** -- enabled accounts in the group with no `AccountExpirationDate` set (a potential configuration gap).

If neither section has entries, no email is sent.

## Logs

Transcripts are written to `logs\ContractorExpirationAlert_<date>.txt` in the script directory. Files older than `-LogHistory` days are deleted automatically on each run.

## Scheduled Task

```
Program:   pwsh.exe   (PowerShell 7+)  or  powershell.exe  (Windows PowerShell 5.1)
Arguments: -NonInteractive -File "C:\scripts\contractor-expiration-alert\Send-ContractorExpirationAlert.ps1"
           -GroupName "Contractors"
           -To "it-team@contoso.com"
           -Cc "manager@contoso.com"
           -WarnDays 14
Trigger:   Daily, 07:00
Run as:    A service account with AD read access
```

The service account running the scheduled task must have access to the `config\<username>.xml` file written during setup. Run `-Setup` as the same account the task will use.

## Version

2.0.0 - Replaced SMTP with Microsoft Graph for email delivery.
