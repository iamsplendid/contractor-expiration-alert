# contractor-expiration-alert

PowerShell script that sends a daily HTML digest email listing contractor AD accounts expiring within a configurable window. Intended to run as a scheduled task.

## Requirements

- PowerShell 5.1 or later
- ActiveDirectory module (RSAT: Active Directory Domain Services)
- Domain-joined machine with network access to a domain controller
- SMTP relay accessible from the host machine

## Quick Start

```powershell
.\Send-ContractorExpirationAlert.ps1 `
    -GroupName 'Contractors' `
    -To 'it-team@contoso.com' `
    -Cc 'manager@contoso.com' `
    -SmtpServer 'smtp.contoso.com' `
    -FromAddress 'noreply@contoso.com' `
    -WarnDays 14
```

Dry-run (no email sent):

```powershell
.\Send-ContractorExpirationAlert.ps1 `
    -GroupName 'Contractors' `
    -To 'it-team@contoso.com' `
    -SmtpServer 'smtp.contoso.com' `
    -FromAddress 'noreply@contoso.com' `
    -ReportOnly -Diagnostics
```

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-GroupName` | string | mandatory | AD security group containing contractor accounts |
| `-To` | string[] | mandatory | Primary recipient email addresses |
| `-Cc` | string[] | — | CC recipient email addresses |
| `-SmtpServer` | string | mandatory | SMTP relay hostname |
| `-FromAddress` | string | mandatory | Sender email address |
| `-WarnDays` | int | 14 | Warn for accounts expiring within this many days |
| `-SmtpPort` | int | 25 | SMTP port |
| `-ReportOnly` | switch | — | Log intended email without sending |
| `-Diagnostics` | switch | — | Print member counts and filter statistics |
| `-LogHistory` | int | 30 | Days of transcript logs to retain |
| `-SkipUpdateCheck` | switch | — | Skip GitHub version check at startup |

## Email Output

The digest contains up to two sections:

1. **Expiring Within N Days** — table of enabled accounts whose `AccountExpirationDate` falls within the warning window, sorted ascending by days remaining. Rows with ≤ 3 days remaining are highlighted red.
2. **No Expiration Date Configured** — enabled accounts in the group with no `AccountExpirationDate` set (a potential configuration gap).

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
           -SmtpServer "smtp.contoso.com"
           -FromAddress "noreply@contoso.com"
           -WarnDays 14
Trigger:   Daily, 07:00
Run as:    A service account with AD read access
```
