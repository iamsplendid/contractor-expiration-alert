# Contractor Expiration Alert — Design Spec

**Date:** 2026-05-05
**Status:** Approved

---

## Overview

A PowerShell script (`Send-ContractorExpirationAlert.ps1`) that queries an on-premises Active Directory security group for contractor accounts whose `AccountExpirationDate` is approaching, then sends a single HTML digest email to a configured recipient list. Intended to run as a scheduled task (daily).

---

## Parameters

| Parameter | Type | Mandatory | Default | Description |
|---|---|---|---|---|
| `-GroupName` | `string` | yes | — | AD security group to scan |
| `-To` | `string[]` | yes | — | Primary email recipients |
| `-Cc` | `string[]` | no | — | CC recipients |
| `-SmtpServer` | `string` | yes | — | SMTP relay hostname |
| `-FromAddress` | `string` | yes | — | Sender email address |
| `-WarnDays` | `int` | no | `14` | Alert window in days |
| `-SmtpPort` | `int` | no | `25` | SMTP port |
| `-ReportOnly` | `switch` | no | — | Dry-run: log intended send, skip actual send |
| `-Diagnostics` | `switch` | no | — | Print member counts and filter stats to console |
| `-LogHistory` | `int` | no | `30` | Days of transcript logs to retain |
| `-SkipUpdateCheck` | `switch` | no | — | Skip GitHub version check at startup |

---

## Core Logic / Flow

1. **Auto-update check** — fetches the script's own raw URL from GitHub, compares `$ScriptVersion`, and re-runs the updated version if a newer one is available. Same pattern used in the `diagnose-mailbox` scripts. Skipped if `-SkipUpdateCheck` is passed.

2. **Transcript logging** — starts a transcript at `<script-dir>\logs\ContractorExpirationAlert_<yyyy-MM-dd_HHmmss>.txt`. On startup, removes transcript files older than `-LogHistory` days.

3. **AD module check** — verifies the `ActiveDirectory` module is available; exits with `Write-Error` if not.

4. **Group query** — calls `Get-ADGroupMember -Identity $GroupName -Recursive`. If the group is not found, exits with `Write-Error` (no partial results sent).

5. **User property fetch** — for each group member, calls `Get-ADUser` to retrieve: `DisplayName`, `SamAccountName`, `EmailAddress`, `Enabled`, `AccountExpirationDate`. Individual lookup failures are logged as warnings and skipped (the run continues).

6. **Filtering**
   - Keep only `Enabled = $true` accounts.
   - Split the remaining accounts into two buckets:
     - **Expiring soon:** `AccountExpirationDate` is set and falls within the next `$WarnDays` days.
     - **No expiration set:** `AccountExpirationDate` is null/not set (these are flagged as a configuration risk).
   - Disabled accounts and accounts expiring beyond `$WarnDays` are silently excluded.

7. **Diagnostics output** — if `-Diagnostics` is passed, prints to console: total group members, count disabled (skipped), count with no expiration (flagged), count expiring within window, count beyond window (skipped).

8. **Nothing to report** — if both buckets are empty, logs "Nothing to report" and exits without sending email.

9. **Email send** — builds the HTML digest (see Email Format below) and calls `Send-MailMessage`. If `-ReportOnly` is set, logs what would have been sent and skips the call.

---

## Email Format

Single HTML email, inline CSS only (no external files).

**Subject:** `Contractor Account Expiration Alert — <N> account(s) expiring within <WarnDays> days`

**Body structure:**

1. **Header** — bold title "Contractor Account Expiration Alert", run timestamp.

2. **Table 1 — Expiring Soon** (present only if bucket is non-empty)
   - Columns: Display Name | Username | Email Address | Expiration Date | Days Remaining
   - Sorted ascending by Days Remaining (most urgent first)
   - Rows where Days Remaining ≤ 3 get a red background highlight; all other rows use default styling.

3. **Table 2 — No Expiration Date Configured** (present only if bucket is non-empty)
   - Columns: Display Name | Username | Email Address
   - Flat warning table; no color coding.
   - Introductory note: "The following contractor accounts are members of the group but have no account expiration date set. This may be a configuration oversight."

4. **Footer** — script version, SMTP server used.

---

## Error Handling

| Scenario | Behavior |
|---|---|
| `ActiveDirectory` module not available | `Write-Error` + exit |
| Group not found in AD | `Write-Error` + exit (no email sent) |
| Individual user lookup failure | `Write-Warning` + skip account (run continues) |
| SMTP send failure | `Write-Warning` with exception message (captured in transcript) |
| Both buckets empty | Log "Nothing to report" + exit (no email sent) |

---

## Repo Structure

```
contractor-expiration-alert/
├── Send-ContractorExpirationAlert.ps1
├── logs/                        # auto-created on first run
├── docs/
│   └── superpowers/
│       └── specs/
│           └── 2026-05-05-contractor-expiration-alert-design.md
└── README.md
```

---

## Scheduled Task Usage

Run daily. Example invocation:

```powershell
.\Send-ContractorExpirationAlert.ps1 `
    -GroupName 'Contractors' `
    -To 'it-team@contoso.com' `
    -Cc 'manager@contoso.com' `
    -SmtpServer 'smtp.contoso.com' `
    -FromAddress 'noreply@contoso.com' `
    -WarnDays 14
```

Dry-run to verify output before scheduling:

```powershell
.\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it-team@contoso.com' -SmtpServer 'smtp.contoso.com' -FromAddress 'noreply@contoso.com' -ReportOnly -Diagnostics
```

---

## Patterns Inherited from Reference Scripts

- Auto-update via GitHub raw URL (`$ScriptVersion` + `Invoke-WebRequest`)
- Transcript logging with retention cleanup
- `-ReportOnly` and `-Diagnostics` switches
- `Send-MailMessage` with `-SmtpServer` / `-Port`
- `[CmdletBinding()]` with typed `param()` block
- Color-coded `Write-Host` console output (Cyan = info, Yellow = warn, Red = fail, Green = ok)
