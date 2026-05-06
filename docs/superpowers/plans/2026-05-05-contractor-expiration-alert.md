# Contractor Expiration Alert — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `Send-ContractorExpirationAlert.ps1` — a scheduled-task-ready PowerShell script that queries an on-prem AD security group and emails a single HTML digest of contractor accounts whose `AccountExpirationDate` is approaching.

**Architecture:** Single self-contained `.ps1` file. Follows patterns from two sibling repos — auto-update from GitHub raw URL (`diagnose-mailbox` pattern), transcript logging + `-ReportOnly`/`-Diagnostics` switches + `Send-MailMessage` over SMTP relay (`passwordreminder` pattern). AD queries use `Get-ADGroupMember` + `Get-ADUser` from the `ActiveDirectory` module. HTML email is generated inline (no external template files).

**Tech Stack:** PowerShell 5.1+, ActiveDirectory RSAT module, `Send-MailMessage`, `gh` CLI (for GitHub repo creation).

---

### Task 1: Repo scaffolding

**Files:**
- Create: `/mnt/c/github/contractor-expiration-alert/.gitignore`
- Create: `/mnt/c/github/contractor-expiration-alert/logs/.gitkeep`

- [ ] **Step 1: Create .gitignore**

Create `/mnt/c/github/contractor-expiration-alert/.gitignore` with this content:

```
logs/*.txt
```

- [ ] **Step 2: Create logs placeholder so git tracks the folder**

```bash
touch /mnt/c/github/contractor-expiration-alert/logs/.gitkeep
```

- [ ] **Step 3: Commit**

```bash
cd /mnt/c/github/contractor-expiration-alert
git add .gitignore logs/.gitkeep
git commit -m "Add repo scaffolding: .gitignore and logs placeholder"
```

---

### Task 2: Create GitHub repo and connect remote

**Files:** None (GitHub + git remote config only).

- [ ] **Step 1: Create the GitHub repo**

```bash
gh repo create iamsplendid/contractor-expiration-alert --public --description "PowerShell script to alert on expiring contractor AD accounts"
```

Expected output: `✓ Created repository iamsplendid/contractor-expiration-alert on GitHub`

- [ ] **Step 2: Add remote and push**

```bash
cd /mnt/c/github/contractor-expiration-alert
git remote add origin https://github.com/iamsplendid/contractor-expiration-alert.git
git push -u origin master
```

Expected: the spec doc commit and scaffolding commit are pushed.

---

### Task 3: Write Send-ContractorExpirationAlert.ps1

**Files:**
- Create: `/mnt/c/github/contractor-expiration-alert/Send-ContractorExpirationAlert.ps1`

- [ ] **Step 1: Create the script**

Create `/mnt/c/github/contractor-expiration-alert/Send-ContractorExpirationAlert.ps1` with the following complete content:

```powershell
<#
.SYNOPSIS
    Send advance expiration alerts for contractor domain accounts.
.DESCRIPTION
    Queries all members of an Active Directory security group for accounts whose
    AccountExpirationDate falls within the configured warning window, then sends
    a single HTML digest email to the configured recipients.
    Accounts with no expiration date set are flagged in a separate section.
    Intended to run as a daily scheduled task.
.PARAMETER GroupName
    Name of the AD security group whose members are contractor accounts.
.PARAMETER To
    One or more primary recipient email addresses.
.PARAMETER Cc
    One or more CC recipient email addresses (optional).
.PARAMETER SmtpServer
    SMTP relay hostname used to send the notification email.
.PARAMETER FromAddress
    Sender email address for the notification email.
.PARAMETER WarnDays
    Number of days ahead to warn about expiring accounts. Default: 14.
.PARAMETER SmtpPort
    SMTP port. Default: 25.
.PARAMETER ReportOnly
    Dry-run switch. Logs what would be sent without actually sending any email.
.PARAMETER Diagnostics
    Prints member counts and filter statistics to the console.
.PARAMETER LogHistory
    Number of days to retain transcript log files. Default: 30.
.PARAMETER SkipUpdateCheck
    Skip the automatic version update check at startup.
.EXAMPLE
    .\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it@contoso.com' -SmtpServer 'smtp.contoso.com' -FromAddress 'noreply@contoso.com'
.EXAMPLE
    .\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it@contoso.com' -SmtpServer 'smtp.contoso.com' -FromAddress 'noreply@contoso.com' -ReportOnly -Diagnostics
.NOTES
    Requires: ActiveDirectory PowerShell module (RSAT).
    This script is read-only with respect to Active Directory.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]  [string]   $GroupName,
    [Parameter(Mandatory = $true)]  [string[]] $To,
                                    [string[]] $Cc,
    [Parameter(Mandatory = $true)]  [string]   $SmtpServer,
    [Parameter(Mandatory = $true)]  [string]   $FromAddress,
                                    [int]      $WarnDays   = 14,
                                    [int]      $SmtpPort   = 25,
                                    [switch]   $ReportOnly,
                                    [switch]   $Diagnostics,
                                    [int]      $LogHistory  = 30,
                                    [switch]   $SkipUpdateCheck
)

$ScriptVersion   = '1.0.0'
$ScriptUpdateUrl = 'https://raw.githubusercontent.com/iamsplendid/contractor-expiration-alert/master/Send-ContractorExpirationAlert.ps1'

# ── Auto-update ──────────────────────────────────────────────────────────────
if (-not $SkipUpdateCheck) {
    try {
        $remote = (Invoke-WebRequest -Uri $ScriptUpdateUrl -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop).Content
        if ($remote -match '\$ScriptVersion\s*=\s*[''"]([^''"]+)[''"]') {
            $remoteVersion = $Matches[1]
            if ([version]$remoteVersion -gt [version]$ScriptVersion) {
                Write-Host "[UPDATE] New version $remoteVersion available. Updating..." -ForegroundColor Cyan
                $scriptPath = $PSCommandPath
                if ($scriptPath -and (Test-Path $scriptPath)) {
                    [System.IO.File]::WriteAllText($scriptPath, $remote, [System.Text.Encoding]::UTF8)
                    Write-Host "[UPDATE] Re-running new version..." -ForegroundColor Green
                    $fwd = @{} + $PSBoundParameters
                    $fwd['SkipUpdateCheck'] = $true
                    & $scriptPath @fwd
                    exit
                } else {
                    Write-Warning "[UPDATE] Cannot determine script path. Download latest: $ScriptUpdateUrl"
                }
            } else {
                Write-Verbose "[UPDATE] Script is current ($ScriptVersion)."
            }
        }
    } catch {
        Write-Verbose "[UPDATE] Version check skipped: $($_.Exception.Message)"
    }
}

# ── Transcript logging ───────────────────────────────────────────────────────
$logsDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir | Out-Null }

$transcriptPath    = Join-Path $logsDir "ContractorExpirationAlert_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"
$transcriptStarted = $false
try {
    Start-Transcript -Path $transcriptPath -UseMinimalHeader | Out-Null
    $transcriptStarted = $true
} catch {
    try {
        Start-Transcript -Path $transcriptPath | Out-Null
        $transcriptStarted = $true
    } catch {
        Write-Warning "Could not start transcript: $($_.Exception.Message)"
    }
}

Get-ChildItem -Path $logsDir -Filter 'ContractorExpirationAlert_*.txt' -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogHistory) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

$start = Get-Date

Write-Host ''
Write-Host ('=' * 70) -ForegroundColor Cyan
Write-Host "  Contractor Expiration Alert  |  Group: $GroupName" -ForegroundColor Cyan
Write-Host "  WarnDays : $WarnDays" -ForegroundColor Cyan
Write-Host "  Time     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "  Version  : $ScriptVersion" -ForegroundColor Cyan
Write-Host ('=' * 70) -ForegroundColor Cyan

# ── AD module check ──────────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error 'ActiveDirectory module is not available. Install RSAT or run on a domain-joined machine with the module present.'
    if ($transcriptStarted) { Stop-Transcript | Out-Null }
    exit 1
}
Import-Module ActiveDirectory -ErrorAction Stop

# ── Group query ──────────────────────────────────────────────────────────────
Write-Host "[INFO] Querying group '$GroupName'..." -ForegroundColor Cyan
try {
    $groupMembers = @(Get-ADGroupMember -Identity $GroupName -Recursive -ErrorAction Stop)
} catch {
    Write-Error "Could not query group '$GroupName': $($_.Exception.Message)"
    if ($transcriptStarted) { Stop-Transcript | Out-Null }
    exit 1
}
Write-Host "[INFO] Found $($groupMembers.Count) member(s) in group '$GroupName'." -ForegroundColor Cyan

# ── User property fetch ──────────────────────────────────────────────────────
$allUsers = [System.Collections.Generic.List[pscustomobject]]::new()
foreach ($member in $groupMembers) {
    if ($member.objectClass -ne 'user') {
        Write-Verbose "Skipping non-user object: $($member.SamAccountName) ($($member.objectClass))"
        continue
    }
    try {
        $u = Get-ADUser -Identity $member.DistinguishedName `
            -Properties DisplayName, SamAccountName, EmailAddress, Enabled, AccountExpirationDate `
            -ErrorAction Stop
        $allUsers.Add([pscustomobject]@{
            DisplayName           = $u.DisplayName
            SamAccountName        = $u.SamAccountName
            EmailAddress          = $u.EmailAddress
            Enabled               = $u.Enabled
            AccountExpirationDate = $u.AccountExpirationDate
        })
    } catch {
        Write-Warning "Failed to retrieve properties for '$($member.SamAccountName)': $($_.Exception.Message)"
    }
}

# ── Filtering ────────────────────────────────────────────────────────────────
$now    = Get-Date
$cutoff = $now.AddDays($WarnDays)

$enabledUsers  = @($allUsers | Where-Object { $_.Enabled })
$disabledUsers = @($allUsers | Where-Object { -not $_.Enabled })

$noExpirationUsers = @($enabledUsers | Where-Object { -not $_.AccountExpirationDate })
$withExpiration    = @($enabledUsers | Where-Object { $_.AccountExpirationDate })

$expiringUsers = @(
    $withExpiration |
        Where-Object { $_.AccountExpirationDate -gt $now -and $_.AccountExpirationDate -le $cutoff } |
        ForEach-Object {
            $days = [math]::Ceiling(($_.AccountExpirationDate - $now).TotalDays)
            [pscustomobject]@{
                DisplayName           = $_.DisplayName
                SamAccountName        = $_.SamAccountName
                EmailAddress          = $_.EmailAddress
                AccountExpirationDate = $_.AccountExpirationDate
                DaysRemaining         = $days
            }
        } |
        Sort-Object DaysRemaining
)

$beyondWindowCount = @($withExpiration | Where-Object { $_.AccountExpirationDate -gt $cutoff }).Count

# ── Diagnostics ──────────────────────────────────────────────────────────────
if ($Diagnostics) {
    Write-Host ''
    Write-Host 'Diagnostics:' -ForegroundColor Cyan
    Write-Host "  Total group members (user objects)      : $($allUsers.Count)"
    Write-Host "  Disabled (skipped)                      : $($disabledUsers.Count)"
    Write-Host "  No expiration date set (flagged)        : $($noExpirationUsers.Count)"
    Write-Host "  Expiring within $WarnDays day(s)               : $($expiringUsers.Count)"
    Write-Host "  Expiring beyond $WarnDays day(s) (skipped)     : $beyondWindowCount"
}

# ── Nothing to report ────────────────────────────────────────────────────────
if ($expiringUsers.Count -eq 0 -and $noExpirationUsers.Count -eq 0) {
    Write-Host "[INFO] Nothing to report. No contractor accounts expiring within $WarnDays day(s) and no accounts missing an expiration date." -ForegroundColor Green
    $elapsed = [int](New-TimeSpan -Start $start).TotalSeconds
    Write-Host "[INFO] Completed in $elapsed second(s)." -ForegroundColor Cyan
    if ($transcriptStarted) { Stop-Transcript | Out-Null }
    exit 0
}

# ── HTML builder ─────────────────────────────────────────────────────────────
function Build-HtmlEmail {
    param(
        [object[]] $ExpiringUsers,
        [object[]] $NoExpirationUsers,
        [int]      $WarnDays,
        [string]   $GroupName,
        [string]   $ScriptVersion,
        [string]   $SmtpServer
    )

    $runTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    # Table 1 — Expiring Soon
    $table1Html = ''
    if ($ExpiringUsers.Count -gt 0) {
        $rows = ''
        foreach ($u in $ExpiringUsers) {
            $expDate  = $u.AccountExpirationDate.ToString('yyyy-MM-dd')
            $rowStyle = if ($u.DaysRemaining -le 3) { ' style="background-color:#ffd6d6;"' } else { '' }
            $email    = if ($u.EmailAddress) { [System.Net.WebUtility]::HtmlEncode($u.EmailAddress) } else { '&mdash;' }
            $rows += "
            <tr$rowStyle>
                <td style='padding:6px 10px; border:1px solid #ddd;'>$([System.Net.WebUtility]::HtmlEncode($u.DisplayName))</td>
                <td style='padding:6px 10px; border:1px solid #ddd;'>$([System.Net.WebUtility]::HtmlEncode($u.SamAccountName))</td>
                <td style='padding:6px 10px; border:1px solid #ddd;'>$email</td>
                <td style='padding:6px 10px; border:1px solid #ddd;'>$expDate</td>
                <td style='padding:6px 10px; border:1px solid #ddd; text-align:center;'>$($u.DaysRemaining)</td>
            </tr>"
        }
        $table1Html = "
    <h2 style='font-size:16px; color:#2f3b52; margin:20px 0 8px 0;'>Expiring Within $WarnDays Days ($($ExpiringUsers.Count) account(s))</h2>
    <p style='font-size:13px; color:#555; margin:0 0 8px 0;'>Rows highlighted in red are expiring within 3 days.</p>
    <table style='border-collapse:collapse; width:100%; font-size:14px;'>
        <thead>
            <tr style='background-color:#2f3b52; color:#fff;'>
                <th style='padding:8px 10px; text-align:left;'>Display Name</th>
                <th style='padding:8px 10px; text-align:left;'>Username</th>
                <th style='padding:8px 10px; text-align:left;'>Email Address</th>
                <th style='padding:8px 10px; text-align:left;'>Expiration Date</th>
                <th style='padding:8px 10px; text-align:center;'>Days Remaining</th>
            </tr>
        </thead>
        <tbody>$rows
        </tbody>
    </table>"
    }

    # Table 2 — No Expiration Date Configured
    $table2Html = ''
    if ($NoExpirationUsers.Count -gt 0) {
        $rows = ''
        foreach ($u in $NoExpirationUsers) {
            $email = if ($u.EmailAddress) { [System.Net.WebUtility]::HtmlEncode($u.EmailAddress) } else { '&mdash;' }
            $rows += "
            <tr>
                <td style='padding:6px 10px; border:1px solid #ddd;'>$([System.Net.WebUtility]::HtmlEncode($u.DisplayName))</td>
                <td style='padding:6px 10px; border:1px solid #ddd;'>$([System.Net.WebUtility]::HtmlEncode($u.SamAccountName))</td>
                <td style='padding:6px 10px; border:1px solid #ddd;'>$email</td>
            </tr>"
        }
        $table2Html = "
    <h2 style='font-size:16px; color:#2f3b52; margin:20px 0 8px 0;'>No Expiration Date Configured ($($NoExpirationUsers.Count) account(s))</h2>
    <p style='font-size:13px; color:#555; margin:0 0 8px 0;'>The following contractor accounts are members of the group but have no account expiration date set. This may be a configuration oversight.</p>
    <table style='border-collapse:collapse; width:100%; font-size:14px;'>
        <thead>
            <tr style='background-color:#2f3b52; color:#fff;'>
                <th style='padding:8px 10px; text-align:left;'>Display Name</th>
                <th style='padding:8px 10px; text-align:left;'>Username</th>
                <th style='padding:8px 10px; text-align:left;'>Email Address</th>
            </tr>
        </thead>
        <tbody>$rows
        </tbody>
    </table>"
    }

    return @"
<!DOCTYPE html>
<html>
<body style="font-family:Segoe UI,Arial,sans-serif; color:#333; margin:0; padding:20px; background:#f5f5f5;">
<div style="max-width:800px; margin:0 auto; background:#fff; padding:24px; border-radius:4px; border:1px solid #ddd;">

    <h1 style="font-size:20px; color:#2f3b52; margin:0 0 4px 0;">Contractor Account Expiration Alert</h1>
    <p style="font-size:13px; color:#777; margin:0 0 20px 0;">
        Group: <strong>$([System.Net.WebUtility]::HtmlEncode($GroupName))</strong> &nbsp;|&nbsp; Run time: $runTime
    </p>
$table1Html
$table2Html
    <hr style="border:none; border-top:1px solid #eee; margin:24px 0 12px 0;">
    <p style="font-size:12px; color:#aaa; margin:0;">
        Send-ContractorExpirationAlert v$ScriptVersion &nbsp;|&nbsp; SMTP: $([System.Net.WebUtility]::HtmlEncode($SmtpServer))
    </p>

</div>
</body>
</html>
"@
}

# ── Build and send email ─────────────────────────────────────────────────────
$htmlBody = Build-HtmlEmail `
    -ExpiringUsers     $expiringUsers `
    -NoExpirationUsers $noExpirationUsers `
    -WarnDays          $WarnDays `
    -GroupName         $GroupName `
    -ScriptVersion     $ScriptVersion `
    -SmtpServer        $SmtpServer

$subject = "Contractor Account Expiration Alert — $($expiringUsers.Count) account(s) expiring within $WarnDays days"

$mailParams = @{
    From       = $FromAddress
    To         = $To
    Subject    = $subject
    Body       = $htmlBody
    BodyAsHtml = $true
    SmtpServer = $SmtpServer
    Port       = $SmtpPort
    Encoding   = 'UTF8'
}
if ($Cc) { $mailParams['Cc'] = $Cc }

if ($ReportOnly) {
    Write-Host ''
    Write-Host '[REPORT ONLY] Would send email:' -ForegroundColor Yellow
    Write-Host "  Subject            : $subject" -ForegroundColor Yellow
    Write-Host "  To                 : $($To -join ', ')" -ForegroundColor Yellow
    if ($Cc) { Write-Host "  Cc                 : $($Cc -join ', ')" -ForegroundColor Yellow }
    Write-Host "  Expiring accounts  : $($expiringUsers.Count)" -ForegroundColor Yellow
    Write-Host "  No-expiry accounts : $($noExpirationUsers.Count)" -ForegroundColor Yellow
} else {
    Write-Host "[INFO] Sending email to: $($To -join ', ')..." -ForegroundColor Cyan
    try {
        Send-MailMessage @mailParams -WarningAction SilentlyContinue
        Write-Host '[INFO] Email sent successfully.' -ForegroundColor Green
    } catch {
        Write-Warning "Failed to send email: $($_.Exception.Message)"
    }
}

$elapsed = [int](New-TimeSpan -Start $start).TotalSeconds
Write-Host "[INFO] Completed in $elapsed second(s)." -ForegroundColor Cyan

if ($transcriptStarted) { Stop-Transcript | Out-Null }
```

- [ ] **Step 2: Verify the script has no syntax errors**

Run this from a PowerShell prompt (Windows) — it parses without executing:

```powershell
$tokens = $null; $errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile(
    'C:\github\contractor-expiration-alert\Send-ContractorExpirationAlert.ps1',
    [ref]$tokens, [ref]$errors
)
$errors
```

Expected: no output (empty `$errors` array). If any parse errors are listed, fix the indicated line number and re-run.

- [ ] **Step 3: Commit**

```bash
cd /mnt/c/github/contractor-expiration-alert
git add Send-ContractorExpirationAlert.ps1
git commit -m "Add Send-ContractorExpirationAlert.ps1"
```

---

### Task 4: Smoke test with -ReportOnly and -Diagnostics

**Files:** None modified.

Run these tests from a domain-joined Windows machine with the `ActiveDirectory` module installed and read access to the target group. Adjust `GroupName`, `To`, and `SmtpServer` to match the environment.

- [ ] **Step 1: Dry-run against a real group**

```powershell
.\Send-ContractorExpirationAlert.ps1 `
    -GroupName '<your-contractor-group>' `
    -To 'test@contoso.com' `
    -SmtpServer 'smtp.contoso.com' `
    -FromAddress 'noreply@contoso.com' `
    -WarnDays 14 `
    -ReportOnly `
    -Diagnostics `
    -SkipUpdateCheck `
    -Verbose
```

Expected output:
- Banner header with group name and version
- `Diagnostics:` section showing member counts for each bucket
- `[REPORT ONLY] Would send email:` block showing subject and account counts
- No email sent, no errors

- [ ] **Step 2: Verify the group-not-found error path**

```powershell
.\Send-ContractorExpirationAlert.ps1 `
    -GroupName 'GroupThatDefinitelyDoesNotExist' `
    -To 'test@contoso.com' `
    -SmtpServer 'smtp.contoso.com' `
    -FromAddress 'noreply@contoso.com' `
    -SkipUpdateCheck
```

Expected: `Write-Error` message referencing the group name, script exits — no email sent.

- [ ] **Step 3: Verify transcript was written**

```powershell
Get-ChildItem -Path '.\logs\' -Filter 'ContractorExpirationAlert_*.txt'
```

Expected: one `.txt` file with today's date in the filename.

---

### Task 5: Write README.md

**Files:**
- Create: `/mnt/c/github/contractor-expiration-alert/README.md`

- [ ] **Step 1: Create README.md**

Create `/mnt/c/github/contractor-expiration-alert/README.md` with the following content:

````markdown
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
2. **No Expiration Date Configured** — enabled accounts in the group with no `AccountExpirationDate` set (a configuration oversight worth reviewing).

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
````

- [ ] **Step 2: Commit**

```bash
cd /mnt/c/github/contractor-expiration-alert
git add README.md
git commit -m "Add README"
```

---

### Task 6: Push everything to GitHub

**Files:** None.

- [ ] **Step 1: Push all commits**

```bash
cd /mnt/c/github/contractor-expiration-alert
git push
```

Expected: all commits (spec, scaffolding, script, README) visible on GitHub.

- [ ] **Step 2: Confirm on GitHub**

```bash
gh repo view iamsplendid/contractor-expiration-alert --web
```

Verify `Send-ContractorExpirationAlert.ps1` and `README.md` are present in the repository root.
