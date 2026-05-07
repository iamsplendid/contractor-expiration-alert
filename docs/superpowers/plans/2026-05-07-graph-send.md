# Graph Send Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace `Send-MailMessage` (SMTP AUTH) with Microsoft Graph `POST /users/{from}/sendMail`, add a per-user DPAPI-encrypted config file, and a first-run setup wizard.

**Architecture:** All changes are confined to `Send-ContractorExpirationAlert.ps1`. Helper functions (validation, config I/O, Graph calls, setup wizard) are added as named functions before `Build-HtmlEmail`. The main flow gains a config-load/setup step between the startup banner and the AD query. The SMTP param block and send block are replaced entirely.

**Tech Stack:** PowerShell 5.1, `Invoke-RestMethod`, `Export-Clixml`/`Import-Clixml` (DPAPI), `Read-Host -AsSecureString`, Pester 5 (unit tests), Microsoft Graph API v1.0, Entra ID OAuth 2.0 client credentials grant.

---

## File Structure

| File | Action | Responsibility |
|---|---|---|
| `Send-ContractorExpirationAlert.ps1` | Modify | Main script — all changes land here |
| `tests/Send-ContractorExpirationAlert.Tests.ps1` | Create | Pester unit tests for pure-logic helpers |

---

## Task 1: Set up Pester test file with failing tests for validation helpers

**Files:**
- Create: `tests/Send-ContractorExpirationAlert.Tests.ps1`

- [ ] **Step 1: Verify Pester is available**

```powershell
Get-Module -ListAvailable Pester | Select-Object Name, Version
```

Expected: Pester 5.x listed. If missing:
```powershell
Install-Module Pester -Force -Scope CurrentUser
```

- [ ] **Step 2: Create the test file**

Create `tests/Send-ContractorExpirationAlert.Tests.ps1` with this content:

```powershell
BeforeAll {
    function Test-IsGuid {
        param([string]$Value)
        # placeholder — will be replaced when Task 2 is done
        throw 'not implemented'
    }

    function Test-IsEmail {
        param([string]$Value)
        # placeholder — will be replaced when Task 2 is done
        throw 'not implemented'
    }
}

Describe 'Test-IsGuid' {
    It 'returns true for a valid lowercase GUID' {
        Test-IsGuid 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' | Should -Be $true
    }
    It 'returns true for a valid uppercase GUID' {
        Test-IsGuid 'AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE' | Should -Be $true
    }
    It 'returns false for an empty string' {
        Test-IsGuid '' | Should -Be $false
    }
    It 'returns false for a plain word' {
        Test-IsGuid 'not-a-guid' | Should -Be $false
    }
    It 'returns false for a partial GUID' {
        Test-IsGuid 'aaaaaaaa-bbbb-cccc' | Should -Be $false
    }
}

Describe 'Test-IsEmail' {
    It 'returns true for a valid email address' {
        Test-IsEmail 'user@domain.com' | Should -Be $true
    }
    It 'returns true for a UPN-style address' {
        Test-IsEmail 'alerts@contoso.onmicrosoft.com' | Should -Be $true
    }
    It 'returns false for a string without @' {
        Test-IsEmail 'nodomain' | Should -Be $false
    }
    It 'returns false for an empty string' {
        Test-IsEmail '' | Should -Be $false
    }
}
```

- [ ] **Step 3: Run tests — verify they fail with 'not implemented'**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: All 9 tests fail with `RuntimeException: not implemented`.

---

## Task 2: Implement Test-IsGuid and Test-IsEmail — make tests pass

**Files:**
- Modify: `Send-ContractorExpirationAlert.ps1` (add helper functions block after line 65, before the auto-update block)
- Modify: `tests/Send-ContractorExpirationAlert.Tests.ps1` (replace placeholder implementations)

- [ ] **Step 1: Add helper functions block to the main script**

Insert the following block after line 65 (`$ScriptUpdateUrl = ...`) and before line 67 (`# ── Auto-update`):

```powershell
# ── Helper functions ─────────────────────────────────────────────────────────
function Test-IsGuid {
    param([string]$Value)
    return $Value -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
}

function Test-IsEmail {
    param([string]$Value)
    return ($Value -match '@') -and ($Value.Length -gt 3)
}

```

- [ ] **Step 2: Update the test file — replace placeholder implementations with real ones**

Replace the entire `BeforeAll` block in `tests/Send-ContractorExpirationAlert.Tests.ps1` with:

```powershell
BeforeAll {
    function Test-IsGuid {
        param([string]$Value)
        return $Value -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    }

    function Test-IsEmail {
        param([string]$Value)
        return ($Value -match '@') -and ($Value.Length -gt 3)
    }
}
```

- [ ] **Step 3: Run tests — verify all 9 pass**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 9, Failed: 0`

- [ ] **Step 4: Commit**

```powershell
git add Send-ContractorExpirationAlert.ps1 tests/Send-ContractorExpirationAlert.Tests.ps1
git commit -m "Add Test-IsGuid and Test-IsEmail helpers with Pester tests"
```

---

## Task 3: Write failing tests for config helpers

**Files:**
- Modify: `tests/Send-ContractorExpirationAlert.Tests.ps1`

- [ ] **Step 1: Append config helper tests to the test file**

Add the following after the existing `Describe 'Test-IsEmail'` block:

```powershell
Describe 'Save-AlertConfig / Read-AlertConfig' {
    BeforeAll {
        function Save-AlertConfig {
            param([hashtable]$Config, [string]$Path)
            throw 'not implemented'
        }
        function Read-AlertConfig {
            param([string]$Path)
            throw 'not implemented'
        }
    }

    It 'round-trips all four fields including the encrypted secret' {
        $tempPath = Join-Path $TestDrive "config\$env:USERNAME.xml"
        $null = New-Item (Split-Path $tempPath) -ItemType Directory -Force

        $config = @{
            TenantId     = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
            ClientId     = 'ffffffff-0000-1111-2222-333333333333'
            ClientSecret = ConvertTo-SecureString 'test-secret' -AsPlainText -Force
            FromAddress  = 'alerts@test.com'
        }
        Save-AlertConfig -Config $config -Path $tempPath
        $loaded = Read-AlertConfig -Path $tempPath

        $loaded.TenantId    | Should -Be 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $loaded.ClientId    | Should -Be 'ffffffff-0000-1111-2222-333333333333'
        $loaded.FromAddress | Should -Be 'alerts@test.com'

        $ptr   = [Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($loaded.ClientSecret)
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
        [Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr)
        $plain | Should -Be 'test-secret'
    }

    It 'returns null when the config file does not exist' {
        Read-AlertConfig -Path (Join-Path $TestDrive 'nonexistent.xml') | Should -BeNullOrEmpty
    }
}
```

- [ ] **Step 2: Run tests — verify new tests fail with 'not implemented'**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 9, Failed: 2` — the two new config tests fail.

---

## Task 4: Implement config helpers — make tests pass

**Files:**
- Modify: `Send-ContractorExpirationAlert.ps1` (add three functions to the helper functions block)
- Modify: `tests/Send-ContractorExpirationAlert.Tests.ps1` (replace placeholder implementations)

- [ ] **Step 1: Add config helpers to the main script**

In `Send-ContractorExpirationAlert.ps1`, append the following inside the `# ── Helper functions` block (after `Test-IsEmail`):

```powershell
function Get-AlertConfigPath {
    return Join-Path $PSScriptRoot "config\$env:USERNAME.xml"
}

function Read-AlertConfig {
    param([string]$Path = (Get-AlertConfigPath))
    if (-not (Test-Path $Path)) { return $null }
    return Import-Clixml -Path $Path
}

function Save-AlertConfig {
    param([hashtable]$Config, [string]$Path = (Get-AlertConfigPath))
    $dir = Split-Path $Path
    if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory | Out-Null }
    $Config | Export-Clixml -Path $Path
}
```

- [ ] **Step 2: Replace the placeholder BeforeAll in the config Describe block**

In `tests/Send-ContractorExpirationAlert.Tests.ps1`, replace the `BeforeAll` inside `Describe 'Save-AlertConfig / Read-AlertConfig'` with:

```powershell
    BeforeAll {
        function Get-AlertConfigPath {
            return Join-Path $PSScriptRoot "config\$env:USERNAME.xml"
        }

        function Read-AlertConfig {
            param([string]$Path = (Get-AlertConfigPath))
            if (-not (Test-Path $Path)) { return $null }
            return Import-Clixml -Path $Path
        }

        function Save-AlertConfig {
            param([hashtable]$Config, [string]$Path = (Get-AlertConfigPath))
            $dir = Split-Path $Path
            if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory | Out-Null }
            $Config | Export-Clixml -Path $Path
        }
    }
```

- [ ] **Step 3: Run tests — verify all 11 pass**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 11, Failed: 0`

- [ ] **Step 4: Commit**

```powershell
git add Send-ContractorExpirationAlert.ps1 tests/Send-ContractorExpirationAlert.Tests.ps1
git commit -m "Add config helpers (Read-AlertConfig, Save-AlertConfig) with Pester tests"
```

---

## Task 5: Add Get-GraphAccessToken and Send-GraphMail functions

**Files:**
- Modify: `Send-ContractorExpirationAlert.ps1`

These functions call external APIs and cannot be unit-tested with Pester without mocking infrastructure. They are validated in Task 9 (integration test).

- [ ] **Step 1: Add Get-GraphAccessToken to the helper functions block**

Append after `Save-AlertConfig`:

```powershell
function Get-GraphAccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [SecureString]$ClientSecret
    )
    $ptr    = [Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($ClientSecret)
    $secret = [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
    [Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr)

    $response = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body @{
            grant_type    = 'client_credentials'
            client_id     = $ClientId
            client_secret = $secret
            scope         = 'https://graph.microsoft.com/.default'
        } `
        -ErrorAction Stop

    return $response.access_token
}
```

- [ ] **Step 2: Add Send-GraphMail to the helper functions block**

Append after `Get-GraphAccessToken`:

```powershell
function Send-GraphMail {
    param(
        [string]   $AccessToken,
        [string]   $FromAddress,
        [string[]] $To,
        [string[]] $Cc,
        [string]   $Subject,
        [string]   $HtmlBody
    )
    $toRecipients = @($To | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    $ccRecipients = @(if ($Cc) { $Cc | ForEach-Object { @{ emailAddress = @{ address = $_ } } } })

    $message = [ordered]@{
        subject      = $Subject
        body         = @{ contentType = 'HTML'; content = $HtmlBody }
        toRecipients = $toRecipients
    }
    if ($ccRecipients.Count -gt 0) { $message['ccRecipients'] = $ccRecipients }

    $payload = @{ message = $message; saveToSentItems = $false } | ConvertTo-Json -Depth 10

    Invoke-RestMethod -Method Post `
        -Uri "https://graph.microsoft.com/v1.0/users/$FromAddress/sendMail" `
        -Headers @{ Authorization = "Bearer $AccessToken"; 'Content-Type' = 'application/json' } `
        -Body $payload `
        -ErrorAction Stop | Out-Null
}
```

- [ ] **Step 3: Run existing Pester tests to confirm no regressions**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 11, Failed: 0`

- [ ] **Step 4: Commit**

```powershell
git add Send-ContractorExpirationAlert.ps1
git commit -m "Add Get-GraphAccessToken and Send-GraphMail helper functions"
```

---

## Task 6: Add Invoke-SetupWizard function

**Files:**
- Modify: `Send-ContractorExpirationAlert.ps1`

- [ ] **Step 1: Add Invoke-SetupWizard to the helper functions block**

Append after `Send-GraphMail`:

```powershell
function Invoke-SetupWizard {
    param([string]$ConfigPath)

    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor Cyan
    Write-Host '  Contractor Alert - First-Time Setup' -ForegroundColor Cyan
    Write-Host ('=' * 70) -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Complete these steps in the Entra ID portal before continuing:' -ForegroundColor Yellow
    Write-Host '  1. Register an application (any name, no redirect URI needed)'
    Write-Host '  2. API permissions -> Add -> Microsoft Graph -> Application permissions -> Mail.Send'
    Write-Host '  3. Grant admin consent for your organization'
    Write-Host '  4. Certificates & secrets -> New client secret -> copy the value (shown once)'
    Write-Host '  5. Note the Tenant ID and Client ID from the app Overview page'
    Write-Host ''

    do {
        $tenantId = Read-Host 'Enter Tenant ID'
        if (-not (Test-IsGuid $tenantId)) { Write-Warning 'Invalid format. Expected: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' }
    } until (Test-IsGuid $tenantId)

    do {
        $clientId = Read-Host 'Enter Client ID'
        if (-not (Test-IsGuid $clientId)) { Write-Warning 'Invalid format. Expected: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' }
    } until (Test-IsGuid $clientId)

    $clientSecret = Read-Host 'Enter Client Secret' -AsSecureString

    do {
        $fromAddress = Read-Host 'Enter From Address (sending mailbox UPN, e.g. alerts@contoso.com)'
        if (-not (Test-IsEmail $fromAddress)) { Write-Warning 'Invalid format. Expected a valid email address.' }
    } until (Test-IsEmail $fromAddress)

    Write-Host ''
    Write-Host '[INFO] Testing credentials against Entra ID...' -ForegroundColor Cyan
    try {
        $null = Get-GraphAccessToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
        Write-Host '[INFO] Credentials verified successfully.' -ForegroundColor Green
    } catch {
        Write-Warning "Credential test failed: $($_.Exception.Message)"
        Write-Warning 'Config not saved. Correct the values and re-run with -Setup.'
        exit 1
    }

    $config = @{
        TenantId     = $tenantId
        ClientId     = $clientId
        ClientSecret = $clientSecret
        FromAddress  = $fromAddress
    }
    Save-AlertConfig -Config $config -Path $ConfigPath
    Write-Host "[INFO] Config saved to: $ConfigPath" -ForegroundColor Green
    Write-Host ''
}
```

- [ ] **Step 2: Run existing Pester tests to confirm no regressions**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 11, Failed: 0`

- [ ] **Step 3: Commit**

```powershell
git add Send-ContractorExpirationAlert.ps1
git commit -m "Add Invoke-SetupWizard function"
```

---

## Task 7: Update Build-HtmlEmail — remove SmtpServer parameter

**Files:**
- Modify: `Send-ContractorExpirationAlert.ps1` (lines 218-315)

- [ ] **Step 1: Remove the SmtpServer parameter from the function signature**

In `Build-HtmlEmail`'s `param()` block, delete the line:
```powershell
        [string]   $SmtpServer
```

- [ ] **Step 2: Replace the footer line that references $SmtpServer**

Find line 308:
```powershell
        Send-ContractorExpirationAlert v$ScriptVersion &nbsp;|&nbsp; SMTP: $([System.Net.WebUtility]::HtmlEncode($SmtpServer))
```

Replace with:
```powershell
        Send-ContractorExpirationAlert v$ScriptVersion &nbsp;|&nbsp; Microsoft Graph
```

- [ ] **Step 3: Update the call site — remove the -SmtpServer argument**

Find lines 318-324:
```powershell
$htmlBody = Build-HtmlEmail `
    -ExpiringUsers     $expiringUsers `
    -NoExpirationUsers $noExpirationUsers `
    -WarnDays          $WarnDays `
    -GroupName         $GroupName `
    -ScriptVersion     $ScriptVersion `
    -SmtpServer        $SmtpServer
```

Replace with:
```powershell
$htmlBody = Build-HtmlEmail `
    -ExpiringUsers     $expiringUsers `
    -NoExpirationUsers $noExpirationUsers `
    -WarnDays          $WarnDays `
    -GroupName         $GroupName `
    -ScriptVersion     $ScriptVersion
```

- [ ] **Step 4: Run Pester tests to confirm no regressions**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 11, Failed: 0`

- [ ] **Step 5: Commit**

```powershell
git add Send-ContractorExpirationAlert.ps1
git commit -m "Remove SmtpServer from Build-HtmlEmail, update footer to show Microsoft Graph"
```

---

## Task 8: Update param block and help docs

**Files:**
- Modify: `Send-ContractorExpirationAlert.ps1` (lines 1-62)

- [ ] **Step 1: Replace the comment-based help block (lines 1-45)**

Replace the entire `<# ... #>` block with:

```powershell
<#
.SYNOPSIS
    Send advance expiration alerts for contractor domain accounts.
.DESCRIPTION
    Queries all members of an Active Directory security group for accounts whose
    AccountExpirationDate falls within the configured warning window, then sends
    a single HTML digest email via Microsoft Graph.
    Accounts with no expiration date set are flagged in a separate section.
    Intended to run as a daily scheduled task.

    On first run, or when -Setup is specified, an interactive wizard collects
    Entra ID app registration credentials and saves them to an encrypted per-user
    config file at .\config\<username>.xml.
.PARAMETER GroupName
    Name of the AD security group whose members are contractor accounts.
.PARAMETER To
    One or more primary recipient email addresses.
.PARAMETER Cc
    One or more CC recipient email addresses (optional).
.PARAMETER WarnDays
    Number of days ahead to warn about expiring accounts. Default: 14.
.PARAMETER Setup
    Force re-run of the first-time setup wizard, even if a config file already
    exists. Use this to update credentials when the client secret rotates.
.PARAMETER ReportOnly
    Dry-run switch. Logs what would be sent without actually sending any email.
    Config file is not required when this switch is used.
.PARAMETER Diagnostics
    Prints member counts and filter statistics to the console.
.PARAMETER LogHistory
    Number of days to retain transcript log files. Default: 30.
.PARAMETER SkipUpdateCheck
    Skip the automatic version update check at startup.
.EXAMPLE
    .\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it@contoso.com' -Setup
.EXAMPLE
    .\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it@contoso.com'
.EXAMPLE
    .\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'it@contoso.com' -ReportOnly -Diagnostics
.NOTES
    Requires: ActiveDirectory PowerShell module (RSAT).
    This script is read-only with respect to Active Directory.
    Entra ID app registration must have Mail.Send application permission with admin consent.
#>
```

- [ ] **Step 2: Replace the param block (lines 47-62)**

Replace the entire `[CmdletBinding()]` + `param(...)` block with:

```powershell
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]  [string]   $GroupName,
    [Parameter(Mandatory = $true)]  [string[]] $To,
                                    [string[]] $Cc,
                                    [int]      $WarnDays   = 14,
                                    [switch]   $Setup,
                                    [switch]   $ReportOnly,
                                    [switch]   $Diagnostics,
                                    [int]      $LogHistory  = 30,
                                    [switch]   $SkipUpdateCheck
)
```

- [ ] **Step 3: Run Pester tests to confirm no regressions**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 11, Failed: 0`

- [ ] **Step 4: Commit**

```powershell
git add Send-ContractorExpirationAlert.ps1
git commit -m "Remove SMTP parameters, add -Setup switch, update help docs"
```

---

## Task 9: Wire up main script flow — config load and Graph send block

**Files:**
- Modify: `Send-ContractorExpirationAlert.ps1`

- [ ] **Step 1: Add config load / setup wizard call after the startup banner**

After line 125 (`Write-Host ('=' * 70) -ForegroundColor Cyan`) and before the `# ── AD module check` comment, insert:

```powershell
# ── Config / setup ───────────────────────────────────────────────────────────
$configPath = Get-AlertConfigPath
if ($Setup) {
    Invoke-SetupWizard -ConfigPath $configPath
}

$config = $null
if (-not $ReportOnly) {
    $config = Read-AlertConfig -Path $configPath
    if (-not $config) {
        Write-Warning "No config file found for user '$env:USERNAME' -- run the script manually with -Setup to complete first-time configuration."
        if ($transcriptStarted) { Stop-Transcript | Out-Null }
        exit 1
    }
}

```

- [ ] **Step 2: Replace the send block**

Find and replace everything from the `$subject = ` line through the closing `}` of the `else` block (the block that currently contains `Send-MailMessage`). This replaces both the subject assignment and the entire mailParams/send block:

```powershell
$subject = "Contractor Account Expiration Alert - $($expiringUsers.Count) account(s) expiring within $WarnDays days"

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
        $token = Get-GraphAccessToken `
            -TenantId     $config.TenantId `
            -ClientId     $config.ClientId `
            -ClientSecret $config.ClientSecret
    } catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $hint = if ($statusCode -in 400, 401) {
            ' Your client secret may be wrong or expired -- re-run with -Setup to update the config.'
        } else { '' }
        Write-Warning "Failed to get access token: $($_.Exception.Message).$hint"
        if ($transcriptStarted) { Stop-Transcript | Out-Null }
        exit 1
    }

    try {
        Send-GraphMail `
            -AccessToken $token `
            -FromAddress $config.FromAddress `
            -To          $To `
            -Cc          $Cc `
            -Subject     $subject `
            -HtmlBody    $htmlBody
        Write-Host '[INFO] Email sent successfully.' -ForegroundColor Green
    } catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $hint = if ($statusCode -in 401, 403) {
            ' Your client secret may be wrong or expired -- re-run with -Setup to update the config.'
        } else { '' }
        Write-Warning "Failed to send email: $($_.Exception.Message).$hint"
        if ($transcriptStarted) { Stop-Transcript | Out-Null }
        exit 1
    }
}
```

- [ ] **Step 3: Run Pester tests to confirm no regressions**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 11, Failed: 0`

- [ ] **Step 4: Smoke test — ReportOnly mode (no config needed)**

```powershell
.\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'Admin@customer.com' -ReportOnly -Diagnostics -SkipUpdateCheck
```

Expected: Runs without error, prints diagnostics and `[REPORT ONLY]` output. No prompt for config.

- [ ] **Step 5: Commit**

```powershell
git add Send-ContractorExpirationAlert.ps1
git commit -m "Wire up Graph send path and config load in main script flow"
```

---

## Task 10: Bump version to 2.0.0, update README, final commit

**Files:**
- Modify: `Send-ContractorExpirationAlert.ps1`
- Modify: `README.md`

- [ ] **Step 1: Bump version**

In `Send-ContractorExpirationAlert.ps1`, change:
```powershell
$ScriptVersion   = '1.1.0'
```
to:
```powershell
$ScriptVersion   = '2.0.0'
```

- [ ] **Step 2: Update README**

In `README.md`, update the Parameters section to remove `SmtpServer`, `SmtpPort`, `Credential`, `UseSSL`, `FromAddress` and add `Setup`. Update the Usage/Examples section to reflect the new Graph-based invocation. Add a Prerequisites section describing the Entra ID app registration steps (same steps printed by the wizard).

- [ ] **Step 3: Run Pester tests one final time**

```powershell
Invoke-Pester .\tests\Send-ContractorExpirationAlert.Tests.ps1 -Output Detailed
```

Expected: `Tests Passed: 11, Failed: 0`

- [ ] **Step 4: Final commit**

```powershell
git add Send-ContractorExpirationAlert.ps1 README.md
git commit -m "v2.0.0 - Replace SMTP with Microsoft Graph send"
```

- [ ] **Step 5: Push to GitHub**

```powershell
git push
```

---

## Task 11: Integration test — first-run setup and live send

This task requires the Entra ID app registration to exist with `Mail.Send` application permission and admin consent granted.

- [ ] **Step 1: Run setup wizard**

```powershell
.\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'Admin@customer.com' -Setup -ReportOnly -SkipUpdateCheck
```

Expected: Wizard prints prerequisites, prompts for Tenant ID, Client ID, Client Secret, From Address. After entering values, prints `[INFO] Credentials verified successfully.` and `[INFO] Config saved to: ...\config\<username>.xml`. Then runs in ReportOnly mode.

Verify config file exists:
```powershell
Test-Path ".\config\$env:USERNAME.xml"
```
Expected: `True`

- [ ] **Step 2: Run without -Setup to confirm config is loaded silently**

```powershell
.\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'Admin@customer.com' -ReportOnly -Diagnostics -SkipUpdateCheck
```

Expected: No setup wizard. Diagnostics and ReportOnly output as before.

- [ ] **Step 3: Run live send**

```powershell
.\Send-ContractorExpirationAlert.ps1 -GroupName 'Contractors' -To 'Admin@customer.com' -SkipUpdateCheck
```

Expected: `[INFO] Sending email to: Admin@customer.com...` followed by `[INFO] Email sent successfully.` Verify email arrives in the inbox.
