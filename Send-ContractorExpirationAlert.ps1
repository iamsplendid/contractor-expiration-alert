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

$ScriptVersion   = '2.0.0'
$ScriptUpdateUrl = 'https://raw.githubusercontent.com/iamsplendid/contractor-expiration-alert/master/Send-ContractorExpirationAlert.ps1'

# ── Helper functions ──────────────────────────────────────────────────────────
function Test-IsGuid {
    param([string]$Value)
    return $Value -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
}

function Test-IsEmail {
    param([string]$Value)
    return $Value -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

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

function Invoke-SetupWizard {
    param(
        [string]$ConfigPath,
        [bool]  $TranscriptStarted = $false
    )

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
        if ($TranscriptStarted) { Stop-Transcript | Out-Null }
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
                    exit $LASTEXITCODE
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

# ── Config / setup ───────────────────────────────────────────────────────────
$configPath = Get-AlertConfigPath
if ($Setup) {
    Invoke-SetupWizard -ConfigPath $configPath -TranscriptStarted $transcriptStarted
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
        [string]   $ScriptVersion
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
                <td style='padding:6px 10px; border:1px solid #ddd;'>$(if ($u.DisplayName) { [System.Net.WebUtility]::HtmlEncode($u.DisplayName) } else { '&mdash;' })</td>
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
                <td style='padding:6px 10px; border:1px solid #ddd;'>$(if ($u.DisplayName) { [System.Net.WebUtility]::HtmlEncode($u.DisplayName) } else { '&mdash;' })</td>
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
        Send-ContractorExpirationAlert v$ScriptVersion &nbsp;|&nbsp; Microsoft Graph
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
    -ScriptVersion     $ScriptVersion

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
        $hint = if ($statusCode -eq 401) {
            ' Your client secret may be wrong or expired -- re-run with -Setup to update the config.'
        } elseif ($statusCode -eq 403) {
            ' The app may lack Mail.Send permission or admin consent, or the From address may not be a licensed mailbox.'
        } else { '' }
        Write-Warning "Failed to send email: $($_.Exception.Message).$hint"
        if ($transcriptStarted) { Stop-Transcript | Out-Null }
        exit 1
    }
}

$elapsed = [int](New-TimeSpan -Start $start).TotalSeconds
Write-Host "[INFO] Completed in $elapsed second(s)." -ForegroundColor Cyan

if ($transcriptStarted) { Stop-Transcript | Out-Null }
