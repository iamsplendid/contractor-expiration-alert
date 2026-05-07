# Graph Send — Design Spec
**Date:** 2026-05-07
**Scope:** Replace `Send-MailMessage` (SMTP AUTH) with Microsoft Graph API send, add per-user encrypted config file and first-run setup wizard.

---

## Background

Microsoft is retiring Basic Auth for SMTP AUTH in Exchange Online (disabled by default December 2026, full removal second half 2027). Outbound port 25 is also blocked by AWS, making the current SMTP path non-functional. This change replaces the send section with Microsoft Graph `POST /users/{from}/sendMail` using OAuth 2.0 client credentials, which works over HTTPS/443 and has no deprecation timeline.

---

## Parameter Changes

**Removed:**
- `-SmtpServer` (mandatory)
- `-SmtpPort` (default 587)
- `-Credential` (PSCredential)
- `-UseSSL` (switch)
- `-FromAddress` (mandatory) — moves into config file

**Added:**
- `-Setup` (switch) — forces re-run of the setup wizard regardless of whether a config file exists

**Unchanged:** `-GroupName`, `-To`, `-Cc`, `-WarnDays`, `-ReportOnly`, `-Diagnostics`, `-LogHistory`, `-SkipUpdateCheck`

---

## Config File

**Path:** `.\config\$env:USERNAME.xml`

The `config\` subfolder is created automatically by the setup wizard if it does not exist. Each Windows user account gets its own config file, which is required because encryption is DPAPI-bound to the user who ran setup. If the scheduled task account changes, setup must be re-run under the new account.

**Contents:**

| Field | Type | Notes |
|---|---|---|
| `TenantId` | Plain text | Entra ID tenant GUID |
| `ClientId` | Plain text | App registration client GUID |
| `ClientSecret` | DPAPI-encrypted SecureString | Never stored in plaintext |
| `FromAddress` | Plain text | Sending mailbox UPN |

Serialized via `Export-Clixml` / `Import-Clixml`, which natively handles `SecureString` fields.

---

## First-Run Setup Wizard

Runs automatically when no config file exists for the current user, or explicitly when `-Setup` is passed.

**Flow:**
1. Print header explaining prerequisites: create app registration in Entra ID, grant `Mail.Send` application permission, grant admin consent, create a client secret.
2. Prompt: `Tenant ID` — validated against the standard 8-4-4-4-12 GUID regex pattern.
3. Prompt: `Client ID` — validated against the standard 8-4-4-4-12 GUID regex pattern.
4. Prompt: `Client Secret` — collected via `Read-Host -AsSecureString` (never plaintext in memory).
5. Prompt: `From Address` — validated as non-empty string containing `@`.
6. Test credentials immediately: POST token request to `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`. If it fails, print the error and exit with code 1 — config file is not saved.
7. On success: create `config\` subfolder if needed, save config via `Export-Clixml`, print confirmation.

---

## Graph Send Section

Replaces the `$mailParams` / `Send-MailMessage` block entirely.

**Step 1 — Get access token**
POST to `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`:
- `grant_type=client_credentials`
- `client_id={clientId}`
- `client_secret={clientSecret}` (decrypted from SecureString in memory only)
- `scope=https://graph.microsoft.com/.default`

Token held in memory only for the script run duration.

**Step 2 — Send mail**
POST to `https://graph.microsoft.com/v1.0/users/{fromAddress}/sendMail` with JSON payload:
- `subject` — existing subject string (unchanged)
- `body.contentType` — `HTML`
- `body.content` — output of existing `Build-HtmlEmail` (unchanged)
- `toRecipients` — mapped from `-To`
- `ccRecipients` — mapped from `-Cc` (omitted if empty)

The `-SmtpServer` parameter is removed from `Build-HtmlEmail`. The footer that previously displayed the SMTP server name is updated to show the static label `Microsoft Graph`.

Both calls use `Invoke-RestMethod`. No new modules or dependencies required.

---

## Error Handling

**Token request fails:**
Catch the exception. If the HTTP status is 400 or 401, append: "Your client secret may be wrong or expired — re-run with `-Setup` to update the config." Exit code 1 in all cases.

**Send call fails:**
Same pattern. 401/403 gets the secret-expiry hint. Any other error prints the raw message. Exit code 1.

**No config file, running unattended:**
If no config file exists and `-Setup` was not passed, print: "No config file found for user '$env:USERNAME' — run the script manually with `-Setup` to complete first-time configuration." Exit code 1. The script does not attempt to launch an interactive wizard in this case.

---

## Entra ID Prerequisites (manual, one-time)

Before running `-Setup`, an admin must:
1. Register an application in Entra ID (any name, no redirect URI needed)
2. Under API permissions → Add → Microsoft Graph → Application permissions → `Mail.Send`
3. Grant admin consent
4. Under Certificates & secrets → New client secret → copy the value immediately (shown once)
5. Note the Tenant ID and Client ID from the app's Overview page

These steps are summarized in the setup wizard output so the user can follow them without leaving the terminal.
