# Admin Guide — Design Spec
**Date:** 2026-05-07
**Output:** `docs/admin-guide.html`

---

## Purpose

A single-file HTML administrator guide for `Send-ContractorExpirationAlert.ps1`. Written for a non-technical audience (e.g. an office manager without an IT background) who is responsible for keeping the script running. Avoids jargon; explains all concepts in plain English. Includes a linked table of contents for navigation. Can be printed to PDF from any browser.

---

## Audience

A non-technical administrator — someone comfortable using a computer and following step-by-step instructions, but with no assumed knowledge of PowerShell, Active Directory, Azure, or software development concepts. Technical terms are either avoided or explained in plain language on first use.

---

## File Location

`docs/admin-guide.html` — in the repo alongside the script, version-controlled.

---

## Structure and Content

### Table of Contents (linked anchors)

1. Overview
2. Requirements
3. Setting Up the Security Permission (App Registration)
4. The Contractors Security Group
5. Script Options (Parameters)
6. First-Time Setup
7. Scheduled Task
8. Renewing the Security Credential (Client Secret Rotation)
9. Troubleshooting

---

### Section 1 — Overview

What the script does in plain English: once a day it checks a list of contractor accounts in your organization and sends an email warning if any of those accounts are about to expire. It looks 14 days ahead by default. If nothing is expiring, no email is sent.

Why it exists: contractor accounts in Active Directory expire on a set date. If no one notices in time, the contractor loses access to systems and work stops. This script gives advance warning so the account can be renewed before that happens.

How it works at a high level: the script runs automatically on a schedule, reads from a group called "Contractors" in your organization's directory, and sends an email via Microsoft 365.

---

### Section 2 — Requirements

Plain-English list of what is needed before the script can be used:

- A Windows server or computer that is connected to the company network and can see the company's user directory (Active Directory)
- PowerShell installed (it comes with Windows — no installation needed)
- A special "directory reading" tool called RSAT installed on the same computer (your IT person can install this; it lets PowerShell read user account information)
- A Microsoft 365 account for the script to send email from (a shared mailbox works — no license required)
- A security permission set up in Microsoft's cloud portal (Entra ID) — covered in Section 3

---

### Section 3 — Setting Up the Security Permission (App Registration)

Plain-English explanation: before the script can send email, it needs a "permission slip" from Microsoft. This is called an App Registration. It tells Microsoft 365 "this script is allowed to send email on behalf of your organization."

This only needs to be done once. Step-by-step instructions (non-technical language):

1. Sign in to the Microsoft Entra portal (entra.microsoft.com) with an administrator account
2. Go to "App registrations" and create a new one with any name (e.g. "Contractor Expiration Alert")
3. Go to "API permissions," add Microsoft Graph > Application permissions > Mail.Send, and grant admin consent
4. Go to "Certificates & secrets," create a new client secret, and copy the value immediately (it is only shown once) — note the expiration date
5. From the app's Overview page, copy the Application (client) ID and the Directory (tenant) ID

These three values — Tenant ID, Client ID, and the client secret — are entered during first-time setup (Section 6).

Plain-English explanation of what the client secret is: think of it like a password for the script. It proves to Microsoft that the script has permission to send email. It expires (typically after 1-2 years) and must be renewed — see Section 8.

---

### Section 4 — The Contractors Security Group

Plain-English explanation: in your organization's user directory, there is a group called "Contractors." The script monitors every account that is a member of this group. It does not monitor all accounts in the organization — only the ones in this specific group.

What this means practically:
- If a contractor account is not in the "Contractors" group, the script will not warn about it
- If a contractor account is in the group but has no expiration date set, the script flags it in a separate section of the email as a potential oversight
- Accounts that are disabled are ignored

Who manages group membership: your IT administrator. If a contractor should be monitored and they are not appearing in the alert emails, ask IT to add their account to the "Contractors" group and make sure an expiration date is set on the account.

---

### Section 5 — Script Options (Parameters)

Plain-English parameter reference table. For each parameter: what it is called, what it does, whether it is required, and an example. Avoid the word "parameter" where possible — use "option" instead.

Options covered:
- `-GroupName` — the name of the security group to monitor (required; default in your setup is "Contractors")
- `-To` — who receives the alert email (required; can be multiple addresses)
- `-Cc` — additional recipients who get a copy of the email (optional)
- `-WarnDays` — how many days ahead to look for expiring accounts (default: 14)
- `-Setup` — runs the first-time setup wizard; also used to update the security credential when it expires
- `-ReportOnly` — a "preview mode" that shows what the email would say without actually sending it; useful for testing
- `-Diagnostics` — prints extra detail to the screen while running; useful when troubleshooting
- `-LogHistory` — how many days of log files to keep (default: 30)
- `-SkipUpdateCheck` — skips checking GitHub for a newer version of the script

---

### Section 6 — First-Time Setup

Step-by-step instructions for running setup for the first time. Written as a numbered walkthrough:

1. Open PowerShell on the server where the script is installed
2. Run the script with the `-Setup` option (exact command shown)
3. The script will display a checklist of things to do in the Microsoft portal first (covered in Section 3)
4. It will ask for the Tenant ID — paste the value copied from the portal
5. It will ask for the Client ID — paste the value
6. It will ask for the Client Secret — paste the value (the cursor will not move; this is normal)
7. It will ask for the "From" address — the email address the alerts will be sent from (the shared mailbox)
8. The script tests the values against Microsoft's servers. If they are correct, it saves a configuration file and continues. If not, it displays an error and does not save.

Important note in plain English: the configuration file is locked to the Windows account that ran setup. If the scheduled task runs as a different account, setup must be re-run while logged in as (or running as) that account. The script will create a separate configuration file for each Windows user automatically.

---

### Section 7 — Scheduled Task

Plain-English explanation of the scheduled task: the script is set up to run automatically every day so that no one has to remember to run it manually.

Covers:
- The task runs daily at a configured time (e.g. 10:00 AM)
- It runs as a specific service account (a Windows user created just for this purpose)
- That service account must have been used to run `-Setup` first (see Section 6) — otherwise the script cannot find its configuration and will not run
- Log files are saved automatically in the `logs` folder inside the script's directory
- How to verify the task is running: open Task Scheduler, find "Contractor Expiration Alert," and check the "Last Run Time" and "Last Run Result" columns. A result of 0 means success.

---

### Section 8 — Renewing the Security Credential (Client Secret Rotation)

Plain-English explanation: the client secret (the "password" that lets the script send email) has an expiration date — 24 months from when it was created. When it expires, the script will stop sending emails and log an error.

What to do when it expires (step by step):
1. Sign in to the Microsoft Entra portal
2. Go to the app registration, then Certificates & secrets
3. Create a new client secret and copy the value immediately
4. Note the new expiration date for future reference
5. On the server, open PowerShell and run the script with `-Setup` as the service account (same as first-time setup)
6. Enter the same Tenant ID and Client ID as before, but paste the new client secret value
7. The script will test and save the new credential

Plain-English reminder: do not delete the old secret until the new one has been confirmed working.

Calendar reminder suggestion: note the expiration date somewhere visible (a calendar reminder 30 days before expiry is recommended).

---

### Section 9 — Troubleshooting

Common scenarios with plain-English explanations and fixes:

| Symptom | Plain-English cause | What to do |
|---|---|---|
| No email received and nothing is expiring | Script is working correctly | Nothing — check the log file to confirm it ran |
| No email received but accounts should be expiring | Script may not have run, or accounts are not in the Contractors group | Check Task Scheduler last run result; ask IT to verify group membership |
| Log shows the script failed to get a security token | Security credential is wrong or expired | Re-run with `-Setup` to update the credential (Section 8) |
| Log shows "no config file found" | Setup was not run as the task service account | Re-run `-Setup` as the service account (Section 6) |
| Log shows the email failed to send (permission error) | App registration is missing the email-sending permission | Ask IT to verify the app registration in the portal (Section 3) |
| Log shows it could not read the user directory | Script cannot reach Active Directory, or the group name is wrong | Verify the server is on the network and the group name matches exactly |
| Script stopped running after a Windows Update or server restart | Scheduled task may have been disabled | Open Task Scheduler and check if the task is still enabled |

---

## Style Notes

- Use "you" throughout (second person)
- Use "the script" not "the tool" or "the program"
- Use "security credential" or "security password" instead of "client secret" in headings; explain the technical term in parentheses on first use
- Use "Microsoft's cloud portal" instead of "Entra ID" on first mention; "the portal" thereafter
- Short paragraphs; numbered lists for procedures; bullet lists for reference information
- No code blocks in the HTML output — commands shown in a `<code>` style but explained in plain English next to them

---

## HTML Implementation Notes

- Self-contained single file (no external CSS or JS dependencies)
- Inline CSS only — styled to be clean and readable when printed
- Table of contents at top with anchor links to each section
- Each section has a clear `<h2>` heading with an `id` attribute for anchor navigation
- Print-friendly: no dark backgrounds, adequate contrast, page-break-friendly
