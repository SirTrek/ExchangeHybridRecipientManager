# Exchange Hybrid Recipient Manager

A local web UI for managing Exchange Online recipients in a hybrid setup, for organisations
that have **decommissioned their last on-premises Exchange server** but still have
directory-synced objects whose mail attributes only Active Directory can change.

It runs entirely on your own machine: a PowerShell `HttpListener` serves the pages and calls
the Exchange Management Tools locally. There is no Exchange server, no IIS, no database and
no service to install.

> **Credit.** This is a maintained continuation of
> [spgoodman/ExchangeRecipientAdmin](https://github.com/spgoodman/ExchangeRecipientAdmin)
> by **Steve Goodman**, whose original work this builds on and whose commits are preserved
> in the history here. That project has had no commits on `main` since 2022 and none anywhere
> since mid-2024, with issues unanswered, so this fork was published to keep it usable and
> findable. It remains MIT licensed under Steve's original copyright. Improvements are still
> offered upstream — see
> [PR #12](https://github.com/spgoodman/ExchangeRecipientAdmin/pull/12).

## Why this exists

Remove your last Exchange server and the supported way to edit a synced user's proxy
addresses, remote routing address or GAL visibility is raw ADSI edits or hand-written
PowerShell. Microsoft's answer is the *Management Tools-only* install of Exchange 2019
(CU12+), which gives you `Get-RemoteMailbox`, `Set-RemoteMailbox` and friends without a
server. This puts a web UI on top of those cmdlets.

## What it manages

| Section | Actions |
|---|---|
| **Remote Mailboxes** | Create a new shared, room or equipment mailbox (AD object and mailbox together); enable a mailbox for an existing AD user as regular, shared, room or equipment; convert an existing mailbox between those types; edit display name, alias and remote routing address; add and remove proxy addresses; hide from or show in the GAL; disable the remote mailbox |
| **Distribution Groups** | Create a group; mail-enable an existing AD group; edit display name and alias; mail-disable (keeps the AD group); delete |
| **Contacts** | Create, edit external address and display name, delete |
| **Email Address Policies** | Create with an address template, change priority and recipient filter, delete |
| **Accepted Domains** | Add, change domain type, delete |

Every list page has live search and sortable columns. Destructive actions sit behind a
type-to-confirm control that is **re-verified server-side** — the browser guard is treated
as a convenience, not a control.

## Shared, room and equipment mailboxes

There are two entry points, and they are not interchangeable:

- **New shared mailbox** calls `New-RemoteMailbox`, which creates the Active Directory
  object *and* its remote mailbox together. Use this when no AD account exists yet.
- **Enable remote mailbox** calls `Enable-RemoteMailbox`, which mail-enables an account you
  already have. It cannot create one, so it is no use for a brand-new shared mailbox.

Only shared, room and equipment mailboxes can be created from the UI. `New-RemoteMailbox`
makes `-UserPrincipalName` and `-Password` mandatory in its *Regular* parameter set and
optional in the other three, and this tool deliberately never handles a password — the
enable form is a GET and every request's full query string is written to the console and the
in-memory log. For a regular mailbox, create the AD user however you normally do and then use
**Enable remote mailbox**.

Shared, room and equipment mailboxes need no licence under 50 GB. Disable the backing AD
account so nobody can sign in as the mailbox.

Each mailbox's edit page also has a **Mailbox Type** card that calls
`Set-RemoteMailbox -Type`. That exists because converting a mailbox in the Exchange Online
admin centre changes the *cloud* mailbox only: the type lives in `msExchRemoteRecipientType`
on the on-premises AD object, which the cloud-side change never touches. Until both sides are
set, `Get-RemoteMailbox` keeps reporting the old type and anything driven off those
attributes disagrees with what Exchange Online is actually serving.

## Granting access to a shared mailbox

Not in the web UI, and not an oversight. Full Access, Send As and Send on Behalf for a
mailbox that lives in Exchange Online are stored **in Exchange Online**; Entra Connect does
not sync mailbox permissions upward, and the Recipient Management snap-in has no
`Add-MailboxPermission` or `Add-RecipientPermission` to call. Holding a live Exchange Online
session inside a web app that binds to localhost with no authentication would also extend
that unauthenticated surface across the whole tenant.

So it ships as a separate script you run from a prompt:

```powershell
.\Grant-SharedMailboxAccess.ps1 -Mailbox hr@contoso.com -User bob@contoso.com -FullAccess -SendAs
.\Grant-SharedMailboxAccess.ps1 -Mailbox hr@contoso.com -User bob@contoso.com,sue@contoso.com -FullAccess -AutoMapping:$false
.\Grant-SharedMailboxAccess.ps1 -Mailbox hr@contoso.com -User bob@contoso.com -FullAccess -SendAs -Remove
```

It reuses an existing `Connect-ExchangeOnline` session if there is one, supports `-WhatIf`,
applies each right independently so one refusal doesn't abandon the rest, and skips unknown
recipients with a warning rather than failing the batch. The three rights are separate:
Full Access alone does **not** let anyone send as the mailbox. `-AutoMapping:$false` is worth
knowing for mailboxes with many delegates, where Outlook auto-mounting is a common cause of
slow profiles; changing it later needs the permission removed and re-added, as Exchange does
not update it in place.

A newly created mailbox will not be visible to this script until directory sync has run and
Exchange Online has provisioned it.

## Requirements

- A domain-joined Windows machine with the **Exchange Server 2019 CU12 or later
  [Management Tools-only install](https://learn.microsoft.com/en-us/exchange/manage-hybrid-exchange-recipients-with-management-tools)**
- Windows PowerShell 5.1 (the Exchange snap-in is not available in PowerShell 7)
- An account with Exchange recipient management rights
- Local administrator rights — the Exchange assemblies will not load without elevation

## Running it

```powershell
git clone https://github.com/SirTrek/ExchangeHybridRecipientManager.git
cd ExchangeHybridRecipientManager
.\Create-Shortcut.ps1        # optional: puts a shortcut on your Desktop
```

Then launch `Launch-ExchangeRecipientAdmin.bat`, Run .\Create-Shortcut.ps1 (to create a desktop shortcut), or run the server directly:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Start-ExchangeRecipientAdminCenter.ps1
```

It picks a random localhost port, opens your browser, and logs each request to the console.
Click **Exit** in the UI to stop it.

## Security

The server **binds to `localhost` only and has no authentication of any kind**. That binding
is the entire access control model. Do not change it to `http://+:port/` or a machine name to
reach it from another computer: anyone who could reach that port could delete mailboxes,
delete accepted domains and break mail flow for whole namespaces, with no credentials. If you
need it from another machine, RDP to the machine running it, or install the Management Tools
where you actually want to work.

## Tests

```powershell
.\Test-ExchangeRecipientAdminCenter.ps1
```

177 assertions. It boots the real server script against stubbed Exchange cmdlets and drives
every route over real HTTP, so the whole request path is exercised rather than mocked. It
needs no Exchange, no Active Directory and no admin rights, and it never issues an LDAP query
— it is safe to run on the management box itself.

The stubs deliberately declare each cmdlet's **real parameter set** and nothing else, so
calling one with a parameter it does not have fails the way Exchange fails. Several genuine
bugs reached production precisely because an earlier, more permissive harness accepted
anything.

## Changes since the original

Beyond the features above, this fork fixes a number of defects in the original, several of
which caused silent data corruption:

- Values from Active Directory were substituted into HTML raw. A double quote in a display
  name truncated an input's `value` attribute, so opening a contact and saving silently
  renamed it; an apostrophe in a proxy address broke out of an inline `confirm('...')`
  string, which nulls the handler per spec and removed the confirmation prompt from a
  destructive action entirely.
- The query parser unescaped the whole query string *before* splitting it, so a `%3D` inside
  a value became a real delimiter — typing `a=b` in the SMTP box provisioned `a@domain` and
  reported success.
- `break` in the static-file branch exits the `switch`, not the file-serving block, so some
  requests returned the *previous* caller's page verbatim.
- `$HTMLRESPONSE`, `$HTML_RESULT` and `$Error` persisted across requests in the single
  long-lived loop scope, pinning stale banners — including outcome claims about destructive
  operations — onto later, unrelated page loads.
- Four modals posted to routes that did not exist, 404ing to a white page and discarding
  whatever had been typed. Edit pages for Contacts and Email Address Policies never existed;
  the Distribution Groups one existed under the wrong filename.
- `Set-AcceptedDomain -DomainName` is not a valid parameter, so changing a domain's type
  always failed. Email address policies could not be created without an address template, or
  updated when their priority was `Lowest`, and priorities outside Exchange's contiguous run
  were rejected with no guidance.

## Licence

MIT — see [LICENSE](LICENSE). Copyright (c) 2022 Steve Goodman.
