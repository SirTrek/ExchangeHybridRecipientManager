<#
.SYNOPSIS
Grants or revokes access to a shared mailbox in Exchange Online.

.DESCRIPTION
Mailbox permissions for a mailbox that lives in Exchange Online are stored in
Exchange Online. Entra Connect does not sync them upward, and the on-premises
Recipient Management tools have no cmdlet that can grant rights on a mailbox
which does not exist on-premises - so this cannot be done from the web UI in
this repository, which only ever talks to Active Directory.

This script is deliberately separate from that web UI. The UI binds to localhost
and has no authentication of any kind; holding a live Exchange Online session
inside it would extend that unauthenticated surface across the whole tenant.
Run this from a prompt instead, when you need it.

Three distinct rights are available, and they are not interchangeable:

  FullAccess   - open the mailbox and read/manage its contents. Does NOT allow
                 sending as the mailbox.
  SendAs       - mail appears to come from the shared mailbox itself.
  SendOnBehalf - mail appears as "<user> on behalf of <mailbox>".

.PARAMETER Mailbox
The shared mailbox, by primary SMTP address, alias or display name.

.PARAMETER User
One or more people to grant or revoke. Accepts UPNs, SMTP addresses or aliases.

.PARAMETER FullAccess
Grant (or with -Remove, revoke) Full Access.

.PARAMETER SendAs
Grant (or with -Remove, revoke) Send As.

.PARAMETER SendOnBehalf
Grant (or with -Remove, revoke) Send on Behalf.

.PARAMETER Remove
Revoke the named rights instead of granting them.

.PARAMETER AutoMapping
Whether Outlook should mount the mailbox automatically for people given Full
Access. Defaults to $true, matching Add-MailboxPermission. Turn it off for
mailboxes with many delegates or a large store, where auto-mounting is a common
cause of slow Outlook profiles: -AutoMapping:$false. Only affects FullAccess.

Note that changing this for someone who already holds Full Access requires the
permission to be removed and re-added; Exchange does not update it in place.

.PARAMETER NoConnect
Skip the Connect-ExchangeOnline step and use the session already in this shell.

.EXAMPLE
.\Grant-SharedMailboxAccess.ps1 -Mailbox hr@contoso.com -User bob@contoso.com -FullAccess -SendAs

.EXAMPLE
.\Grant-SharedMailboxAccess.ps1 -Mailbox hr@contoso.com -User bob@contoso.com,sue@contoso.com -FullAccess -AutoMapping:$false

.EXAMPLE
.\Grant-SharedMailboxAccess.ps1 -Mailbox hr@contoso.com -User bob@contoso.com -FullAccess -SendAs -Remove

.EXAMPLE
.\Grant-SharedMailboxAccess.ps1 -Mailbox hr@contoso.com -User bob@contoso.com -FullAccess -WhatIf

.NOTES
Part of https://github.com/SirTrek/ExchangeHybridRecipientManager
Requires the ExchangeOnlineManagement module and an account with Exchange
recipient management rights in the tenant.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,

    [Parameter(Mandatory, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string[]]$User,

    [switch]$FullAccess,
    [switch]$SendAs,
    [switch]$SendOnBehalf,
    [switch]$Remove,
    [bool]$AutoMapping = $true,
    [switch]$NoConnect
)

$ErrorActionPreference = 'Stop'

if (-not ($FullAccess -or $SendAs -or $SendOnBehalf)) {
    throw "Specify at least one of -FullAccess, -SendAs or -SendOnBehalf. They are separate rights: Full Access alone does not let anyone send as the mailbox."
}

# Only import when the cmdlets are not already present. Re-importing over a live
# session is wasteful, and it would shadow anything the caller has already set up.
if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
        throw "The ExchangeOnlineManagement module is not installed. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    Import-Module ExchangeOnlineManagement
}

if (-not $NoConnect) {
    $Existing = @(Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' })
    if ($Existing) {
        Write-Verbose "Reusing the existing Exchange Online session ($($Existing[0].UserPrincipalName))."
    }
    else {
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

# Resolve the mailbox first, so a typo fails here rather than half way through a
# list of users. A mailbox created on-premises will not be visible until Entra
# Connect has synced and Exchange Online has provisioned it, which is the usual
# reason this lookup comes up empty right after creating one.
$Target = Get-Mailbox -Identity $Mailbox -ErrorAction SilentlyContinue
if (-not $Target) {
    throw "No mailbox in Exchange Online matched '$Mailbox'. If it was just created on-premises, it will not appear until directory sync has run and the mailbox has been provisioned."
}

if ($Target.RecipientTypeDetails -ne 'SharedMailbox') {
    Write-Warning "'$($Target.PrimarySmtpAddress)' is a $($Target.RecipientTypeDetails), not a SharedMailbox. Proceeding, but check this is the object you meant."
}

$Verb = if ($Remove) { "Revoked" } else { "Granted" }
$Results = @()

foreach ($Account in $User) {

    $Recipient = Get-Recipient -Identity $Account -ErrorAction SilentlyContinue
    if (-not $Recipient) {
        Write-Warning "Skipping '$Account': no such recipient in Exchange Online."
        $Results += [pscustomobject]@{
            Mailbox = $Target.PrimarySmtpAddress; User = $Account
            Right   = '(all)'; Result = 'Skipped - recipient not found'
        }
        continue
    }
    $Who = $Recipient.PrimarySmtpAddress

    # Each right is applied independently so one failure does not abandon the
    # others - a tenant can legitimately refuse SendAs while allowing FullAccess.
    $Jobs = @()
    if ($FullAccess)   { $Jobs += 'FullAccess' }
    if ($SendAs)       { $Jobs += 'SendAs' }
    if ($SendOnBehalf) { $Jobs += 'SendOnBehalf' }

    foreach ($Right in $Jobs) {
        $What = "$Verb $Right on $($Target.PrimarySmtpAddress)"
        if (-not $PSCmdlet.ShouldProcess($Who, $What)) { continue }

        try {
            switch ($Right) {
                'FullAccess' {
                    if ($Remove) {
                        Remove-MailboxPermission -Identity $Target.PrimarySmtpAddress -User $Who `
                            -AccessRights FullAccess -Confirm:$false | Out-Null
                    }
                    else {
                        Add-MailboxPermission -Identity $Target.PrimarySmtpAddress -User $Who `
                            -AccessRights FullAccess -AutoMapping $AutoMapping -Confirm:$false | Out-Null
                    }
                }
                'SendAs' {
                    # Send As in Exchange Online is Add-RecipientPermission, not
                    # Add-MailboxPermission, and its parameter is -Trustee.
                    if ($Remove) {
                        Remove-RecipientPermission -Identity $Target.PrimarySmtpAddress -Trustee $Who `
                            -AccessRights SendAs -Confirm:$false | Out-Null
                    }
                    else {
                        Add-RecipientPermission -Identity $Target.PrimarySmtpAddress -Trustee $Who `
                            -AccessRights SendAs -Confirm:$false | Out-Null
                    }
                }
                'SendOnBehalf' {
                    # This one is a property of the mailbox, edited with the
                    # Add/Remove hashtable form so the other delegates survive.
                    $Edit = if ($Remove) { @{Remove = $Who } } else { @{Add = $Who } }
                    Set-Mailbox -Identity $Target.PrimarySmtpAddress -GrantSendOnBehalfTo $Edit | Out-Null
                }
            }

            $Results += [pscustomobject]@{
                Mailbox = $Target.PrimarySmtpAddress; User = $Who
                Right   = $Right; Result = $Verb
            }
        }
        catch {
            Write-Warning "$Right for ${Who}: $($_.Exception.Message)"
            $Results += [pscustomobject]@{
                Mailbox = $Target.PrimarySmtpAddress; User = $Who
                Right   = $Right; Result = "Failed - $($_.Exception.Message)"
            }
        }
    }
}

if ($FullAccess -and -not $Remove -and -not $AutoMapping) {
    Write-Host "Auto-mapping is off: delegates must add the mailbox to Outlook themselves (File > Account Settings > Change > More Settings > Advanced)."
}

Write-Host "Permission changes can take a few minutes to take effect, and Outlook may need a restart."
$Results
