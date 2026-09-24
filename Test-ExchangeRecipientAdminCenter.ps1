<#
.Synopsis
Regression tests for Start-ExchangeRecipientAdminCenter.ps1
.Description
Boots the REAL, unmodified server script against stubbed Exchange cmdlets and
drives every route over HTTP. No Exchange, AD or admin rights required, and
nothing outside this folder is touched - the stubs shadow Get-PSSnapIn /
Add-PSSnapIn (so the snap-in guard passes), Get-Random (deterministic port) and
Start-Process (no browser). Everything else is the production code path.

The stub data deliberately contains characters a real directory allows and that
previously broke the app: an apostrophe, a double quote, an ampersand and HTML
markup in a DisplayName.
.Example
.\Test-ExchangeRecipientAdminCenter.ps1
.Example
.\Test-ExchangeRecipientAdminCenter.ps1 -Port 19200
#>
param(
    [string]$AppRoot = $PSScriptRoot,
    [int]$Port = 19100
)
$script = Join-Path $AppRoot "Start-ExchangeRecipientAdminCenter.ps1"
$base   = "http://localhost:$Port"

$prelude = {
    param($ScriptPath,$Port)
    # Values a real directory can legally hold: apostrophe, double quote, ampersand, markup.
    $DN   = 'Say "Hi" O''Brien & <b>Sons</b>'
    $SMTP = "o'brien@contoso.com"

    # ---------------------------------------------------------------------------
    # Exchange cmdlet stubs.
    #
    # These declare the REAL parameter set of each cmdlet and deliberately do NOT use
    # -ValueFromRemainingArguments, so calling one with a parameter it does not have
    # fails the way Exchange fails: "A parameter cannot be found that matches parameter
    # name 'X'". An earlier version of this harness accepted any argument, which let a
    # Set-AcceptedDomain -DomainName call pass every test and still break in production.
    #
    # If a stub and the real cmdlet ever disagree, the stub is what needs fixing.
    # ---------------------------------------------------------------------------

    function Get-PSSnapIn { [CmdletBinding()] param([string]$Name,[switch]$Registered) [pscustomobject]@{Name='s'} }
    function Add-PSSnapIn { [CmdletBinding()] param([Parameter(ValueFromRemainingArguments)]$Name) }
    function Start-Process { [CmdletBinding()] param([string]$FilePath) }
    function Get-Random { [CmdletBinding()] param([int]$Minimum,[int]$Maximum) $Port }

    function Get-User {
        [CmdletBinding()] param([string]$Identity,[string]$Filter,$ResultSize)
        [pscustomobject]@{ UserPrincipalName = "u@x.internal" } }

    function Get-RemoteMailbox {
        [CmdletBinding()] param([string]$Identity,[string]$Filter,$ResultSize)
        if ($Identity -and $Identity -notlike "*contoso.com*") { return $null }
        # A second fixture that is already a shared mailbox, so the type picker's
        # preselect is tested against something other than the default.
        $Rtd = if ($Identity -like "*shared*") { "RemoteSharedMailbox" } else { "RemoteUserMailbox" }
        [pscustomobject]@{ Name=$DN;DisplayName=$DN;Alias="ob";PrimarySmtpAddress=$SMTP
            RemoteRoutingAddress="ob@t.mail.onmicrosoft.com";RecipientTypeDetails=$Rtd
            WhenChanged="2026-08-26";HiddenFromAddressListsEnabled=$false
            EmailAddresses=@("SMTP:$SMTP","smtp:o'brien.alt@contoso.com") } }
    function Set-RemoteMailbox {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity,$EmailAddresses,
            [bool]$HiddenFromAddressListsEnabled,[string]$DisplayName,[string]$Alias,
            [string]$RemoteRoutingAddress,[string]$PrimarySmtpAddress,[string]$Type)
        # Logged only when -Type is actually passed, so the every-other-call paths
        # are unaffected and "the cmdlet was never reached" stays assertable.
        if ($PSBoundParameters.ContainsKey('Type')) {
            Add-Content -Path (Join-Path $env:TEMP "erac_setmbxtype_test.txt") -Value "$Identity|$Type"
        } }
    function Enable-RemoteMailbox {
        # -Archive is deliberately declared alongside the three type switches. It is a
        # real parameter of the real cmdlet, so if the handler ever splatted a
        # request-supplied name straight through, "mailboxtype=Archive" would quietly
        # provision an archive here rather than failing - which is what makes the
        # whitelist testable.
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity,[string]$PrimarySMTPAddress,
            [string]$RemoteRoutingAddress,[switch]$Archive,[switch]$Shared,[switch]$Room,[switch]$Equipment)
        $fired = @()
        foreach ($n in 'Archive','Shared','Room','Equipment') { if ($PSBoundParameters[$n]) { $fired += $n } }
        $sw = if ($fired) { $fired -join '+' } else { 'none' }
        Set-Content -Path (Join-Path $env:TEMP "erac_enable_test.txt") -Value "$PrimarySMTPAddress|$sw" }
    function Disable-RemoteMailbox {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity) }

    function Get-AcceptedDomain {
        [CmdletBinding()] param([string]$Identity)
        @([pscustomobject]@{Name="contoso.com";DomainName="contoso.com";DomainType="InternalRelay";Default=$true},
          [pscustomobject]@{Name="t.mail.onmicrosoft.com";DomainName="t.mail.onmicrosoft.com";DomainType="InternalRelay";Default=$false}) |
          Where-Object { -not $Identity -or $_.Name -eq $Identity } }
    # NOTE: no -DomainName here. The SMTP namespace is fixed at creation; only
    # New-AcceptedDomain accepts it. This omission is the point of the strict stubs.
    function Set-AcceptedDomain {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity,[string]$DomainType,
            [string]$Name,[switch]$MakeDefault,[bool]$AddressBookEnabled,[bool]$MatchSubDomains,[bool]$OutboundOnly) }
    function New-AcceptedDomain {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Name,[string]$DomainName,[string]$DomainType) }
    function Remove-AcceptedDomain {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity) }

    function Get-DistributionGroup {
        [CmdletBinding()] param([string]$Identity,[string]$Filter,$ResultSize)
        [pscustomobject]@{DisplayName=$DN;Alias="dl";PrimarySmtpAddress=$SMTP
            RecipientTypeDetails="MailUniversalDistributionGroup";WhenCreated="2026-01-01"} }
    function Set-DistributionGroup {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity,[string]$DisplayName,
            [string]$Alias,[string]$Name,[string]$PrimarySmtpAddress) }
    function New-DistributionGroup {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Name,[string]$PrimarySmtpAddress,
            [string]$Type,[string]$Alias,[string]$DisplayName,[string]$OrganizationalUnit,$Members) }
    function Enable-DistributionGroup {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity,[string]$PrimarySmtpAddress,[string]$Alias) }
    function Disable-DistributionGroup {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity)
        Add-Content -Path (Join-Path $env:TEMP "erac_deletes_test.txt") -Value "disablegroup:$Identity" }
    function Remove-DistributionGroup {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity)
        Add-Content -Path (Join-Path $env:TEMP "erac_deletes_test.txt") -Value "group:$Identity" }
    # The app queries AD over ADSI for groups that are not yet mail-enabled. Its
    # $ERAC_GroupLookup override is used here so the suite never issues an LDAP query -
    # a plain function stub would not work, because dot-sourcing the app defines the
    # real one afterwards and wins. The sentinel file drives the degraded path; an
    # environment variable would not, since the server runs in a separate process.
    $ERAC_GroupLookup = {
        if (Test-Path (Join-Path $env:TEMP "erac_groups_fail.txt")) {
            throw "The term 'Get-Group' is not recognized as the name of a cmdlet."
        }
        [pscustomobject]@{ Name = "Plain AD Group"; Dn = "CN=Plain AD Group,OU=Groups,DC=contoso,DC=com" }
    }

    function Get-MailContact {
        [CmdletBinding()] param([string]$Identity,[string]$Filter,$ResultSize)
        if ($Identity -and $Identity -notlike "*contoso.com*") { return $null }
        [pscustomobject]@{DisplayName=$DN;PrimarySmtpAddress=$SMTP;RecipientType="MailContact";ExternalEmailAddress="x@y.com"} }
    function Set-MailContact {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity,[string]$DisplayName,
            [string]$ExternalEmailAddress,[string]$Alias,[string]$Name,[bool]$HiddenFromAddressListsEnabled) }
    function New-MailContact {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Name,[string]$ExternalEmailAddress,
            [string]$DisplayName,[string]$Alias,[string]$OrganizationalUnit,[string]$FirstName,[string]$LastName) }
    function Remove-MailContact {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity)
        Add-Content -Path (Join-Path $env:TEMP "erac_deletes_test.txt") -Value "contact:$Identity" }

    function Get-EmailAddressPolicy {
        [CmdletBinding()] param([string]$Identity)
        # The default policy's priority really is "Lowest", not a number - which is what
        # an <input type="number"> could not hold, blanking the field on submit. Two
        # numerically-prioritised policies sit alongside it, so the valid range for an
        # edit is 1-2 and for a new policy 1-3.
        $all = @(
            [pscustomobject]@{Name="Default Policy";Priority="Lowest";RecipientFilter="Company -eq 'A&B'"}
            [pscustomobject]@{Name="Sales Policy";Priority=1;RecipientFilter="Department -eq 'Sales'"}
            [pscustomobject]@{Name="Ops Policy";Priority=2;RecipientFilter="Department -eq 'Ops'"}
        )
        if ($Identity) { return ($all | Where-Object { $_.Name -eq $Identity }) }
        return $all }
    function Set-EmailAddressPolicy {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity,$Priority,[string]$RecipientFilter,
            [string]$Name,$EnabledEmailAddressTemplates,$IncludedRecipients,[switch]$ForceUpgrade)
        # lets a test force a failure without making every call fail
        if ($RecipientFilter -eq 'TRIGGER-FAILURE') { throw "SIMULATED-FAILURE" }
        # mirrors the real -Priority type: an empty string is rejected outright
        if ($PSBoundParameters.ContainsKey('Priority')) {
            if ([string]::IsNullOrWhiteSpace([string]$Priority)) {
                throw 'Cannot bind parameter ''Priority''. Cannot convert value "" to type "EmailAddressPolicyPriority". The correct values include "Lowest" and positive 32-bit integers.'
            }
        }
        # mirrors the contiguity rule: only 1..(number of numeric policies) is accepted
        if ($PSBoundParameters.ContainsKey('Priority')) {
            $n = 0
            if ([int]::TryParse([string]$Priority, [ref]$n)) {
                if ($n -lt 1 -or $n -gt 2) {
                    throw "The specified priority `"$Priority`" isn't valid. The priority should be equivalent to or immediately higher or lower than the priority of an existing policy. Parameter name: Priority"
                }
            }
        }
        $sent = if ($PSBoundParameters.ContainsKey('Priority')) { "priority=$Priority" } else { "priority=<omitted>" }
        Add-Content -Path (Join-Path $env:TEMP "erac_setpolicy_test.txt") -Value "$Identity|$sent" }
    function New-EmailAddressPolicy {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Name,$Priority,[string]$RecipientFilter,
            $IncludedRecipients,$EnabledEmailAddressTemplates,[string]$EnabledPrimarySMTPAddressTemplate)
        # Mirrors the real constraint: without an address template Exchange cannot pick a
        # parameter set and fails before doing anything, whatever else you supply.
        if (-not $EnabledEmailAddressTemplates -and -not $EnabledPrimarySMTPAddressTemplate) {
            throw "Parameter set cannot be resolved using the specified named parameters."
        }
        Add-Content -Path (Join-Path $env:TEMP "erac_policy_test.txt") -Value "$Name|$EnabledEmailAddressTemplates|$Priority" }
    function Remove-EmailAddressPolicy {
        [CmdletBinding(SupportsShouldProcess)] param([string]$Identity)
        Add-Content -Path (Join-Path $env:TEMP "erac_deletes_test.txt") -Value "policy:$Identity" }

    . $ScriptPath
}

Remove-Item (Join-Path $env:TEMP "erac_groups_fail.txt") -ErrorAction SilentlyContinue
$job = Start-Job -ScriptBlock $prelude -ArgumentList $script,$Port
foreach($i in 1..40){Start-Sleep -m 400; try{Invoke-WebRequest "$base/" -UseBasicParsing -TimeoutSec 3|Out-Null;break}catch{}}

$pass=0; $fail=0
function Req($m,$p,$b){
  $a=@{Uri="$base$p";Method=$m;UseBasicParsing=$true;TimeoutSec=20}
  if($b){$a.Body=$b;$a.ContentType="application/x-www-form-urlencoded"}
  try{ $r=Invoke-WebRequest @a
       $c = if($r.Content -is [byte[]]){[Text.Encoding]::UTF8.GetString($r.Content)}else{[string]$r.Content}
       [pscustomobject]@{Status=[int]$r.StatusCode;Body=$c;CT=[string]$r.Headers['Content-Type']} }
  catch{ $st=-1; if($_.Exception.Response){$st=[int]$_.Exception.Response.StatusCode}
       [pscustomobject]@{Status=$st;Body="";CT=""} } }
function Check($name,$cond,$detail){
  if($cond){$script:pass++; "  PASS  $name"} else {$script:fail++; "  FAIL  $name  --> $detail"} }

"=== 1. Dead routes now implemented (were 404, data lost on submit) ==="
foreach($r in @(
  @{p="/addcontact";           b="displayName=New+Contact&externalEmailAddress=n%40v.com"},
  @{p="/adddistributiongroup"; b="groupName=New+Group&groupEmail=g%40contoso.com&groupType=Distribution"},
  @{p="/mailenablegroup";      b="existingGroup=Plain+AD+Group&groupEmail=g%40contoso.com"},
  @{p="/addemailaddresspolicy";b="policyName=New+Policy&priority=2&recipientFilter=x+-eq+%27y%27"})){
  $x = Req POST $r.p $r.b
  Check "POST $($r.p) -> 200 and renders a page" ($x.Status -eq 200 -and $x.Body.Length -gt 500) "status=$($x.Status) bytes=$($x.Body.Length)"
}

"`n=== 2. Malformed URLs no longer return blank 200s ==="
foreach($u in @("/editremotemailbox","/editremotemailbox?id=","/editaccepteddomain?id=","/editemailaddresspolicy?id=","/editcontact?nonsense","/editdistributiongroup?id=")){
  $x = Req GET $u
  Check "GET $u -> real page, not blank" ($x.Status -eq 200 -and $x.Body.Length -gt 200) "bytes=$($x.Body.Length)"
}
$x = Req GET "/remotemailboxes?username=a&username=b"
Check "duplicate query key no longer blanks the page" ($x.Body.Length -gt 500) "bytes=$($x.Body.Length)"

"`n=== 3. No stale page served for unrouted requests ==="
$a = Req GET "/accepteddomains"
$b = Req POST "/" ""
Check "POST / does not echo the previous page" ($b.Body -ne $a.Body) "identical bodies ($($b.Body.Length) bytes)"
$c = Req GET "//"
Check "GET // does not echo the previous page" ($c.Body -ne $a.Body) "identical bodies"

"`n=== 4. Silent SMTP truncation fixed ==="
Req GET "/remotemailboxes?username=u%40x.internal&primarysmtpaddress_local=a%3Db&primarysmtpaddress_accepteddomain=contoso.com&remoteroutingaddress_local=r&remoteroutingaddress_accepteddomain=t.mail.onmicrosoft.com" | Out-Null
Start-Sleep -Milliseconds 300
$got = (Get-Content (Join-Path $env:TEMP "erac_enable_test.txt") -EA SilentlyContinue)
Check "'a=b' reaches Exchange intact as a=b@contoso.com" ($got -eq 'a=b@contoso.com|none') "Exchange got '$got'"
$x = Req GET "/remotemailboxes?foo=bar"
Check "stray query no longer fires a provisioning attempt" ($x.Body -notmatch 'enabled as Remote Mailbox') "banner appeared"

"`n=== 5. Escaping: guards intact, fields not truncated ==="
$x = Req GET "/editremotemailbox?id=o%27brien%40contoso.com"
Check "apostrophe alias keeps its confirm guard" ($x.Body -match 'data-confirm="Remove smtp:o&#39;brien\.alt') "guard markup not found"
Check "no raw apostrophe left inside a confirm() literal" ($x.Body -notmatch "confirm\('[^']*o'brien") "raw apostrophe still in JS"
Check "disable form carries expected addr in data-expect" ($x.Body -match 'data-expect="o&#39;brien@contoso\.com"') "data-expect missing"
Check "DisplayName double quote is encoded, not truncating" ($x.Body -match 'value="Say &quot;Hi&quot; O&#39;Brien &amp; &lt;b&gt;Sons&lt;/b&gt;"') "value attribute not encoded"
$x = Req GET "/contacts"
Check "markup in a DisplayName is inert in list rows" ($x.Body -notmatch '<b>Sons</b>' -and $x.Body -match '&lt;b&gt;Sons&lt;/b&gt;') "markup rendered live"

"`n=== 6. Error banner shows only the current failure ==="
$b1 = (Req POST "/editemailaddresspolicy" "Name=Default+Policy&Priority=1&RecipientFilter=TRIGGER-FAILURE").Body
$b2 = (Req POST "/editemailaddresspolicy" "Name=Default+Policy&Priority=1&RecipientFilter=TRIGGER-FAILURE").Body
$n1 = ([regex]::Matches($b1,'SIMULATED-FAILURE')).Count
$n2 = ([regex]::Matches($b2,'SIMULATED-FAILURE')).Count
Check "error is not repeated on the 2nd failure" ($n1 -eq 1 -and $n2 -eq 1) "1st=$n1 2nd=$n2 occurrences"

"`n=== 7. Result banner does not leak to later page loads ==="
Req POST "/addaccepteddomain" "domainName=x.com&domainType=InternalRelay" | Out-Null
$x = Req GET "/remotemailboxes"
Check "GET /remotemailboxes is clean after an unrelated POST" ($x.Body -notmatch 'class="alert') "stale banner present"

"`n=== 8. Headers, favicon, placeholders ==="
$x = Req GET "/accepteddomains"
Check "HTML responses declare Content-Type" ($x.CT -match 'text/html') "got '$($x.CT)'"
$x = Req GET "/favicon.ico"
Check "favicon served instead of 404" ($x.Status -eq 200) "status=$($x.Status)"
foreach($u in @("/","/remotemailboxes","/distributiongroups","/contacts","/emailaddresspolicies","/accepteddomains")){
  $x = Req GET $u
  $leak = [regex]::Matches($x.Body,'<!--\s*\{[a-z_]+\}\s*-->') | ForEach-Object {$_.Value}
  Check "no unreplaced placeholder on $u" ($leak.Count -eq 0) "$($leak -join ',')"
}
$x = Req GET "/distributiongroups"
Check "mail-enable group picker is populated" ($x.Body -match '<option value="CN=Plain AD Group[^"]*">Plain AD Group</option>') "dropdown still empty"

"`n=== 9. Still safe, still alive ==="
foreach($u in @("/../Start-ExchangeRecipientAdminCenter.ps1","/%2e%2e/Start-ExchangeRecipientAdminCenter.ps1")){
  $x = Req GET $u
  Check "traversal blocked: $u" ($x.Status -eq 404) "status=$($x.Status)"
}
$x = Req GET "/accepteddomains"
Check "server alive after the whole battery" ($x.Status -eq 200) "status=$($x.Status)"

"`n=== 10. Delete: guarded, server-verified, and actually performed ==="
$delLog = Join-Path $env:TEMP "erac_deletes_test.txt"
if (Test-Path $delLog) { Remove-Item $delLog -Force }

# every edit page must offer a delete behind a type-to-confirm guard
foreach($pg in @(
  @{u="/editcontact?id=o%27brien%40contoso.com";        a="/deletecontact"},
  @{u="/editdistributiongroup?id=o%27brien%40contoso.com"; a="/deletedistributiongroup"},
  @{u="/editemailaddresspolicy?id=Default%20Policy";    a="/deleteemailaddresspolicy"})){
  $x = Req GET $pg.u
  Check "delete form present on $($pg.u.Split('?')[0])" ($x.Body -match [regex]::Escape('action="' + $pg.a + '"')) "form missing"
  Check "  guarded by type-to-confirm + disabled button" ($x.Body -match 'data-expect="' -and $x.Body -match 'data-confirm-button') "guard attributes missing"
}

# a WRONG confirmation must delete nothing
Req POST "/deletecontact" "id=o%27brien%40contoso.com&confirmText=not-the-address" | Out-Null
Req POST "/deletedistributiongroup" "id=o%27brien%40contoso.com&confirmText=wrong" | Out-Null
Req POST "/deleteemailaddresspolicy" "id=Default+Policy&confirmText=wrong" | Out-Null
Start-Sleep -Milliseconds 200
Check "wrong confirmation deletes nothing" (-not (Test-Path $delLog)) "a Remove-* cmdlet ran: $(Get-Content $delLog -EA SilentlyContinue)"
$x = Req POST "/deletecontact" "id=o%27brien%40contoso.com&confirmText=not-the-address"
Check "wrong confirmation explains itself" ($x.Body -match "didn&#39;t match|didn't match") "no explanatory banner"

# the CORRECT confirmation must actually delete
Req POST "/deletecontact" "id=o%27brien%40contoso.com&confirmText=o%27brien%40contoso.com" | Out-Null
Req POST "/deletedistributiongroup" "id=o%27brien%40contoso.com&confirmText=o%27brien%40contoso.com" | Out-Null
Req POST "/deleteemailaddresspolicy" "id=Default+Policy&confirmText=Default+Policy" | Out-Null
Start-Sleep -Milliseconds 200
$done = @(Get-Content $delLog -EA SilentlyContinue)
Check "correct confirmation deletes the contact" ($done -contains "contact:o'brien@contoso.com") "log=$($done -join ',')"
Check "correct confirmation deletes the group"   ($done -contains "group:o'brien@contoso.com")   "log=$($done -join ',')"
Check "correct confirmation deletes the policy"  ($done -contains "policy:Default Policy")       "log=$($done -join ',')"
if (Test-Path $delLog) { Remove-Item $delLog -Force }

# deleting an object that is already gone must not explode
$x = Req POST "/deletecontact" "id=ghost%40nowhere.com&confirmText=ghost%40nowhere.com"
Check "deleting a missing object degrades gracefully" ($x.Status -eq 200 -and $x.Body -match 'not found') "status=$($x.Status)"

"`n=== 11. Mail-disable a group: distinct from delete, and reversible ==="
$delLog = Join-Path $env:TEMP "erac_deletes_test.txt"
if (Test-Path $delLog) { Remove-Item $delLog -Force }

$x = Req GET "/editdistributiongroup?id=o%27brien%40contoso.com"
Check "mail-disable form offered alongside delete" ($x.Body -match [regex]::Escape('action="/disabledistributiongroup"') -and $x.Body -match [regex]::Escape('action="/deletedistributiongroup"')) "one of the two forms is missing"
Check "the two actions are visually distinguished" ($x.Body -match 'border-warning' -and $x.Body -match 'border-danger') "cards not differentiated"
Check "mail-disable is described as reversible" ($x.Body -match 'Mail-Enable a Group') "reversibility not explained"

Req POST "/disabledistributiongroup" "id=o%27brien%40contoso.com&confirmText=wrong" | Out-Null
Start-Sleep -Milliseconds 200
Check "wrong confirmation mail-disables nothing" (-not (Test-Path $delLog)) "cmdlet ran: $(Get-Content $delLog -EA SilentlyContinue)"

$x = Req POST "/disabledistributiongroup" "id=o%27brien%40contoso.com&confirmText=o%27brien%40contoso.com"
Start-Sleep -Milliseconds 200
$done = @(Get-Content $delLog -EA SilentlyContinue)
Check "correct confirmation mail-disables the group" ($done -contains "disablegroup:o'brien@contoso.com") "log=$($done -join ',')"
Check "mail-disable did NOT delete the AD group" ($done -notcontains "group:o'brien@contoso.com") "Remove-DistributionGroup was also called"
Check "result explains the AD group was kept" ($x.Body -match 'kept') "no reassurance in the banner"
if (Test-Path $delLog) { Remove-Item $delLog -Force }

"`n=== 12. Changing an accepted domain's type (the reported failure) ==="
$x = Req POST "/editaccepteddomain" "Name=contoso.com&DomainType=Authoritative"
Check "InternalRelay -> Authoritative succeeds" ($x.Body -match 'alert-success') "banner: $((([regex]::Match($x.Body,'<div class=\"alert[^\"]*\"[^>]*>(.*?)</div>')).Groups[1].Value) -replace '\s+',' ')"
Check "  no 'parameter cannot be found' error" ($x.Body -notmatch 'parameter cannot be found') "cmdlet rejected a parameter"
$x = Req GET "/editaccepteddomain?id=contoso.com"
Check "accepted domain name shown read-only (it is fixed at creation)" ($x.Body -match 'id="DomainNameDisplay"[^>]*disabled' -or $x.Body -match 'disabled[^>]*id="DomainNameDisplay"') "still editable"

"`n=== 13. Creating an email address policy (needs an address template) ==="
$polLog = Join-Path $env:TEMP "erac_policy_test.txt"
if (Test-Path $polLog) { Remove-Item $polLog -Force }

$x = Req GET "/emailaddresspolicies"
Check "modal collects an address template" ($x.Body -match 'name="addressTemplate"') "field missing"
Check "  template syntax is explained" ($x.Body -match '%m' -and $x.Body -match '%g') "no placeholder help"

$x = Req POST "/addemailaddresspolicy" "policyName=Test+Policy&priority=3&recipientFilter=RecipientType+-eq+%27UserMailbox%27&addressTemplate=%25m%40contoso.com"
Check "policy with a template is created" ($x.Body -match 'alert-success') "banner: $((([regex]::Match($x.Body,'<div class=\"alert[^\"]*\"[^>]*>(.*?)</div>')).Groups[1].Value) -replace '\s+',' ')"
Check "  no parameter-set error" ($x.Body -notmatch 'Parameter set cannot be resolved') "cmdlet could not resolve a set"
Start-Sleep -Milliseconds 200
$made = @(Get-Content $polLog -EA SilentlyContinue)
Check "  template reaches Exchange with an SMTP: prefix" ($made -contains "Test Policy|SMTP:%m@contoso.com|3") "got: $($made -join ',')"

# a template typed WITH the prefix must not end up double-prefixed
if (Test-Path $polLog) { Remove-Item $polLog -Force }
Req POST "/addemailaddresspolicy" "policyName=P2&priority=2&recipientFilter=x&addressTemplate=smtp%3A%25m%40contoso.com" | Out-Null
Start-Sleep -Milliseconds 200
$made = @(Get-Content $polLog -EA SilentlyContinue)
Check "an already-prefixed template is normalised, not doubled" ($made -contains "P2|SMTP:%m@contoso.com|2") "got: $($made -join ',')"

# omitting the template must explain itself rather than surfacing the raw cmdlet error
if (Test-Path $polLog) { Remove-Item $polLog -Force }
$x = Req POST "/addemailaddresspolicy" "policyName=P3&priority=1&recipientFilter=x&addressTemplate="
Check "missing template gives an actionable message" ($x.Body -match 'template is required') "banner: $((([regex]::Match($x.Body,'<div class=\"alert[^\"]*\"[^>]*>(.*?)</div>')).Groups[1].Value) -replace '\s+',' ')"
Check "  and creates nothing" (-not (Test-Path $polLog)) "a policy was created anyway"
if (Test-Path $polLog) { Remove-Item $polLog -Force }

"`n=== 14. Updating a policy whose priority is Lowest (the reported failure) ==="
$setLog = Join-Path $env:TEMP "erac_setpolicy_test.txt"
if (Test-Path $setLog) { Remove-Item $setLog -Force }

$x = Req GET "/editemailaddresspolicy?id=Default%20Policy"
Check "default policy shows Lowest and cannot be edited" ($x.Body -match 'value="Lowest"' -and $x.Body -match 'id="Priority"[^>]*disabled') "Lowest not shown read-only"

# resubmitting the form unchanged must not try to re-set Lowest
$x = Req POST "/editemailaddresspolicy" "Name=Default+Policy&RecipientFilter=Company+-eq+%27A%26B%27"
Check "unchanged priority is not resent" ($x.Body -notmatch 'Cannot convert value') "banner: $((([regex]::Match($x.Body,'<div class=\"alert[^\"]*\"[^>]*>(.*?)</div>')).Groups[1].Value) -replace '\s+',' ')"
Start-Sleep -Milliseconds 200
$sent = @(Get-Content $setLog -EA SilentlyContinue)
Check "  Priority omitted from the cmdlet call" ($sent -contains "Default Policy|priority=<omitted>") "got: $($sent -join ',')"

# a blank priority must not be sent as an empty string
if (Test-Path $setLog) { Remove-Item $setLog -Force }
$x = Req POST "/editemailaddresspolicy" "Name=Default+Policy&Priority=&RecipientFilter=Company+-eq+%27A%26B%27"
Check "blank priority does not reach the cmdlet" ($x.Body -notmatch 'Cannot convert value') "empty string was passed through"
Start-Sleep -Milliseconds 200
$sent = @(Get-Content $setLog -EA SilentlyContinue)
Check "  Priority omitted when left blank" ($sent -contains "Default Policy|priority=<omitted>") "got: $($sent -join ',')"

# a genuinely changed priority must still be applied
if (Test-Path $setLog) { Remove-Item $setLog -Force }
$x = Req POST "/editemailaddresspolicy" "Name=Sales+Policy&Priority=2&RecipientFilter=Department+-eq+%27Sales%27"
Start-Sleep -Milliseconds 200
$sent = @(Get-Content $setLog -EA SilentlyContinue)
Check "a changed priority IS sent" ($sent -contains "Sales Policy|priority=2") "got: $($sent -join ',')"
Check "  and reports success" ($x.Body -match 'alert-success') "no success banner"
if (Test-Path $setLog) { Remove-Item $setLog -Force }

"`n=== 15. Accepted domain delete lives only on the domain's own page ==="
$x = Req GET "/accepteddomains"
Check "list page has no per-row delete form" ($x.Body -notmatch 'action="/deleteaccepteddomain"') "a delete form is still rendered in the list"
Check "  and no leftover empty header column" (([regex]::Matches($x.Body,'<th scope="col">')).Count -eq 3) "header column count is $((([regex]::Matches($x.Body,'<th scope=\"col\">')).Count))"
Check "  domains still listed and clickable" ($x.Body -match 'href="/editaccepteddomain\?id=contoso\.com"') "rows missing"
$x = Req GET "/editaccepteddomain?id=contoso.com"
Check "delete still offered in the Danger Zone" ($x.Body -match 'action="/deleteaccepteddomain"' -and $x.Body -match 'Danger Zone') "delete no longer reachable at all"
$x = Req POST "/deleteaccepteddomain" "name=contoso.com"
Check "  and the route still works" ($x.Status -eq 200 -and $x.Body -match 'alert-success') "status=$($x.Status)"

"`n=== 16. Priority is offered as valid choices only (the reported failure) ==="
$setLog = Join-Path $env:TEMP "erac_setpolicy_test.txt"
$polLog = Join-Path $env:TEMP "erac_policy_test.txt"
foreach ($l in @($setLog,$polLog)) { if (Test-Path $l) { Remove-Item $l -Force } }

# 2 numeric policies exist -> edit range 1-2, new range 1-3
$x = Req GET "/editemailaddresspolicy?id=Sales%20Policy"
Check "edit page offers a priority dropdown" ($x.Body -match '<select[^>]*name="Priority"') "still a free-text field"
Check "  limited to the valid range 1-2" (($x.Body -match '<option[^>]*value="1"') -and ($x.Body -match '<option[^>]*value="2"') -and ($x.Body -notmatch '<option[^>]*value="3"')) "range wrong"
Check "  current value preselected" ($x.Body -match '<option selected value="1">') "not preselected"

$x = Req GET "/emailaddresspolicies"
Check "add modal offers a priority dropdown" ($x.Body -match '<select[^>]*name="priority"') "still a number box"
Check "  allows one past the end (1-3)" (($x.Body -match '<option[^>]*value="3"') -and ($x.Body -notmatch '<option[^>]*value="4"')) "range wrong"
Check "  defaults to appending at the end" ($x.Body -match '<option selected value="3">') "no sensible default"

# the reported value: 11 is way outside the run and must be refused before the cmdlet
$x = Req POST "/editemailaddresspolicy" "Name=Sales+Policy&Priority=11&RecipientFilter=Department+-eq+%27Sales%27"
Check "an out-of-range priority is refused" ($x.Body -match 'between 1 and 2') "banner: $((([regex]::Match($x.Body,'<div class=\"alert[^\"]*\"[^>]*>(.*?)</div>')).Groups[1].Value) -replace '\s+',' ')"
Start-Sleep -Milliseconds 200
Check "  and never reaches Exchange" (-not (Test-Path $setLog)) "Set-EmailAddressPolicy was called: $(Get-Content $setLog -EA SilentlyContinue)"

$x = Req POST "/addemailaddresspolicy" "policyName=P9&priority=11&recipientFilter=x&addressTemplate=%25m%40contoso.com"
Check "out-of-range priority refused when adding too" ($x.Body -match 'between 1 and 3') "banner: $((([regex]::Match($x.Body,'<div class=\"alert[^\"]*\"[^>]*>(.*?)</div>')).Groups[1].Value) -replace '\s+',' ')"
Start-Sleep -Milliseconds 200
Check "  and creates nothing" (-not (Test-Path $polLog)) "a policy was created"
foreach ($l in @($setLog,$polLog)) { if (Test-Path $l) { Remove-Item $l -Force } }

"`n=== 17. No stray output leaks into the page (the reported failure) ==="
# A bare string inside a PowerShell function becomes part of its RETURN VALUE, so it
# gets joined onto the HTML and rendered at the top of the page. Every HTML response
# must therefore begin exactly with the doctype - this catches the whole class, not
# just the one instance that was reported.
foreach ($u in @("/","/remotemailboxes","/distributiongroups","/contacts","/emailaddresspolicies",
                 "/accepteddomains","/editremotemailbox?id=o%27brien%40contoso.com",
                 "/editdistributiongroup?id=o%27brien%40contoso.com","/editcontact?id=o%27brien%40contoso.com",
                 "/editemailaddresspolicy?id=Sales%20Policy","/editaccepteddomain?id=contoso.com")) {
  $x = Req GET $u
  Check "page starts with the doctype: $u" ($x.Body.TrimStart().StartsWith('<!doctype')) "starts with: '$($x.Body.Substring(0,[Math]::Min(70,$x.Body.Length)) -replace '\s+',' ')'"
}

"`n=== 18. Group lookup failing degrades without breaking the page ==="
$failFile = Join-Path $env:TEMP "erac_groups_fail.txt"
Set-Content -Path $failFile -Value "1"
try {
  $x = Req GET "/distributiongroups"
  Check "page still renders when the AD lookup fails" ($x.Status -eq 200 -and $x.Body.Length -gt 2000) "status=$($x.Status) bytes=$($x.Body.Length)"
  Check "  and still begins with the doctype (no leaked log line)" ($x.Body.TrimStart().StartsWith('<!doctype')) "starts with: '$($x.Body.Substring(0,[Math]::Min(90,$x.Body.Length)) -replace '\s+',' ')'"
  Check "  the failure is not rendered into the markup" ($x.Body -notmatch "is not recognized") "the cmdlet error text reached the page"
  Check "  picker says so instead of being silently empty" ($x.Body -match 'could not load groups') "no explanation offered"
  Check "  the group tables are unaffected" ($x.Body -match 'href="/editdistributiongroup\?id=') "group rows missing"
} finally { Remove-Item $failFile -ErrorAction SilentlyContinue }

"`n=== 19. Mailbox type: provisioning and converting ==="
$enLog  = Join-Path $env:TEMP "erac_enable_test.txt"
$typLog = Join-Path $env:TEMP "erac_setmbxtype_test.txt"
foreach ($l in @($enLog,$typLog)) { if (Test-Path $l) { Remove-Item $l -Force } }

$base4 = "primarysmtpaddress_local=hr&primarysmtpaddress_accepteddomain=contoso.com&remoteroutingaddress_local=hr&remoteroutingaddress_accepteddomain=t.mail.onmicrosoft.com"

# the picker itself
$x = Req GET "/remotemailboxes"
Check "enable modal offers a mailbox type picker" ($x.Body -match '<select[^>]*name="mailboxtype"') "no type picker"
Check "  offers all four types" (@('Regular','Shared','Room','Equipment' | Where-Object { $x.Body -match "<option[^>]*value=`"$_`"" }).Count -eq 4) "missing options"
Check "  defaults to Regular" ($x.Body -match '<option selected value="Regular">') "Regular not preselected"

# Regular must not pass a switch at all - there is no -Regular on the real cmdlet
Req GET "/remotemailboxes?username=u%40x.internal&mailboxtype=Regular&$base4" | Out-Null
Start-Sleep -Milliseconds 300
Check "Regular passes no type switch" ((Get-Content $enLog -EA SilentlyContinue) -eq 'hr@contoso.com|none') "got '$(Get-Content $enLog -EA SilentlyContinue)'"

# and each real type maps to its own switch
foreach ($t in @('Shared','Room','Equipment')) {
  Remove-Item $enLog -Force -EA SilentlyContinue
  $x = Req GET "/remotemailboxes?username=u%40x.internal&mailboxtype=$t&$base4"
  Start-Sleep -Milliseconds 300
  Check "$t provisions with -$t" ((Get-Content $enLog -EA SilentlyContinue) -eq "hr@contoso.com|$t") "got '$(Get-Content $enLog -EA SilentlyContinue)'"
  Check "  and the banner says so" ($x.Body -match "as a $($t.ToLower()) mailbox") "banner did not mention the type"
}

# omitting the field entirely still provisions a regular mailbox (back-compat)
Remove-Item $enLog -Force -EA SilentlyContinue
Req GET "/remotemailboxes?username=u%40x.internal&$base4" | Out-Null
Start-Sleep -Milliseconds 300
Check "a request with no type field still works" ((Get-Content $enLog -EA SilentlyContinue) -eq 'hr@contoso.com|none') "got '$(Get-Content $enLog -EA SilentlyContinue)'"

# the type names a PARAMETER, so an arbitrary one must not reach the cmdlet.
# -Archive is real: without the whitelist this would silently provision an archive.
foreach ($bad in @('Archive','Bogus')) {
  Remove-Item $enLog -Force -EA SilentlyContinue
  $x = Req GET "/remotemailboxes?username=u%40x.internal&mailboxtype=$bad&$base4"
  Start-Sleep -Milliseconds 300
  Check "'$bad' is refused as a mailbox type" ($x.Body -match 'is not a valid mailbox type') "no refusal banner"
  Check "  and never reaches Exchange" (-not (Test-Path $enLog)) "Enable-RemoteMailbox ran: $(Get-Content $enLog -EA SilentlyContinue)"
}

# edit page: current type shown, picker preselected from it
$x = Req GET "/editremotemailbox?id=o%27brien%40contoso.com"
Check "edit page shows the AD recipient type" ($x.Body -match 'RemoteUserMailbox') "type not displayed"
Check "  picker preselects Regular for a user mailbox" ($x.Body -match '<option selected value="Regular">') "wrong preselect"
$x = Req GET "/editremotemailbox?id=shared%40contoso.com"
Check "  and preselects Shared for a shared mailbox" ($x.Body -match '<option selected value="Shared">') "wrong preselect"
Check "  showing the long form in the badge" ($x.Body -match 'RemoteSharedMailbox') "badge not updated"

# converting
$x = Req POST "/editremotemailbox" "id=o%27brien%40contoso.com&Action=changetype&MailboxType=Shared"
Start-Sleep -Milliseconds 200
Check "changetype calls Set-RemoteMailbox -Type Shared" (@(Get-Content $typLog -EA SilentlyContinue) -contains "o'brien@contoso.com|Shared") "got: $(Get-Content $typLog -EA SilentlyContinue)"
Check "  and says the sync is still pending" ($x.Body -match 'directory sync') "no sync caveat shown"

Remove-Item $typLog -Force -EA SilentlyContinue
$x = Req POST "/editremotemailbox" "id=o%27brien%40contoso.com&Action=changetype&MailboxType=Nonsense"
Start-Sleep -Milliseconds 200
Check "an invalid type is refused on convert too" ($x.Body -match 'is not a valid mailbox type') "no refusal banner"
Check "  and never reaches Exchange" (-not (Test-Path $typLog)) "Set-RemoteMailbox ran: $(Get-Content $typLog -EA SilentlyContinue)"
Check "  refusal still renders the edit page" ($x.Status -eq 200 -and $x.Body -match 'name="MailboxType"') "page broken"

# casing is the operator's, canonical casing is ours
Remove-Item $typLog -Force -EA SilentlyContinue
Req POST "/editremotemailbox" "id=o%27brien%40contoso.com&Action=changetype&MailboxType=shared" | Out-Null
Start-Sleep -Milliseconds 200
Check "lower-case 'shared' is canonicalised to Shared" (@(Get-Content $typLog -EA SilentlyContinue) -contains "o'brien@contoso.com|Shared") "got: $(Get-Content $typLog -EA SilentlyContinue)"

foreach ($l in @($enLog,$typLog)) { if (Test-Path $l) { Remove-Item $l -Force } }

"`n================ $pass passed, $fail failed ================"
try{Invoke-WebRequest "$base/exit" -UseBasicParsing -TimeoutSec 5|Out-Null}catch{}
Start-Sleep -Seconds 1
Stop-Job $job -EA SilentlyContinue; Remove-Job $job -Force -EA SilentlyContinue
