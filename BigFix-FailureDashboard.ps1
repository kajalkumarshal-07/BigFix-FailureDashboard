#Requires -Version 5.1
<#
.SYNOPSIS
    BigFix Failed Device Dashboard - GUI automation for failed BigFix clients.

.DESCRIPTION
    Shows detailed status of failed devices with error details (BigFix fixlet/action
    failures, offline clients, compliance failures, and software deployment
    errors) and maps each error to a root cause + remediation guidance.

    Data sources:
      1) BigFix REST API (live)  - default https://<server>:52311
      2) Import fallback         - CSV or XML reports
      3) Built-in demo data      - works with no server

.PARAMETER SelfTest
    Load demo data, run analysis, print summary, and exit (no GUI). For validation.

.PARAMETER Server
    BigFix REST API base URL. Example: https://bes.corp.local:52311

.PARAMETER ExportCsv
    Headless: export records to CSV then exit (no GUI).

.PARAMETER ExportHtml
    Headless: export executive HTML report then exit (no GUI).

.PARAMETER ExportJson
    Headless: export records to JSON then exit (no GUI).

.PARAMETER Email
    Headless: send SMTP digest (settings.json) after loading records, then exit.

.PARAMETER ImportPath
    Headless: load failures from CSV/XML before export/email.

.PARAMETER UseCache
    Headless/GUI: load last successful API pull from cache instead of live API.

.PARAMETER ApiWorker
    Internal: background worker mode for async Connect API (do not use directly).

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\BigFix-FailureDashboard.ps1

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\BigFix-FailureDashboard.ps1 -SelfTest

.EXAMPLE
    powershell.exe -File .\BigFix-FailureDashboard.ps1 -ExportHtml C:\out\bf.html -ExportCsv C:\out\bf.csv -UseCache
#>
[CmdletBinding()]
param(
    [string]$Server = '',
    [string]$Username = '',
    [string]$Password = $env:BIGFIX_DASH_PASSWORD,
    [switch]$SelfTest,
    [switch]$NoRelaunch,
    [string]$ExportCsv = '',
    [string]$ExportHtml = '',
    [string]$ExportJson = '',
    [switch]$Email,
    [string]$ImportPath = '',
    [switch]$UseCache,
    [string]$ApiWorker = ''
)

# ---------------------------------------------------------------------------
# STA required for WinForms
# ---------------------------------------------------------------------------
if (-not $NoRelaunch -and -not $SelfTest -and -not $ApiWorker -and -not $ExportCsv -and -not $ExportHtml -and -not $ExportJson -and -not $Email -and -not $ImportPath) {
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
        $argsList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-WindowStyle', 'Hidden', '-File', $PSCommandPath)
        if ($Server)     { $argsList += @('-Server', $Server) }
        if ($Username)   { $argsList += @('-Username', $Username) }
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argsList -WindowStyle Hidden | Out-Null
        exit 0
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try {
    [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
    [System.Windows.Forms.Application]::add_ThreadException({
        param($sender, $e)
        $line = 'unknown'
        if ($e.Exception -and $e.Exception.StackTrace -match 'BigFix-FailureDashboard\.ps1: line (\d+)') { $line = $Matches[1] }
        $msg = "UI exception @ line $line : $($e.Exception.Message)"
        try { [System.IO.File]::AppendAllText((Join-Path $env:TEMP 'BigFixDashboard-UIErrors.log'), "$(Get-Date -Format o) $msg`r`n$($e.Exception)`r`n`r`n") } catch { }
    })
} catch { }
[System.Windows.Forms.Application]::EnableVisualStyles()

$ErrorActionPreference = 'Stop'
$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:AppVersion = '2.0.0'
$script:SettingsPath = Join-Path $script:ScriptRoot 'settings.json'
$script:CacheDir = Join-Path $script:ScriptRoot 'cache'
$script:CachePath = Join-Path $script:CacheDir 'last-pull.json'

# ---------------------------------------------------------------------------
# Settings persistence (passwords are NEVER written to disk)
# ---------------------------------------------------------------------------
function Get-DefaultAppSettings {
    [pscustomobject]@{
        Version              = 1
        Server               = 'https://localhost:52311'
        Username             = ''
        OfflineThresholdHours = 4
        SkipTlsCertCheck     = $false
        AutoRefreshMinutes   = 0
        AlertOnNewCritical   = $true
        CategoryFilter       = 'All'
        SeverityFilter       = 'All'
        ExportDirectory      = ''
        DeviceGroupView      = $false
        Smtp                 = [pscustomobject]@{
            Enable = $false; Host = ''; Port = 587; UseSsl = $true
            From = ''; To = ''; Username = ''
        }
        Ticket               = [pscustomobject]@{
            Enable = $false; WebhookUrl = ''; HeaderJson = ''
        }
    }
}

function Get-AppSettings {
    $s = Get-DefaultAppSettings
    if (Test-Path -LiteralPath $script:SettingsPath) {
        try {
            $j = Get-Content -LiteralPath $script:SettingsPath -Raw | ConvertFrom-Json
            foreach ($p in $j.PSObject.Properties) {
                if ($p.Name -eq 'Version') { continue }
                if ($p.Name -in @('Smtp', 'Ticket')) {
                    $target = $s.($p.Name)
                    foreach ($sp in $p.Value.PSObject.Properties) {
                        if ($target.PSObject.Properties[$sp.Name]) { $target.($sp.Name) = $sp.Value }
                    }
                } elseif ($s.PSObject.Properties[$p.Name]) {
                    $s.($p.Name) = $p.Value
                }
            }
        } catch { }
    }
    $script:AppSettings = $s
    return $s
}

function Save-AppSettings {
    param($Settings = $script:AppSettings)
    if (-not $Settings) { return $false }
    try {
        $Settings | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
        return (Test-Path -LiteralPath $script:SettingsPath)
    } catch {
        try { [System.IO.File]::AppendAllText((Join-Path $env:TEMP 'BigFixDashboard-UIErrors.log'), "$(Get-Date -Format o) Save-AppSettings: $($_.Exception.Message)`r`n") } catch { }
        return $false
    }
}

$null = Get-AppSettings

# ---------------------------------------------------------------------------
# TLS / self-signed cert support (BigFix often uses self-signed certs)
# ---------------------------------------------------------------------------
function Set-BigFixTls {
    param([switch]$SkipCertCheck)
    try {
        $tls12 = [Net.SecurityProtocolType]::Tls12
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $tls12
    } catch { }
    if ($SkipCertCheck) {
        if (-not ('BigFixDashTrustAll' -as [type])) {
            Add-Type -TypeDefinition @'
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class BigFixDashTrustAll : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }
}
'@
        }
        [Net.ServicePointManager]::CertificatePolicy = New-Object BigFixDashTrustAll
    }
}

# ---------------------------------------------------------------------------
# Root-cause knowledge base (BigFix + compliance + software patterns)
# ---------------------------------------------------------------------------
$script:RootCauseRules = @(
    @{ Id='BF-DOWNLOAD'; Pattern='failed to download|download (error|failure)|prefetch.*(fail|error)|relay.*(unavailable|failed|unreachable)|failed to (connect|contact) relay'; Category='FixletFailure'; Severity='High'
       RootCause='BigFix client cannot download content from relay/root server.'
       Remediation='Verify BESClient service, relay reachability (443/52311), DNS, _BESData permissions, and relay affinity/list.' }
    @{ Id='BF-DISK'; Pattern='insufficient (disk )?space|no space left|disk full|0x80070070'; Category='FixletFailure'; Severity='High'
       RootCause='Not enough free disk space to complete the BigFix action.'
       Remediation='Free space on system drive; clear Temp and BES cache (_BESData); re-run the action.' }
    @{ Id='BF-ACCESS'; Pattern='access is denied|access denied|unauthorized|logon failure|privilege.*not held'; Category='FixletFailure'; Severity='High'
       RootCause='Permission/authorization failure during BigFix action execution.'
       Remediation='Run as SYSTEM, verify ACLs on target paths, review action script and execution account context.' }
    @{ Id='BF-INSTALL-EXIT'; Pattern='exit code 16[0-9][0-9]|error 3010|error 1603|error 1618|reboot required|restart (required|pending)|restarting is required'; Category='FixletFailure'; Severity='Medium'
       RootCause='Installer returned MSI/reboot error (3010 / 1603 / 1618 or similar).'
       Remediation='Schedule reboot if required; for 1603/1618 check MSI verbose logs, concurrent installer lock, and free disk space.' }
    @{ Id='BF-SERVICE'; Pattern='service .* (could not|failed|unable)|failed to start service|error 1068|error 1053'; Category='FixletFailure'; Severity='Medium'
       RootCause='Windows service failed to start/stop during remediation.'
       Remediation='Check service dependencies, System Event Log, hung services; reboot device and retry action.' }
    @{ Id='BF-RELEVANCE'; Pattern='relevance.*(error|failed)|incorrect plural|fixed substitution|Singular statements'; Category='FixletFailure'; Severity='Medium'
       RootCause='Relevance evaluation error in the BigFix fixlet/task definition.'
       Remediation='Open Fixlet Debugger in BigFix console, correct relevance, re-deploy the fixlet.' }
    @{ Id='BF-LOCK'; Pattern='file (in use|is locked|locked)|another installation|installation in progress|0x80070663|error 1618'; Category='FixletFailure'; Severity='Medium'
       RootCause='Target file locked or another installation already in progress.'
       Remediation='Close blocking processes, wait for other installs to finish, reboot if pending, then retry.' }
    @{ Id='BF-RELAY-OFFLINE'; Pattern='BESClient.*(stopped|not running)|client not reporting|no report|heartbeat.*(fail|miss)'; Category='Offline'; Severity='Critical'
       RootCause='BESClient service stopped or unable to send reports to the relay/root server.'
       Remediation='Structured remote recovery: restart BESClient; inspect daily logs (YYYYMMDD/BESClient.log) for register/HTTP/socket errors; test TCP 52311 to relay; resolve relay DNS; if service+network OK but reports still fail, reset __BESData cache; see Runbook > Clients Not Reporting.' }
    @{ Id='BF-NET-PORT'; Pattern='port 52311|tcp.52311|firewall.*(block|drop)|connection (refused|timed? ?out)|test-netconnection'; Category='Offline'; Severity='High'
       RootCause='Endpoint cannot reach relay/root on TCP 52311 (firewall, routing, or relay down).'
       Remediation='Run remote TCP 52311 test to assigned relay; inspect local/edge firewalls and routing; confirm relay listening; see Runbook > Clients Not Reporting / Relay Connectivity.' }
    @{ Id='BF-DNS-RELAY'; Pattern='dns.*(fail|error|resolution|lookup)|name resolution|resolve.*relay|relay.*(fqdn|hostname).*(fail|unknown)'; Category='Offline'; Severity='High'
       RootCause='Endpoint cannot resolve the relay FQDN from clientsettings/actionsite (DNS or HOSTS issue).'
       Remediation='Run remote Resolve-DnsName for relay; compare to active relay IP; fix DNS/hosts; see Runbook > Clients Not Reporting.' }
    @{ Id='BF-CACHE-CORRUPT'; Pattern='corrupt.*(client|database|cache|sqlite)|__BESData.*(corrupt|reset)|stuck (gather|gatherer)|client database'; Category='Offline'; Severity='Critical'
       RootCause='BESClient internal database/cache (__BESData) corrupted or gather stuck; service may run but reports fail.'
       Remediation='Stop BESClient, delete __BESData, start BESClient (rebuild site defs); verify new log contacts relay; see Runbook > Clients Not Reporting.' }

    @{ Id='COMP-BASELINE'; Pattern='not compliant|non-compliant|compliance (check )?failed|baseline (failed|non-compliant)|failed (audit|configuration)'; Category='ComplianceFailure'; Severity='High'
       RootCause='Configuration/compliance baseline evaluated as failed on the device.'
       Remediation='Open failing baseline properties in BigFix console, compare against expected state, apply remediation fixlet, or document exception.' }
    @{ Id='COMP-HARDENING'; Pattern='security (policy|baseline).*(fail|violation)|hardening.*(fail|drift)|CIS |STIG '; Category='ComplianceFailure'; Severity='High'
       RootCause='Security hardening / benchmark drift detected (CIS/STIG/policy).'
       Remediation='Review benchmark results, apply hardening fixlet, or record a sanctioned exception.' }

    @{ Id='PATCH-FAIL'; Pattern='patch (failed|error)|update (failed|error)|KB[0-9]+.*(fail|error)|windows update failed|CBS.*(error|fail)'; Category='SoftwareDeployment'; Severity='High'
       RootCause='Patch / OS update deployment failed.'
       Remediation='Inspect CBS.log and WindowsUpdate.log, run DISM /Online /Cleanup-Image /RestoreHealth, retry deployment.' }
    @{ Id='APP-DEPLOY'; Pattern='application (deployment|install).*(fail|error)|package (failed|error)|deployment failed|AppDeploy'; Category='SoftwareDeployment'; Severity='High'
       RootCause='Application or package deployment failed.'
       Remediation='Review installer logs, confirm content availability, check prerequisites/permissions, redeploy.' }
    @{ Id='DEFENDER'; Pattern='defender|virus|malscan|threat.*(found|detected)|malware'; Category='SoftwareDeployment'; Severity='High'
       RootCause='Endpoint protection / threat-related failure.'
       Remediation='Run full scan, review quarantine, update definitions, check Microsoft-Windows-Windows Defender operational log.' }

    @{ Id='BF-DL-FAILED'; Pattern='download failed|Download Failed|pending downloads|waiting for downloads'; Category='FixletFailure'; Severity='High'
       RootCause='BigFix action Download Failed - required content download did not complete (relay unreachable or mirror pending).'
       Remediation='Check relay reachability (443/52311), content mirror status, free disk; re-push the action. See Runbook > Action Status.' }
    @{ Id='BF-HASH-MISMATCH'; Pattern='hash mismatch|failed a hash comparison'; Category='FixletFailure'; Severity='High'
       RootCause='Download completed but failed hash comparison - possible network corruption between agent and parent relay.'
       Remediation='Investigate network between agent and parent relay; re-download content; check relay disk/CPU health. See Runbook > Action Status.' }
    @{ Id='BF-ACTION-WAITING'; Pattern='action status.*(waiting|locked|pending)|pending (message|login|restart)|computer is locked|waiting for user'; Category='FixletFailure'; Severity='Medium'
       RootCause='Action is Waiting/Locked/Pending - waiting on user input, login, restart, or time constraint.'
       Remediation='Open Action History, clear stuck Running actions, or stop and re-issue action. See Runbook > Action Status.' }
    @{ Id='BF-NOT-REPORTED'; Pattern='not reported|no report on this action|action.*not reported'; Category='FixletFailure'; Severity='Medium'
       RootCause='Action Not Reported - endpoint never received or processed the action (UDP notification or connectivity issue).'
       Remediation='ForceRefresh the client, check GatherHashMV in BESClient.log, enable command polling if UDP blocked. See Runbook > Slow Clients.' }
    @{ Id='BF-GATHERHASH'; Pattern='GatherHashMV|forcerefresh.*(not|fail|miss)|command received.*(absent|miss)|content.*not.*(received|evaluat)'; Category='FixletFailure'; Severity='Medium'
       RootCause='Client is not receiving new content notifications (UDP/52311 blocked) so actions/analyses evaluate late.'
       Remediation='Enable _BESClient_Comm_CommandPollEnable and persistent connections; verify UDP 52311; ForceRefresh. See Runbook > Slow Clients.' }
    @{ Id='BF-ONE-ACTION'; Pattern='stuck.*running|action.*block|only one action|concurrent action'; Category='FixletFailure'; Severity='Medium'
       RootCause='BigFix client runs one action at a time - a stuck Running action blocks new actions.'
       Remediation='Open computer > Action History > By Action Status; Stop the stuck Running action or restart BESClient. See Runbook > Action Status.' }

    @{ Id='BFI-UPLOAD'; Pattern='MaxArchiveSize|maxarchivesize|scan.*(not|fail).*(upload|import)|force reupload|Computer Support Data|software scan.*(fail|error)|capacity scan.*(fail|error)'; Category='SoftwareDeployment'; Severity='High'
       RootCause='BigFix Inventory scan results finished locally but failed to upload/import (MaxArchiveSize exceeded, blocked upload, or scan failure).'
       Remediation='Use Computer Support Data panel to force upload; run Force Reupload of Software Scan Results; check MaxArchiveSizeExceeded; check scan return codes. See Runbook > Inventory Uploads.' }
    @{ Id='BFI-UUID'; Pattern='duplicate uuid|uuid.*(duplicate|conflict)|VM UUID|capacity calculation'; Category='SoftwareDeployment'; Severity='Medium'
       RootCause='Duplicate virtual machine UUIDs break BigFix Inventory capacity calculation and virtualization hierarchy.'
       Remediation='Resolve duplicate UUIDs (To Do list / import log); for ROK media hosts set SMBIOS.reflectHost=false in .vmx; force capacity scan reupload. See Runbook > Inventory Uploads.' }
    @{ Id='BFI-SIGNATURE'; Pattern='software id tag|signature.*(missing|not found)|bundl(ing|e).*(fail|missing)|not discovered'; Category='SoftwareDeployment'; Severity='Medium'
       RootCause='Installed software not discovered because Software ID tag / signature is missing on the endpoint.'
       Remediation='Update signature catalog (Management > Catalog Upload); re-initiate software scan; check bundling rules. See Runbook > Inventory Uploads.' }

    @{ Id='BF-INSTALL-1920'; Pattern='error 1920|error 1923|service.*failed to start.*(install|setup)|BESClient.*(install|upgrade).*(fail|error)'; Category='FixletFailure'; Severity='High'
       RootCause='BESClient install/upgrade failed - service could not start during setup (dependency, AV injection, permissions).'
       Remediation='Check service logon (Local System), exclude BESClient.exe from DLL injection (e.g. 0patch), review upgrade log, reinstall previous version if needed. See Runbook > Client Install.' }
    @{ Id='BF-WINSOCK'; Pattern='winsock|4294967288|4294967286|client.*register.*(fail|error)|relay registration'; Category='FixletFailure'; Severity='High'
       RootCause='Client cannot register to relay - Winsock corruption or network error preventing registration.'
       Remediation='Run network tests, netsh winsock reset if corrupted, verify 52311 reachability, check BESClient.log relay selection. See Runbook > Client Install.' }

    @{ Id='BF-FILLDB'; Pattern='FillDB|FillDb|filldb.*(log|buffer|stop)|buffer.*back.?up'; Category='FixletFailure'; Severity='High'
       RootCause='FillDB service/log issue - reports reach server but console Last Report Time does not update.'
       Remediation='Stop FillDb service, rename FillDB log, start FillDb; clear console cache; issue blank action test. See Runbook > Console Performance.' }
    @{ Id='BF-CONSOLE-CACHE'; Pattern='console.*(cache|sluggish|slow|freeze)|stale.*(console|dashboard|last report)|visual.*(delay|freeze)'; Category='FixletFailure'; Severity='Medium'
       RootCause='Corrupted/bloated local console cache or heavy DB queries causing stale/sluggish UI.'
       Remediation='Clear %LOCALAPPDATA%\\BigFix cache, restart console, blank action test, trim Stop/Expired actions, check FillDB. See Runbook > Console Performance.' }

    @{ Id='BF-TIMESYNC'; Pattern='w32time|time service|clock skew|time.*(sync|synchroniz)|system time.*(wrong|off)'; Category='FixletFailure'; Severity='Medium'
       RootCause='Windows Time service stopped or clock skew breaks certificate/token evaluation and BigFix action windows.'
       Remediation='Restart W32Time, run w32tm /resync /force, verify NTP source; check domain/GPO time settings.' }
    @{ Id='BF-PROXY'; Pattern='proxy.*(fail|error|config|settings)|WinHttp|default proxy|unexpected proxy'; Category='Offline'; Severity='High'
       RootCause='System/WinHTTP proxy misconfigured so BESClient cannot reach relay/root directly.'
       Remediation='Inspect netsh winhttp show proxy and IE proxy settings; clear bad proxy for BigFix ports; verify WPAD if used.' }
    @{ Id='BF-CERT'; Pattern='certificate.*(expir|invalid|trust)|TLS.*(error|fail|handshake)|SSL.*(error|handshake|alert)'; Category='Offline'; Severity='High'
       RootCause='TLS/certificate trust failure between client and relay/root (expired, untrusted CA, hostname mismatch).'
       Remediation='Check relay/root cert expiry and chain on the endpoint; install/correct CA trust; fix server cert if expired.' }
    @{ Id='BF-FIREWALL'; Pattern='firewall.*(disable|block|drop)|windows defender firewall|profile.*(block)'; Category='FixletFailure'; Severity='Medium'
       RootCause='Local firewall rules/profile blocking BigFix traffic or the action target port.'
       Remediation='Verify firewall state and rules for 52311/443; re-enable Windows Firewall profiles; align with security baseline.' }
    @{ Id='BF-LOW-MEMORY'; Pattern='out of memory|memory pressure|insufficient memory|not enough memory|0x8007000E'; Category='FixletFailure'; Severity='High'
       RootCause='Endpoint ran out of memory during action/scan execution.'
       Remediation='Identify top memory processes, free RAM, reduce concurrent loads, retry the action; check for memory leaks.' }
    @{ Id='BF-PENDING-REBOOT'; Pattern='pending reboot|restart pending.*(install|update)|reboot.*(pending|required) before'; Category='FixletFailure'; Severity='Medium'
       RootCause='A previous install/update left a pending reboot that blocks the current BigFix action.'
       Remediation='Schedule/perform reboot, then re-run the action; verify CBS/WU pending reboot registry keys clear.' }
    @{ Id='BF-GPO'; Pattern='group policy|gpupdate|policy.*(denied|failed)|resultant set of policy'; Category='FixletFailure'; Severity='Medium'
       RootCause='Group Policy application failure or policy conflict affecting the target configuration.'
       Remediation='Run gpupdate /force, review event logs for GPSvc errors, fix GPO scope/filtering; coordinate with AD team.' }
    @{ Id='BF-VPN'; Pattern='vpn.*(disconnect|fail|not connected|tunnel)|always.?on.?vpn|anyconnect|globalprotect'; Category='Offline'; Severity='High'
       RootCause='Endpoint is on VPN/tunnel outage so BigFix relay path is unreachable.'
       Remediation='Confirm VPN session and split-tunnel routes; test DNS+TCP 52311 while on VPN; escalate network if tunnel is up but relay fails.' }
    @{ Id='BF-DISK-IO'; Pattern='disk.*(error|i/o|i/o error|corrupt)|status.*bad|SMART.*(fail|error)|uncorrectable'; Category='FixletFailure'; Severity='Critical'
       RootCause='Disk I/O or SMART health failure preventing reliable action/script execution.'
       Remediation='Back up data, run chkdsk/S.M.A.R.T. diagnostics, replace failing disk; do not force large installs until healthy.' }
    @{ Id='BF-SCCM-CONFLICT'; Pattern='ccmexec|configuration manager|SCCM.*(conflict|lock)|another management agent'; Category='FixletFailure'; Severity='Medium'
       RootCause='Configuration Manager (or peer agent) lock/content conflict with BigFix action execution.'
       Remediation='Check ccmexec activity and local locks, sequence agents, avoid concurrent installs, coordinate deployment windows.' }
    @{ Id='BF-WMI'; Pattern='WMI.*(error|corrupt|fail|repository)|repository.*(corrupt|broken)|winmgmt'; Category='FixletFailure'; Severity='High'
       RootCause='WMI repository corruption or WinMgmt failure breaking inventory/action queries.'
       Remediation='winmgmt /verifyrepository; if broken, /salvagerepository or resetrepository; restart WinMgmt; retest inventory.' }
    @{ Id='BF-ACCOUNT-LOCKED'; Pattern='account.*(lock|disabled)|logon failure.*lock|user name.*bad|525|775'; Category='FixletFailure'; Severity='High'
       RootCause='Account locked/disabled or bad credentials used by the action/service context.'
       Remediation='Unlock/reset account in AD, verify service logon accounts, re-issue action with correct context.' }
    @{ Id='BF-UAC'; Pattern='elevation required|run as administrator|UAC|user account control.*(block|deny)|740'; Category='FixletFailure'; Severity='Medium'
       RootCause='Action needed elevation and was blocked by UAC/policy in the execution context.'
       Remediation='Ensure action runs as SYSTEM/admin, adjust installer elevation, or provide offer/elevate settings in the fixlet.' }
)

# ---------------------------------------------------------------------------
# Remote fix solutions (WinRM / PS remoting) - keyed by MatchedRule Id
# Manual=$true means no safe remote script (console/server action only).
# Script runs on the target via Invoke-Command; may set $global:FixNote.
# ---------------------------------------------------------------------------
$script:FixSolutions = @{
    'BF-DOWNLOAD' = @{ Title='Restart BESClient + clear pending download cache'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service -Name $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out="Service $svc not installed"} }
Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$besData = Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData'
if (Test-Path $besData) {
  Get-ChildItem -Path $besData -Recurse -Directory -Filter '__PendingDownload' -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
}
Start-Service -Name $svc
Start-Sleep -Seconds 2
$st=(Get-Service -Name $svc).Status
return @{Ok=($st -eq 'Running');Out="BESClient=$st; pending download folders cleared"}
'@ }
    'BF-DL-FAILED' = @{ Title='Restart BESClient + requeue failed downloads'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service -Name $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed'} }
Stop-Service $svc -Force -ErrorAction SilentlyContinue
Start-Sleep 1
$roots=@(Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData')
foreach ($root in $roots) {
  if (Test-Path $root) {
    Get-ChildItem $root -Recurse -Directory -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -match 'Pending|Download' } |
      ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
  }
}
Start-Service $svc
Start-Sleep 2
return @{Ok=((Get-Service $svc).Status -eq 'Running');Out='BESClient restarted; download cache cleared'}
'@ }
    'BF-HASH-MISMATCH' = @{ Title='Flush BES cache and restart client (hash mismatch)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed'} }
Stop-Service $svc -Force -ErrorAction SilentlyContinue
Start-Sleep 1
$cache=@(
  (Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData'),
  (Join-Path $env:ProgramData 'BigFix')
)
foreach ($c in $cache) {
  if (Test-Path $c) {
    Get-ChildItem $c -Recurse -Force -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -match 'sha1|hash|__Download' } |
      ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
  }
}
Start-Service $svc
Start-Sleep 2
return @{Ok=((Get-Service $svc).Status -eq 'Running');Out='Cache flushed; BESClient restarted (retry action from console)'}
'@ }
    'BF-RELAY-OFFLINE' = @{ Title='Offline recovery: service, logs, TCP 52311, DNS, cache reset'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$steps=New-Object System.Collections.Generic.List[string]
$svc='BESClient'
$besRoot=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client'
$besData=Join-Path $besRoot '__BESData'
$logDir=Join-Path $besData '__Global\Logs'
if (-not (Get-Service -Name $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed - deploy client'} }
# 1) Service check / restart
$s=Get-Service -Name $svc
if ($s.Status -ne 'Running') {
  Set-Service $svc -StartupType Automatic -ErrorAction SilentlyContinue
  try { Start-Service $svc -ErrorAction Stop; $steps.Add('Svc: started') } catch { $steps.Add("Svc: start failed - $($_.Exception.Message)") }
} else { $steps.Add('Svc: Running') }
Start-Sleep -Seconds 2
$st=(Get-Service $svc).Status
$procs=@(Get-Process BESClient -ErrorAction SilentlyContinue).Count
$steps.Add("Svc=$st procs=$procs")
# 2) Inspect logs (today's YYYYMMDD.log + BESClient.log)
$logText=''
$hits=@()
$pats='Failed to register|public key exchange|HTTP Error (404|502|503)|socket error|Winsock error|ReportSummary|relay.*(reject|unavailable|failed)|Report posted'
if (Test-Path $logDir) {
  $daily=Join-Path $logDir ("{0:yyyyMMdd}.log" -f (Get-Date))
  $files=@()
  if (Test-Path $daily) { $files += $daily }
  $besLog=Join-Path $logDir 'BESClient.log'
  if (Test-Path $besLog) { $files += $besLog }
  $lines=@()
  foreach ($f in $files) { $lines += (Get-Content $f -Tail 60 -ErrorAction SilentlyContinue) }
  $logText=($lines -join ' | ')
  $hits=@($lines | Where-Object { $_ -match $pats } | Select-Object -Last 5)
} else { $steps.Add('Logs: dir missing') }
if ($hits.Count -gt 0) { $steps.Add("LogHits: " + (($hits | ForEach-Object { ($_ -replace '\s+',' ').Trim() }) -join ' // ')) }
elseif ($logText) { $steps.Add('LogHits: none in tail') }
else { $steps.Add('LogHits: no log files') }
# 3+4) Discover relay, test TCP 52311 + DNS
$relay=$null
$cfg=Join-Path $besRoot 'clientsettings.cfg'
if (Test-Path $cfg) {
  $cl=Get-Content $cfg -ErrorAction SilentlyContinue
  foreach ($line in $cl) {
    if ($line -match '^(RelayList1|Master_BESRelay|_BESClientRelay_|_BESClient_RelayList1)=') {
      if ($line -match '=([^:,\s]+)') { $relay=$Matches[1].Trim(); break }
    }
  }
}
if (-not $relay -and $logText) {
  if ($logText -match 'Relay select[^\r\n]*?([A-Za-z0-9][A-Za-z0-9._-]+)') { $relay=$Matches[1] }
}
$portOk=$null; $dnsOk=$null; $relayIp=''
if ($relay) {
  try {
    $dns=Resolve-DnsName -Name $relay -Type A -ErrorAction Stop | Where-Object { $_.IPAddress } | Select-Object -First 1
    if ($dns) { $relayIp=$dns.IPAddress; $dnsOk=$true } else { $dnsOk=$false }
    $steps.Add("DNS: $relay -> $(if ($relayIp) { $relayIp } else { 'no A record' })")
  } catch { $dnsOk=$false; $steps.Add("DNS: FAIL $($_.Exception.Message)") }
  try {
    $tn=Test-NetConnection -ComputerName $relay -Port 52311 -WarningAction SilentlyContinue -ErrorAction Stop
    $portOk=[bool]$tn.TcpTestSucceeded
    $steps.Add("TCP52311: $relay :52311 = $portOk")
  } catch { $portOk=$false; $steps.Add("TCP52311: FAIL $($_.Exception.Message)") }
} else {
  $steps.Add('Relay: not found in clientsettings.cfg/logs (skip port/DNS)')
}
# 5) Full cache reset only if service running AND network path looks OK but reports still broken
$needReset=$false
if ($st -eq 'Running' -and $relay -and $dnsOk -eq $true -and $portOk -eq $true) {
  $badLog=($hits -match 'Failed to register|HTTP Error|socket error|Winsock error|relay.*(reject|unavailable)')
  $noSuccess= -not ($hits -match 'Report posted')
  if ($badLog -or $noSuccess) { $needReset=$true }
}
if ($needReset -and (Test-Path $besData)) {
  Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
  Start-Sleep -Seconds 1
  try {
    Remove-Item -Path $besData -Recurse -Force -ErrorAction Stop
    $steps.Add('Cache: __BESData removed')
  } catch { $steps.Add("Cache: remove failed - $($_.Exception.Message)") }
  Start-Service -Name $svc -ErrorAction SilentlyContinue
  Start-Sleep -Seconds 3
  $st=(Get-Service $svc).Status
  $steps.Add("Svc after reset=$st")
} elseif ($st -eq 'Running' -and -not $needReset) {
  $steps.Add('Cache: reset skipped (service up; no hard fail signals or network incomplete)')
} else {
  $steps.Add('Cache: reset skipped (service not running)')
}
$procs=@(Get-Process BESClient -ErrorAction SilentlyContinue).Count
$ok=($st -eq 'Running' -and $procs -gt 0)
return @{Ok=$ok;Out=($steps -join '; ')}
'@ }
    'BF-NET-PORT' = @{ Title='Test TCP 52311 connectivity to BigFix relay/root'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$besRoot=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client'
$relay=$null
$cfg=Join-Path $besRoot 'clientsettings.cfg'
if (Test-Path $cfg) {
  foreach ($line in (Get-Content $cfg -ErrorAction SilentlyContinue)) {
    if ($line -match '^(RelayList1|Master_BESRelay|_BESClientRelay_|_BESClient_RelayList1)=') {
      if ($line -match '=([^:,\s]+)') { $relay=$Matches[1].Trim(); break }
    }
  }
}
if (-not $relay) {
  $log=Join-Path $besRoot '__BESData\__Global\Logs\BESClient.log'
  if (Test-Path $log) {
    $t=(Get-Content $log -Tail 80 -ErrorAction SilentlyContinue) -join ' '
    if ($t -match 'Relay select[^\r\n]*?([A-Za-z0-9][A-Za-z0-9._-]+)') { $relay=$Matches[1] }
  }
}
if (-not $relay) { return @{Ok=$false;Out='Relay host not found in clientsettings.cfg or BESClient.log'} }
try {
  $tn=Test-NetConnection -ComputerName $relay -Port 52311 -WarningAction SilentlyContinue -ErrorAction Stop
  $ping=if ($tn.PingSucceeded) { 'ping=OK' } else { 'ping=fail' }
  return @{Ok=[bool]$tn.TcpTestSucceeded;Out="Relay=$relay; TCP52311=$($tn.TcpTestSucceeded); $ping; RemoteAddr=$($tn.RemoteAddress)"}
} catch {
  return @{Ok=$false;Out="Relay=$relay; TCP52311 test failed - $($_.Exception.Message)"}
}
'@ }
    'BF-DNS-RELAY' = @{ Title='Resolve relay hostname (DNS) from client settings'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$besRoot=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client'
$relay=$null
$cfg=Join-Path $besRoot 'clientsettings.cfg'
if (Test-Path $cfg) {
  foreach ($line in (Get-Content $cfg -ErrorAction SilentlyContinue)) {
    if ($line -match '^(RelayList1|Master_BESRelay|_BESClientRelay_|_BESClient_RelayList1)=') {
      if ($line -match '=([^:,\s]+)') { $relay=$Matches[1].Trim(); break }
    }
  }
}
if (-not $relay) { return @{Ok=$false;Out='Relay FQDN not found in clientsettings.cfg (RelayList1/Master_BESRelay)'} }
try {
  $a=Resolve-DnsName -Name $relay -Type A -ErrorAction Stop | Where-Object { $_.IPAddress } | Select-Object -First 1
  if ($a) { return @{Ok=$true;Out="DNS OK: $relay -> $($a.IPAddress) (check matches active relay)"} }
  return @{Ok=$false;Out="DNS: $relay returned no A record"}
} catch {
  return @{Ok=$false;Out="DNS FAIL: $relay - $($_.Exception.Message)"}
}
'@ }
    'BF-CACHE-CORRUPT' = @{ Title='Reset client cache: stop BESClient, delete __BESData, restart'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service -Name $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed'} }
Stop-Service -Name $svc -Force -ErrorAction Stop
Start-Sleep -Seconds 1
$besData=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData'
$removed=$false
if (Test-Path $besData) {
  Remove-Item -Path $besData -Recurse -Force -ErrorAction Stop
  $removed=$true
}
Start-Service -Name $svc
Start-Sleep -Seconds 3
$st=(Get-Service $svc).Status
$procs=@(Get-Process BESClient -ErrorAction SilentlyContinue).Count
$log=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData\__Global\Logs\BESClient.log'
$logNote=if (Test-Path $log) { 'new log present' } else { 'log rebuilding' }
return @{Ok=($st -eq 'Running' -and $procs -gt 0);Out="__BESData removed=$removed; BESClient=$st; procs=$procs; $logNote (verify relay contact + register)"}
'@ }
    'BF-DISK' = @{ Title='Free disk space (Temp + BES download folders)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$freed=0
$targets=@(
  (Join-Path $env:TEMP '*'),
  'C:\Windows\Temp\*',
  (Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData\__Global\__BESData\__PendingDownload\*')
)
foreach ($t in $targets) {
  Get-ChildItem -Path $t -Force -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      $len=if ($_.PSIsContainer) { 0 } else { $_.Length }
      Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop
      $freed+=$len
    } catch { }
  }
}
$d=Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':')) -ErrorAction SilentlyContinue
$freeGb=[math]::Round(($d.Free/1GB),2)
return @{Ok=($freeGb -gt 1);Out="Approx freed=$([math]::Round($freed/1MB,1)) MB; free=${freeGb} GB on $($env:SystemDrive)"}
'@ }
    'BF-ACCESS' = @{ Title='Collect access-denied diagnostics (ACLs + event log)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$lines=@()
$lines += "User=$env:USERNAME; Computer=$env:COMPUTERNAME"
Get-WinEvent -FilterHashtable @{LogName='Application';StartTime=(Get-Date).AddHours(-2)} -MaxEvents 15 -ErrorAction SilentlyContinue |
  Where-Object { $_.LevelDisplayName -in @('Error','Warning') } |
  ForEach-Object { $lines += ("{0} {1}" -f $_.TimeCreated, ($_.Message -replace '\s+',' ').Substring(0,[Math]::Min(160,($_.Message -replace '\s+',' ').Length))) }
$besLog=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData\__Global\Logs\BESClient.log'
if (Test-Path $besLog) { $lines += (Get-Content $besLog -Tail 10 -ErrorAction SilentlyContinue) }
return @{Ok=$true;Out=($lines -join ' | ')}
'@ }
    'BF-INSTALL-EXIT' = @{ Title='Check reboot-pending state after MSI failure'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$pending=$false
$keys=@(
 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
)
foreach ($k in $keys) { if (Test-Path $k) { $pending=$true } }
$svr=Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
if ($svr) { $pending=$true }
$msi=Get-Process msiexec -ErrorAction SilentlyContinue
return @{Ok=$true;Out="RebootPending=$pending; msiexecRunning=$([bool]$msi)"}
'@ }
    'BF-SERVICE' = @{ Title='Restart failed Windows service related to action'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$names=@('BESClient','wuauserv','BITS','Winmgmt','ccmexec','Spooler')
$rep=@()
foreach ($n in $names) {
  $s=Get-Service -Name $n -ErrorAction SilentlyContinue
  if ($s -and $s.Status -ne 'Running') {
    try { Start-Service $n -ErrorAction Stop; $rep += "$n=Running" }
    catch { $rep += "$n=Fail($($_.Exception.Message))" }
  } elseif ($s) { $rep += "$n=Running" }
}
if (-not $rep) { $rep += 'No known related services found' }
return @{Ok=$true;Out=($rep -join '; ')}
'@ }
    'BF-RELEVANCE' = @{ Title='Relevance error - fix in BigFix console (manual)'
        Manual=$true
        Script='' }
    'BF-LOCK' = @{ Title='Clear stuck MSI / installation lock'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$msi=Get-Process msiexec -ErrorAction SilentlyContinue
if (-not $msi) { return @{Ok=$true;Out='No msiexec running (lock cleared or never held)'} }
$age=((Get-Date) - ($msi | Sort-Object StartTime | Select-Object -First 1).StartTime).TotalMinutes
if ($age -gt 30) {
  $msi | Stop-Process -Force -ErrorAction SilentlyContinue
  return @{Ok=$true;Out="Stopped stale msiexec (age=$([int]$age) min)"}
}
return @{Ok=$true;Out="mssiexec active age=$([int]$age) min - wait or retry later"}
'@ }
    'BF-ACTION-WAITING' = @{ Title='Restart BESClient to clear stuck Waiting/Locked action'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed'} }
Restart-Service $svc -Force
Start-Sleep 2
return @{Ok=((Get-Service $svc).Status -eq 'Running');Out='BESClient restarted to unblock action queue (re-issue stuck action from console if needed)'}
'@ }
    'BF-ONE-ACTION' = @{ Title='Restart BESClient to clear stuck Running action'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed'} }
Restart-Service $svc -Force
Start-Sleep 2
return @{Ok=((Get-Service $svc).Status -eq 'Running');Out='BESClient restarted (stuck Running action should clear)'}
'@ }
    'BF-NOT-REPORTED' = @{ Title='Enable command polling + restart BESClient (ForceRefresh path)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed'} }
$cfgRoot=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client'
$cfg=Join-Path $cfgRoot 'clientsettings.cfg'
$line='_BESClient_Comm_CommandPollEnable=1'
if (Test-Path $cfg) {
  $c=Get-Content $cfg -ErrorAction SilentlyContinue
  if ($c -notmatch [regex]::Escape($line)) { Add-Content -Path $cfg -Value $line -Encoding ASCII }
} else {
  New-Item -Path $cfg -ItemType File -Force | Out-Null
  Set-Content -Path $cfg -Value $line -Encoding ASCII
}
Restart-Service $svc -Force
Start-Sleep 2
return @{Ok=((Get-Service $svc).Status -eq 'Running');Out='CommandPollEnable=1 written; BESClient restarted'}
'@ }
    'BF-GATHERHASH' = @{ Title='Enable command polling (UDP/52311 notifications blocked)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed'} }
$cfg=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\clientsettings.cfg'
$lines=@('_BESClient_Comm_CommandPollEnable=1','_BESClient_Comm_CommandPollIntervalSeconds=300')
if (Test-Path $cfg) { $existing = @(Get-Content $cfg -ErrorAction SilentlyContinue | ForEach-Object { $_ }) } else { $existing = @() }
foreach ($l in $lines) { if ($existing -notcontains $l) { Add-Content -Path $cfg -Value $l -Encoding ASCII } }
Restart-Service $svc -Force
Start-Sleep 2
return @{Ok=((Get-Service $svc).Status -eq 'Running');Out='Command polling enabled; BESClient restarted'}
'@ }
    'BF-INSTALL-1920' = @{ Title='Fix BESClient service startup (1920/1923)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$svc='BESClient'
$s=Get-Service $svc -ErrorAction SilentlyContinue
if (-not $s) { return @{Ok=$false;Out='BESClient not installed'} }
try {
  $si=Get-CimInstance Win32_Service -Filter "Name='BESClient'" -ErrorAction Stop
  if ($si.StartName -ne 'LocalSystem' -and $si.StartName -ne '.\LocalSystem') {
    # leave account change to admin - report only
  }
  if ($s.Status -ne 'Running') { Start-Service $svc -ErrorAction Stop }
  Start-Sleep 2
  $st=(Get-Service $svc).Status
  return @{Ok=($st -eq 'Running');Out="BESClient=$st; StartName=$($si.StartName)"}
} catch {
  return @{Ok=$false;Out="Start failed: $($_.Exception.Message)"}
}
'@ }
    'BF-WINSOCK' = @{ Title='Reset Winsock + restart BESClient (relay registration)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$reset=$false
try {
  $p=Start-Process -FilePath 'netsh' -ArgumentList 'winsock','reset' -Wait -PassThru -WindowStyle Hidden
  $reset=($p.ExitCode -eq 0)
} catch { }
$svc='BESClient'
if (Get-Service $svc -ErrorAction SilentlyContinue) { Restart-Service $svc -Force -ErrorAction SilentlyContinue }
return @{Ok=$true;Out="WinsockReset=$reset (reboot may be required); BESClient restarted if present"}
'@ }
    'COMP-BASELINE' = @{ Title='Compliance failure - remediate from BigFix console (manual)'
        Manual=$true
        Script='' }
    'COMP-HARDENING' = @{ Title='Hardening drift - apply hardening fixlet (manual)'
        Manual=$true
        Script='' }
    'PATCH-FAIL' = @{ Title='Repair Windows Update stack (DISM + reset WU)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
try {
  $dism=Start-Process dism.exe -ArgumentList '/Online','/Cleanup-Image','/RestoreHealth' -Wait -PassThru -WindowStyle Hidden
  $rep += "DISM=$($dism.ExitCode)"
} catch { $rep += "DISM=Fail($($_.Exception.Message))" }
foreach ($n in 'wuauserv','BITS','cryptsvc') {
  if (Get-Service $n -ErrorAction SilentlyContinue) {
    try { Restart-Service $n -Force -ErrorAction Stop; $rep += "$n=Restarted" } catch { $rep += "$n=Fail" }
  }
}
$wuLog=Join-Path $env:SystemRoot 'SoftwareDistribution\Download'
if (Test-Path $wuLog) {
  Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
  Get-ChildItem $wuLog -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
  Start-Service wuauserv -ErrorAction SilentlyContinue
  $rep += 'WUCache=Cleared'
}
return @{Ok=$true;Out=($rep -join '; ')}
'@ }
    'APP-DEPLOY' = @{ Title='Application deploy diagnostics (space + installer processes)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$d=Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':')) -ErrorAction SilentlyContinue
$free=[math]::Round(($d.Free/1GB),2)
$proc=@(Get-Process msiexec,setup,install -ErrorAction SilentlyContinue | ForEach-Object { $_.Name }) -join ','
$rep=@("FreeGB=$free","Installers=$proc")
if ($free -lt 2) {
  Get-ChildItem $env:TEMP -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
  $d=Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':')) -ErrorAction SilentlyContinue
  $rep += "TempCleared; FreeGB=$([math]::Round(($d.Free/1GB),2))"
}
return @{Ok=($free -ge 1);Out=($rep -join '; ')}
'@ }
    'DEFENDER' = @{ Title='Microsoft Defender: update definitions + quick scan'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
try {
  $mp=Get-MpComputerStatus -ErrorAction Stop
  $rep += "AMProduct=$($mp.AMProductVersion); SigAgeDays=$([int]($mp.AntivirusSignatureLastUpdated - (Get-Date)).TotalDays * -1)"
  try { Update-MpSignature -ErrorAction Stop; $rep += 'SigUpdate=OK' } catch { $rep += "SigUpdate=$($_.Exception.Message)" }
} catch {
  return @{Ok=$false;Out="Defender WMI unavailable: $($_.Exception.Message)"}
}
return @{Ok=$true;Out=($rep -join '; ')}
'@ }
    'BFI-UPLOAD' = @{ Title='Restart BESClient to requeue Inventory scan upload'
        Manual=$false
        Script=@'
$ErrorActionPreference='Stop'
$svc='BESClient'
if (-not (Get-Service $svc -ErrorAction SilentlyContinue)) { return @{Ok=$false;Out='BESClient not installed'} }
$scan=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData\__Global'
if (Test-Path $scan) {
  Get-ChildItem $scan -Recurse -Filter '*.lock' -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
}
Restart-Service $svc -Force
Start-Sleep 2
return @{Ok=((Get-Service $svc).Status -eq 'Running');Out='BESClient restarted; scan locks cleared (force reupload from Inventory console if still failing)'}
'@ }
    'BFI-UUID' = @{ Title='Duplicate UUID - fix in VM manager / Inventory (manual)'
        Manual=$true
        Script='' }
    'BFI-SIGNATURE' = @{ Title='Software ID tag missing - catalog upload in Inventory (manual)'
        Manual=$true
        Script='' }
    'BF-FILLDB' = @{ Title='FillDB / console stale - server-side action (manual)'
        Manual=$true
        Script='' }
    'BF-CONSOLE-CACHE' = @{ Title='Clear local BigFix console cache (run on console host, manual)'
        Manual=$true
        Script='' }
    'BF-ACCESS-DEFAULT' = @{ Title='Collect BESClient.log + daily YYYYMMDD.log failure tail'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$lines=@()
$logDir=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData\__Global\Logs'
$pats='Failed to register|public key exchange|HTTP Error (404|502|503)|socket error|Winsock error|ReportSummary|relay.*(reject|unavailable|failed)|Report posted'
$hits=@()
if (Test-Path $logDir) {
  $daily=Join-Path $logDir ("{0:yyyyMMdd}.log" -f (Get-Date))
  if (Test-Path $daily) { $lines += Get-Content $daily -Tail 30 -ErrorAction SilentlyContinue }
  $besLog=Join-Path $logDir 'BESClient.log'
  if (Test-Path $besLog) { $lines += Get-Content $besLog -Tail 30 -ErrorAction SilentlyContinue }
}
$hits=@($lines | Where-Object { $_ -match $pats } | Select-Object -Last 5)
$svc=Get-Service BESClient -ErrorAction SilentlyContinue
$hitTxt=if ($hits) { (($hits | ForEach-Object { ($_ -replace '\s+',' ').Trim() }) -join ' // ') } else { 'none' }
return @{Ok=($svc -and $svc.Status -eq 'Running');Out="BESClient=$($svc.Status); logLines=$($lines.Count); failureHits=$hitTxt"}
'@ }
    'BF-TIMESYNC' = @{ Title='Restart W32Time and force time resync'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
try {
  if (Get-Service W32Time -ErrorAction SilentlyContinue) {
    Set-Service W32Time -StartupType Automatic -ErrorAction SilentlyContinue
    Restart-Service W32Time -Force -ErrorAction Stop
    $rep += 'W32Time=Restarted'
  } else { $rep += 'W32Time=Missing' }
} catch { $rep += "W32Time=Fail($($_.Exception.Message))" }
try {
  $p = Start-Process w32tm.exe -ArgumentList '/resync','/force' -Wait -PassThru -WindowStyle Hidden
  $rep += "ResyncExit=$($p.ExitCode)"
} catch { $rep += "Resync=Fail($($_.Exception.Message))" }
try {
  $st = w32tm /query /status 2>$null | Out-String
  if ($st -match 'Source:\s*(.+)') { $rep += "Source=$($Matches[1].Trim())" }
  if ($st -match 'Last Successful Sync Time:\s*(.+)') { $rep += "LastSync=$($Matches[1].Trim())" }
} catch { }
$now = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$rep += "LocalNow=$now"
return @{Ok=$true;Out=($rep -join '; ')}
'@ }
    'BF-PROXY' = @{ Title='Inspect WinHTTP/system proxy settings'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
try {
  $wh = (netsh winhttp show proxy 2>&1 | Out-String).Trim() -replace '\s+',' '
  $rep += "WinHTTP: $wh"
} catch { $rep += "WinHTTP: Fail($($_.Exception.Message))" }
try {
  $ie = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue
  $rep += "ProxyEnable=$($ie.ProxyEnable); ProxyServer=$($ie.ProxyServer); AutoConfigURL=$($ie.AutoConfigURL)"
} catch { }
return @{Ok=$true;Out=($rep -join ' | ')}
'@ }
    'BF-CERT' = @{ Title='Inspect TLS trust to relay/root (cert store + connectivity)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$besRoot=Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client'
$relay=$null
$cfg=Join-Path $besRoot 'clientsettings.cfg'
if (Test-Path $cfg) {
  foreach ($line in (Get-Content $cfg -ErrorAction SilentlyContinue)) {
    if ($line -match '^(RelayList1|Master_BESRelay|_BESClientRelay_|_BESClient_RelayList1)=') {
      if ($line -match '=([^:,\s]+)') { $relay=$Matches[1].Trim(); break }
    }
  }
}
if (-not $relay) { return @{Ok=$true;Out='Relay host unknown - cannot probe TLS; check clientsettings.cfg'} }
try {
  $tcp = Test-NetConnection -ComputerName $relay -Port 52311 -WarningAction SilentlyContinue -ErrorAction Stop
  $note = "TCP52311=$($tcp.TcpTestSucceeded)"
} catch { $note = "TCP52311=Fail($($_.Exception.Message))" }
try {
  $req = [System.Net.HttpWebRequest]::Create("https://${relay}:52311/")
  $req.Method = 'GET'
  $req.Timeout = 8000
  $req.AllowAutoRedirect = $false
  try { $resp = $req.GetResponse(); $resp.Close(); $note += '; TLS=OK' }
  catch [System.Net.WebException] {
    $r = $_.Exception.Response
    if ($r) { $note += "; TLS=HTTP$([int]$r.StatusCode)" }
    elseif ($_.Exception.Status -match 'Trust|Certificate|SecureChannel') { $note += "; TLS=TrustFail($($_.Exception.Status))" }
    else { $note += "; TLS=Fail($($_.Exception.Message))" }
  }
} catch { $note += "; TLS=Error($($_.Exception.Message))" }
return @{Ok=$true;Out="Relay=$relay; $note"}
'@ }
    'BF-FIREWALL' = @{ Title='Report Windows Firewall profiles and BigFix-related rules'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
try {
  $prof = Get-NetFirewallProfile -ErrorAction Stop | ForEach-Object { "$($_.Name)=$($_.Enabled)" }
  $rep += ('Profiles: ' + ($prof -join ', '))
} catch { $rep += "Profiles: Fail($($_.Exception.Message))" }
try {
  $rules = Get-NetFirewallRule -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match 'BigFix|BES|52311' } |
    Select-Object -First 8 -ExpandProperty DisplayName
  $rep += 'BigFixRules: ' + $(if ($rules) { ($rules -join ', ') } else { 'none matched' })
} catch { }
return @{Ok=$true;Out=($rep -join ' | ')}
'@ }
    'BF-LOW-MEMORY' = @{ Title='Report top memory processes and free RAM'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
$freeGb = if ($os) { [math]::Round($os.FreePhysicalMemory/1MB,2) } else { -1 }
$top = Get-Process -ErrorAction SilentlyContinue |
  Sort-Object WorkingSet64 -Descending |
  Select-Object -First 8 |
  ForEach-Object { '{0}={1}MB' -f $_.Name, [int]($_.WorkingSet64/1MB) }
return @{Ok=($freeGb -ge 0);Out="FreeGB=$freeGb; Top: $($top -join ', ')"}
'@ }
    'BF-PENDING-REBOOT' = @{ Title='Detect pending reboot registry markers'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$pending=@()
$keys=@(
 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations'
)
foreach ($k in $keys) { if (Test-Path $k) { $pending += $k } }
$svr=Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
if ($svr) { $pending += 'PendingFileRenameOperations' }
return @{Ok=$true;Out=$(if ($pending) { 'PENDING REBOOT: ' + ($pending -join '; ') } else { 'No pending reboot markers' })}
'@ }
    'BF-GPO' = @{ Title='Force Group Policy refresh and read GPSvc errors'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
try {
  $p = Start-Process gpupdate.exe -ArgumentList '/force' -Wait -PassThru -WindowStyle Hidden
  $rep += "gpupdate=$($p.ExitCode)"
} catch { $rep += "gpupdate=Fail($($_.Exception.Message))" }
try {
  $ev = Get-WinEvent -FilterHashtable @{LogName='System';Id=1058,1030,1125,1127;StartTime=(Get-Date).AddDays(-2)} -MaxEvents 5 -ErrorAction SilentlyContinue
  $rep += 'GPEvents=' + $(if ($ev) { (($ev | ForEach-Object { $_.Id }) -join ',') } else { 'none' })
} catch { }
return @{Ok=$true;Out=($rep -join '; ')}
'@ }
    'BF-VPN' = @{ Title='Report network adapters and default route (VPN path check)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
try {
  $ad = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object Status -eq 'Up' |
    ForEach-Object { "$($_.Name):$($_.InterfaceDescription)" }
  $rep += 'Up: ' + $(if ($ad) { ($ad -join ', ') } else { 'none' })
} catch { $rep += "Adapters=Fail($($_.Exception.Message))" }
try {
  $routes = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
    Sort-Object RouteMetric | Select-Object -First 3 |
    ForEach-Object { "$($_.NextHop)(metric $($_.RouteMetric))" }
  $rep += 'DefaultRoutes: ' + $(if ($routes) { ($routes -join ', ') } else { 'none' })
} catch { }
return @{Ok=$true;Out=($rep -join ' | ')}
'@ }
    'BF-DISK-IO' = @{ Title='Report logical disk health indicators (free space + chkdsk flags)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
try {
  $d = Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':')) -ErrorAction Stop
  $rep += "FreeGB=$([math]::Round($d.Free/1GB,2))"
} catch { }
try {
  $vol = Get-CimInstance Win32_Volume -Filter "DriveLetter='$env:SystemDrive'" -ErrorAction SilentlyContinue
  if ($vol) { $rep += "HealthStatus=$($vol.HealthStatus); ErrorMethodInfo=$($vol.ErrorMethodInfo)" }
} catch { }
try {
  $bad = Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='disk','Ntfs','Application Popup';StartTime=(Get-Date).AddDays(-7)} -MaxEvents 5 -ErrorAction SilentlyContinue
  $rep += 'DiskEvents=' + $(if ($bad) { $bad.Count } else { 0 })
} catch { }
return @{Ok=$true;Out=($rep -join '; ')}
'@ }
    'BF-SCCM-CONFLICT' = @{ Title='Report CCM/exec agent activity (management conflict check)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
$rep=@()
$svc=Get-Service ccmexec -ErrorAction SilentlyContinue
if ($svc) { $rep += "ccmexec=$($svc.Status)" } else { $rep += 'ccmexec=NotInstalled' }
$msi=Get-Process msiexec -ErrorAction SilentlyContinue
$rep += "msiexecCount=$(@($msi).Count)"
$locks = @()
$besLock = Join-Path ${env:ProgramFiles(x86)} 'BigFix Enterprise\BES Client\__BESData'
if (Test-Path $besLock) {
  $locks = @(Get-ChildItem $besLock -Recurse -Filter '*.lock' -ErrorAction SilentlyContinue | Select-Object -First 5 -ExpandProperty Name)
}
$rep += 'BESLocks: ' + $(if ($locks) { ($locks -join ',') } else { 'none' })
return @{Ok=$true;Out=($rep -join '; ')}
'@ }
    'BF-WMI' = @{ Title='Verify WMI repository (winmgmt /verifyrepository)'
        Manual=$false
        Script=@'
$ErrorActionPreference='Continue'
try {
  $p = Start-Process winmgmt.exe -ArgumentList '/verifyrepository' -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "$env:TEMP\wmiver.txt"
  $out = ''
  if (Test-Path "$env:TEMP\wmiver.txt") { $out = (Get-Content "$env:TEMP\wmiver.txt" -Raw -ErrorAction SilentlyContinue) }
  $msg = if ($out) { ($out -replace '\s+',' ').Trim() } else { "exit=$($p.ExitCode)" }
  $ok = ($msg -match 'consistent' -and $msg -notmatch 'inconsistent')
  return @{Ok=$ok;Out="verifyrepository: $msg"}
} catch {
  return @{Ok=$false;Out="verifyrepository failed: $($_.Exception.Message)"}
}
'@ }
    'BF-ACCOUNT-LOCKED' = @{ Title='Account lockout - manual identity remediation (collect context only)'
        Manual=$true
        Script='' }
    'BF-UAC' = @{ Title='Elevation/UAC blocked action - adjust fixlet context (manual)'
        Manual=$true
        Script='' }
}

$script:FixCategoryDefaults = @{
    'Offline'           = 'BF-RELAY-OFFLINE'
    'FixletFailure'     = 'BF-ACCESS-DEFAULT'
    'ComplianceFailure' = 'COMP-BASELINE'
    'SoftwareDeployment'= 'APP-DEPLOY'
}

# ---------------------------------------------------------------------------
# Day-to-day runbook (HCL BigFix Maintenance & Troubleshooting + community KBs)
# ---------------------------------------------------------------------------
$script:RunbookTopics = @(
    @{ Id='clients-offline'; Title='1. Clients Not Reporting / Offline Endpoints'
       Tags='Offline,BESClient,52311,helper,forcerefresh,not reported'
       Summary='Endpoints show as old / <not reported> / not checking in within the expected timeframe. Last Report Time is not updating in the Console.'
       RootCause='Blocked network ports (TCP/UDP 52311), stopped/hung BES Client service, local firewall or security product blocking traffic, automatic relay selection failure, or UDP command notifications not reaching the client.'
       Steps=@(
           'Verify TCP/UDP port 52311 is open end-to-end across firewalls between endpoints, relays, and the root server (test with telnet/nc from the endpoint).'
           'Check the BES Client service is running on the endpoint; restart it. On Windows install the BES Client Helper Service (Fixlet #591 Install / #592 Uninstall) so a watcher auto-restarts a hung client (helper tries: restart client -> remove revocation file (backup) -> remove _BESData folder -> restart again).'
           'Inspect client log BESClient.log for "Report posted successfully" and "Relay select" messages to confirm communication flow and relay selection.'
           '  Windows logs: C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs'
           '  Linux/UNIX  : /var/opt/BESClient/__BESData/__Global/Logs'
           '  macOS       : /Library/Application Support/BigFix/BES Agent/__BESData/__Global/Logs'
           'Confirm device is powered on, on the network, DNS resolves the relay/root, and relay list/affinity is correct (clientsettings.cfg or automatic relay selection).'
           'Force a client refresh: right-click the computer in BigFix Console -> Refresh, or send ForceRefresh: curl "http://<server>:52311/cgi-bin/bfenterprise/clientregister.exe?RequestType=NotifyClient&Body=ForceRefresh&ComputerID=<id>".'
           'If console shows stale data but logs show reports posting: clear Console cache, issue a blank custom action, and check FillDB (stop FillDb service, rename FillDB log, restart FillDb).'
           'Enable client debug logging (Fixlet #157 BES Client Setting: Enable Debug Logging) and usage profiler (Fixlet #361) to trace relay selection and slow evaluations.'
           'If UDP notifications are blocked, enable Command Polling / persistent connections: _BESClient_Comm_CommandPollEnable=1 and _BESClient_Comm_CommandPollIntervalSeconds.'
           'In this dashboard: filter Category=Offline, sort by HoursSinceReport, apply the remediation on each row.'
       )
       Refs=@(
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_client_helper.html'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_maintenance_and_troubleshootin.html'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/r_client_set.html'
           'https://forum.bigfix.com/t/tip-troubleshooting-client-reponsiveness/35306'
           'https://forum.bigfix.com/t/computers-stuck-in-not-reported-have-to-manually-restart-bes-service/47497'
           'https://forum.bigfix.com/t/how-to-ensure-client-receives-forcerefresh/47752'
           'https://support.hcl-software.com/csm?id=kb_article&sysparm_article=KB0022437'
       ) }
    @{ Id='relay'; Title='2. Relay Connectivity / Communication Lags'
       Tags='Relay,lag,curl,diagnostics,certificates,hierarchy'
       Summary='Relays fail to gather content or forward client reports upstream; communication lags or a relay chain is stuck holding back reports.'
       RootCause='Network interruptions, expired certificates, incorrect port/firewall blocks, overloaded relay, or a stuck relay in the parent chain.'
       Steps=@(
           'Connectivity check from a terminal (use curl on the port - not SSH):'
           '  curl -k https://<relay-address>:52311/cgi-bin/bfenterprise/clientregister.exe?RequestType=Version'
           '  (expect HTTPS response on port 52311)'
           'Check relay parent chains; inspect relay logs for stuck evaluation loops or buffer bottlenecks.'
           'BigFix Management > Analyses > activate relay Status analysis; watch Distance to relay for topology jumps / loops.'
           'Open Dashboards > BES Support > Relay Health Dashboard: Endpoints per Relay (imbalance), Hierarchy (relay loops), Inactive (stale relays / stopped service highlighted red), Version difference.'
           'Minimize clients reporting directly to the server - prefer healthy relays (Deployment Health Checks dashboard).'
           'Run BigFix Diagnostics Tool and Relay/Server diagnostics when problems persist.'
           'Enable debug/verbose logging on Root Server and Relay services while diagnosing (then turn off).'
           'For DMZ relays confirm persistent TCP connection parent<->child settings; check certificate expiry on authenticating relays.'
           'If a specific relay is version-skewed, force site gather on that relay and confirm it catches up to root site versions.'
       )
       Refs=@(
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Installation/c_monitoring_relay_health.html'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Installation/c_relay_diagnostics.html'
           'https://help.hcl-software.com/bigfix/9.5/platform/Platform/Console/c_relay_health_dashboard.html'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_maintenance_and_troubleshootin.html'
       ) }
    @{ Id='console-slow'; Title='3. Console Performance Delays / Visual Freezes'
       Tags='Console,cache,blank action,FillDB,audit trail,sluggish'
       Summary='BigFix administration console lags, freezes, or displays outdated dashboard / Last Report Time content while backend looks healthy.'
       RootCause='Corrupted or bloated local client-side console cache, heavy database queries, stopped/expired action history bloat, or FillDB buffer/log backup.'
       Steps=@(
           'Clear or completely delete the local Console cache and restart the console (close Console first).'
           '  Typical cache locations: %LOCALAPPDATA%\BigFix, console install folder cache/profile dirs.'
           'Issue a blank custom action to test basic deployment propagation end-to-end (watch for notification in client log).'
           'If Last Report Time is stale but client logs show "Report posted successfully": stop FillDb service, rename the FillDB log, start FillDb service again.'
           'Trim action history: delete Stop/Expired actions older than ~1-3 months (or run Audit Trail Cleaner) - large action history makes the console sluggish; export metadata first if you need audit records.'
           'Review expensive relevances on Web Reports if Web UI is slow.'
           'If server-side: check FillDB buffer back-up, disk fragmentation, SQL patching, and BigFix Management domain Fixlets.'
           'Confirm console machine has adequate RAM/CPU and LAN-speed connection to the BigFix server.'
       )
       Refs=@(
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_maintenance_and_troubleshootin.html'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_monitoringexpensiverelevances.html'
           'https://forum.bigfix.com/t/bigfix-endpoints-showing-offline/34543'
           'https://forum.bigfix.com/t/stopped-and-expired-actions/40975'
       ) }
    @{ Id='inventory-upload'; Title='4. BigFix Inventory Scans Not Uploading'
       Tags='Inventory,scan,upload,BFI,UUID,MaxArchiveSize,Computer Support Data'
       Summary='Software/capacity scans finish on the endpoint but results do not populate in BigFix Inventory; upload failures or missing software.'
       RootCause='Missing software signatures / Software ID tags, interrupted scan tasks, blocked upload results, MaxArchiveSize exceeded on the server, or duplicate VM UUIDs disrupting capacity calculation.'
       Steps=@(
           'Open the computer in BigFix Inventory -> Computer Support Data panel: download the full troubleshooting log package, or force an upload of scan files without opening the main BigFix console.'
           'Force reupload: run "Force Reupload of Software Scan Results" task, or "Run Capacity Scan and Upload Results" -> "run a single capacity scan and force upload of results" from BigFix console.'
           'Check for duplicate VM Universally Unique Identifiers (UUIDs) in the To Do list / import log - duplicates break capacity calculation and virtualization hierarchy (set SMBIOS.reflectHost="false" in .vmx for ROK media hosts).'
           'Check MaxArchiveSizeExceeded on the endpoint/server - when MaxArchiveSize is exceeded, scan results (and VM Manager data) cannot upload to the BigFix server.'
           'Verify scan fixlets/tasks completed (not stuck/failing) and signature catalogs are current (Management > Catalog Upload).'
           'Check software scan return codes (e.g. 255 = expected file not found, 9009 = copy/move failed, Upload Failed = check MaxArchiveSizeExceeded).'
           'Check free disk and relay upload path for bottlenecks; retry the scan task.'
           'If software still missing: signature/Software ID tag may not exist on the endpoint - check bundling; re-initiate scanner or uninstall/reinstall scanner on trouble hosts.'
           'Verify the endpoint still reports to BigFix (Runbook item 1) before chasing Inventory-only issues.'
       )
       Refs=@(
           'https://help.hcl-software.com/bigfix/10.0/inventory/Inventory/probdet/t_collecting_logs.html'
           'https://help.hcl-software.com/bigfix/11.0/inventory/Inventory/probdet/data_import_warnings.html'
           'https://help.hcl-software.com/bigfix/10.0/inventory/Inventory/probdet/r_probs_vmmanagers.html'
           'https://help.hcl-software.com/bigfix/10.0/inventory/Inventory/probdet/r_scanner_return_codes.html'
           'https://help.hcl-software.com/bigfix/10.0/inventory/Inventory/probdet/r_import_problems.html'
           'https://www.ibm.com/support/pages/vm-managers-cannot-be-uploaded-bigfix-server-maxarchivesize-exceeded'
           'https://forum.bigfix.com/t/removed-software-showing-in-reports-scan-issues/32970'
       ) }
    @{ Id='db-webui'; Title='5. Database Connection Failures (WebUI / Server)'
       Tags='WebUI,SQL,database,sa,credentials,port 8080'
       Summary='WebUI or root components throw credential/login exceptions or disconnect from the backend database; WebUI fails to initialize.'
       RootCause='Modified database passwords, expired SA credentials, stopped SQL services, or account used for routine password rotation.'
       Steps=@(
           'Log into the WebUI/console configuration utility; update hostname and database credentials (sa or service account).'
           'If the DB account password changed after install: run "Deploy/Update WebUI Database Configuration" Fixlet - also repairs credential-based initialization failure.'
           'Restart the BES WebUI service (Windows Services or Task Manager) after applying changes.'
           'Confirm SQL Server service is running; verify network/ports to the DB host; check SQL login/password matches what BigFix/WebUI expect; rotate consistently if passwords changed.'
           'Port conflicts: Web Reports defaulted to 80 before 9.2.5 and 8080 after - WebUI vs Web Reports port conflict can break WebUI apps (Query, Profile, Patch Policies, Send Notification).'
           'WebUI must reach Web Reports via REST API for some apps - if Web Reports is down/remote, restart WebUI server after Web Reports is restored.'
           'Review WebUI/server logs for the exact login exception before changing config.'
       )
       Refs=@(
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_maintenance_and_troubleshootin.html'
           'https://help.hcl-software.com/bigfix/11.0/platform/pdf/BigFix_WebUI_Admin_Guide.pdf'
           'https://forum.bigfix.com/t/webui-is-very-slow-to-start-any-suggestions/43753'
       ) }
    @{ Id='server-maint'; Title='6. Server Maintenance & Troubleshooting Checklist'
       Tags='Maintenance,SQL,backup,defrag,diagnostics,monitoring,FillDB'
       Summary='Routine HCL-recommended care for the BigFix root server, database, and relays.'
       RootCause='Neglected SQL maintenance, disk pressure, ignored Management-domain health signals, or untrimmed action history.'
       Steps=@(
           'Keep SQL Server patched (Patches for Windows site) and familiar with MS SQL Server Tools.'
           'Back up the BigFix database on a regular schedule; run occasional DBCC error-check to validate data.'
           'On performance degradation: check fragmentation; BigFix writes many temp files - defrag when needed; occasional disk error-check.'
           'Run the BigFix Diagnostics Tool whenever server components misbehave.'
           'Check the BigFix Management domain often - Fixlets detect component problems before they hit the network.'
           'Add/maintain relays for performance; healthy relays = healthy deployments.'
           'Review Deployment Health Checks dashboard for optimizations and failures.'
           'Monitor servers for: power/unavailable, disk failure, event-log errors, service states, FillDB buffer directory back-up.'
           'Regularly clean Stop/Expired actions (Audit Trail Cleaner) to keep console/DB responsive.'
           'Enable debug/verbose logging on Root Server/Relay only while diagnosing (then turn off).'
       )
       Refs=@(
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_maintenance_and_troubleshootin.html'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Installation/c_running_the_bigfix_dia.html'
       ) }
    @{ Id='action-status'; Title='7. Action Status Meanings & Stuck Actions'
       Tags='Action,status,failed,download failed,hash mismatch,waiting,locked,forcerefresh'
       Summary='Action shows Failed / Download Failed / Waiting / Locked / Hash Mismatch / Not Reported - understand what each status means and how to clear stuck actions.'
       RootCause='Failed = action ran but issue still relevant (for patches, patch ran but did not fix); Download Failed = required download failed; Hash Mismatch = download completed but failed hash compare (investigate network agent<->parent); Waiting = user input / retry / time window / login; Locked = computer locked; Not Reported = no report yet; Pending Downloads = waiting for mirror; Constrained/Expired = relevance or expiration false.'
       Steps=@(
           'Open the computer -> Action History tab -> Actions > By Property > By Action Status; look for actions stuck in Running - only one action runs at a time, a stuck Running action blocks new ones.'
           'To clear a stuck Running action: Stop the action from the console, or restart the BES Client service (note: restart does not stop an already-launched executable).'
           'Download Failed / Pending Downloads: check relay reachability (443/52311), content mirror status, and free space; re-push the action.'
           'Hash Mismatch: investigate network between agent and its parent relay to eliminate corruption; re-download content.'
           'Waiting on user input / Pending Message: either wait for user or stop and re-issue with appropriate offer/message settings.'
           'Force a client refresh after clearing: Refresh from console, or ForceRefresh URL (see Runbook item 1).'
           'Not Reported: endpoint never received/processed the action - check client log for GatherHashMV / ForceRefresh command received (see Runbook item 8).'
           'Expired/Constrained: relevance became false or expiration passed - re-evaluate fixlet relevance before redeploying.'
           'Bulk cleanup: delete Stop/Expired actions older than retention window (Audit Trail Cleaner) to reduce console load.'
       )
       Refs=@(
           'https://developer.bigfix.com/action-script/guide/action_statuses.html'
           'https://forum.bigfix.com/t/tip-troubleshooting-client-reponsiveness/35306'
           'https://support.hcl-software.com/csm?id=kb_article&sysparm_article=KB0023207'
           'https://developer.bigfix.com/rest-api/api/action.html'
       ) }
    @{ Id='client-slow'; Title='8. Slow / Unresponsive Clients & Content Not Received'
       Tags='slow client,GatherHashMV,ForceRefresh,command poll,persistent,heartbeat,debug'
       Summary='Client Last Report Time updates but actions/analyses evaluate late, or client does not process new content for a long time.'
       RootCause='UDP command notifications (new content) not reaching the client, another action stuck Running, expensive relevance, or client busy with long evaluations.'
       Steps=@(
           'Check client logs for "GatherHashMV command received" or "ForceRefresh command received" - absence means UDP/52311 notifications are blocked.'
           'If UDP is blocked: enable Command Polling (_BESClient_Comm_CommandPollEnable=1, _BESClient_Comm_CommandPollIntervalSeconds) and/or persistent TCP connections.'
           'Confirm only one action is not stuck Running (Action History - see Runbook item 7).'
           'Look for long-running site evaluations in the log; very long site evaluation times often come from expensive custom relevances.'
           'Enable debug logging (Fixlet #157) and BES Client Usage Profiler (Fixlet #361) to find what is consuming client time.'
           'Check heartbeat interval settings if reports are infrequent but present.'
           'If reports post but console does not update: console cache / FillDB issue (Runbook item 3).'
           'For macOS: logs at /Library/Application Support/BigFix/BES Agent/__BESData/__Global/Logs/'
       )
       Refs=@(
           'https://forum.bigfix.com/t/tip-troubleshooting-client-reponsiveness/35306'
           'https://forum.bigfix.com/t/computers-stuck-in-not-reported-have-to-manually-restart-bes-service/47497'
           'https://forum.bigfix.com/t/how-to-ensure-client-receives-forcerefresh/47752'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/r_client_set.html'
           'https://bigfix-wiki.hcltechsw.com/wikis/home/wiki/BigFix%20Wiki/page/Slow%20and%20Unresponsive%20Clients%20-%20Troubleshooting?lang=en-us'
       ) }
    @{ Id='client-install'; Title='9. Client Install / Upgrade Failures (1920 / 1923 / Winsock)'
       Tags='install,upgrade,1920,1923,Winsock,registration,error'
       Summary='BESClient install or upgrade fails with service errors (1920/1923), Winsock errors, or clients fail to register to relay.'
       RootCause='Service cannot start during install (dependency, AV injection such as 0patch, permissions), stale client settings, or network/Winsock issues preventing relay registration.'
       Steps=@(
           'Error 1920/1923 (service failed to start during install): confirm BES Client service logon is Local System; check for security product DLL injection (e.g. 0patch) - exclude BESClient.exe or uninstall conflicting agent; review upgrade log.'
           'Winsock errors 4294967288 / 4294967286 on registration: run network tests, reset Winsock (netsh winsock reset) if corrupted, verify 52311 reachability.'
           'If upgrade fails, uninstall the new build and reinstall previous known-good version while support investigates.'
           'After install, confirm client registers: BESClient.log shows relay selection + Report posted successfully.'
           'For relay assignment at install: clientsettings.cfg in client folder (RelayList1=... or _BESClientRelay_... settings).'
           'Preserve KeyStorage when upgrading if you must reinstall - avoid full client reset unless necessary.'
       )
       Refs=@(
           'https://forum.bigfix.com/t/installation-besclient-11-0-6-137-fails-because-of-error-1920-or-1923/54302'
           'https://forum.bigfix.com/t/agent-11-0-6-upgrade-http-404-sync-error-stuck-pending-restart-how-to-fix-without-resetting-keystorage/54302'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_maintenance_and_troubleshootin.html'
           'https://help.hcl-software.com/bigfix/11.0/platform/Platform/Installation/c_windows_clients.html'
       ) }
    @{ Id='webreports-slow'; Title='10. Web Reports / WebUI Performance & Expensive Relevance'
       Tags='Web Reports,WebUI,expensive relevance,slow,dashboard'
       Summary='Web Reports pages load slowly, dashboards timeout, or WebUI takes very long to start.'
       RootCause='Expensive custom relevances exceeding evaluation thresholds, WebUI recursive file-permission startup (older versions), or remote Web Reports unreachable.'
       Steps=@(
           'Monitor expensive relevances: enable relevance evaluation time monitoring and review relevances exceeding the custom threshold.'
           'Simplify or index heavy dashboards/reports; avoid unbounded plural relevance over all sites.'
           'WebUI slow start (older versions): upgrade - later versions fixed recursive permission setting at startup that caused large delays with many subscribed sites.'
           'Confirm Web Reports is running and reachable via REST API from WebUI server; restart WebUI after Web Reports is restored.'
           'Check port conflicts (Web Reports 8080 vs WebUI) after upgrades from pre-9.2.5.'
           'Review application server memory; Web Reports primary resource need is memory for fast access.'
       )
       Refs=@(
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Config/c_monitoringexpensiverelevances.html'
           'https://forum.bigfix.com/t/webui-is-very-slow-to-start-any-suggestions/43753'
           'https://help.hcl-software.com/bigfix/11.0/platform/pdf/BigFix_WebUI_Admin_Guide.pdf'
           'https://help.hcl-software.com/bigfix/10.0/platform/Platform/Console/c_web_reports.html'
       ) }
)

# ---------------------------------------------------------------------------
# Analysis helpers
# ---------------------------------------------------------------------------
function Get-SeverityRank {
    param([string]$Severity)
    switch -Regex ($Severity) {
        '^Critical' { 4 }
        '^High'     { 3 }
        '^Medium'   { 2 }
        default     { 1 }
    }
}

function Get-NormalizedCategory {
    param([string]$Category)
    $c = "$Category".Trim()
    if (-not $c) { return 'FixletFailure' }
    $map = @{
        'offline'='Offline'; 'notreporting'='Offline'; 'stale'='Offline'
        'fixlet'='FixletFailure'; 'action'='FixletFailure'; 'bigfix'='FixletFailure'
        'compliance'='ComplianceFailure'; 'baseline'='ComplianceFailure'
        'softwaredeployment'='SoftwareDeployment'; 'patch'='SoftwareDeployment'; 'deployment'='SoftwareDeployment'
        'healthy'='Healthy'; 'ok'='Healthy'; 'online'='Healthy'; 'good'='Healthy'
    }
    $key = ($c -replace '[^a-zA-Z]', '').ToLower()
    if ($map.ContainsKey($key)) { return $map[$key] }
    switch -Regex ($c) {
        'offline|not.?report' { return 'Offline' }
        'complian|baseline'   { return 'ComplianceFailure' }
        'patch|deploy|update|software' { return 'SoftwareDeployment' }
        'healthy|^ok$|online' { return 'Healthy' }
        default { return 'FixletFailure' }
    }
}

function Get-FailureAnalysis {
    <# Returns object with Category, Severity, RootCause, Remediation, MatchedRule #>
    param(
        [string]$ErrorMessage = '',
        [string]$ErrorCode = '',
        [string]$CategoryHint = '',
        [string]$Component = ''
    )
    $text = ("$ErrorMessage $ErrorCode $Component").Trim()
    $hintCat = if ($CategoryHint) { Get-NormalizedCategory -Category $CategoryHint } else { '' }

    $matched = @($script:RootCauseRules | Where-Object { $text -match $_.Pattern })
    if ($matched.Count -gt 0) {
        $rule = $matched[0]
        if ($hintCat) {
            $preferred = @($matched | Where-Object { $_.Category -eq $hintCat })
            if ($preferred.Count -gt 0) { $rule = $preferred[0] }
        }
        $finalCat = if ($hintCat) { $hintCat } else { $rule.Category }
        return [pscustomobject]@{
            Category   = $finalCat
            Severity   = $rule.Severity
            RootCause  = $rule.RootCause
            Remediation= $rule.Remediation
            MatchedRule= $rule.Id
        }
    }

    # Category-based defaults when no pattern matched
    $category = if ($hintCat) { $hintCat } else { 'FixletFailure' }
    $sev = 'Medium'
    $root = 'Unclassified failure. Review the raw error, BigFix action details, and Windows Event Log on the device.'
    $rem  = '1) Open the action/fixlet result for this device in BigFix console. 2) Collect BESClient.log and _BESData logs. 3) Check Application/System event logs. 4) Re-run the action after clearing the cause.'
    switch ($category) {
        'Healthy' {
            $sev = 'Low'
            $root = 'Device is reporting to BigFix within the offline threshold. No failed action/fixlet/compliance signals matched for this device.'
            $rem  = 'No action required for BigFix connectivity. Use Category filters (Offline / FixletFailure / ComplianceFailure) to isolate problems.'
        }
        'Offline' {
            $sev = 'Critical'
            $root = 'Device is not reporting to BigFix within the configured threshold.'
            $rem  = 'Verify power/network, BESClient service, DNS to relay/root, firewall ports, and relay list configuration.'
        }
        'ComplianceFailure' {
            $root = 'Compliance/baseline failure without a specific signature.'
            $rem  = 'Open the failing baseline in BigFix console, inspect each property result, apply remediation or waive with justification.'
        }
        'SoftwareDeployment' {
            $root = 'Software/patch deployment failure without a specific signature.'
            $rem  = 'Review installer logs on device, confirm content and prerequisites, then redeploy.'
        }
    }
    if ($ErrorCode -and $ErrorCode -match '^0x[0-9A-Fa-f]+$') {
        $root = "$root (Error code $ErrorCode)"
    }
    return [pscustomobject]@{
        Category   = $category
        Severity   = $sev
        RootCause  = $root
        Remediation= $rem
        MatchedRule= 'Default'
    }
}

function New-FailureRecord {
    param(
        [Parameter(Mandatory)][string]$DeviceName,
        [string]$IPAddress = '',
        [string]$LastReportTime = '',
        [string]$Category = '',
        [string]$Severity = '',
        [string]$Source = 'Unknown',
        [string]$Component = '',
        [string]$ErrorCode = '',
        [string]$ErrorMessage = '',
        [string]$Status = 'Failed'
    )
    $analysis = Get-FailureAnalysis -ErrorMessage $ErrorMessage -ErrorCode $ErrorCode -CategoryHint $Category -Component $Component
    $finalCat = $analysis.Category
    $sev = if ($Severity) { $Severity } else { $analysis.Severity }

    $hours = ''
    if ($LastReportTime) {
        $dt = [datetime]::MinValue
        if ([datetime]::TryParse($LastReportTime, [ref]$dt) -and $dt -ne [datetime]::MinValue) {
            $hours = [math]::Round(((Get-Date) - $dt).TotalHours, 1)
        }
    }

    [pscustomobject]@{
        DeviceName      = $DeviceName
        IPAddress       = $IPAddress
        LastReportTime  = $LastReportTime
        HoursSinceReport= $hours
        Category        = $finalCat
        Severity        = $sev
        SeverityRank    = Get-SeverityRank $sev
        Source          = $Source
        Component       = $Component
        ErrorCode       = $ErrorCode
        ErrorMessage    = $ErrorMessage
        Status          = $Status
        RootCause       = $analysis.RootCause
        Remediation     = $analysis.Remediation
        MatchedRule     = $analysis.MatchedRule
    }
}

# ---------------------------------------------------------------------------
# Demo data (works with no BigFix server)
# ---------------------------------------------------------------------------
function Get-DemoCsv {
    @'
DeviceName,IPAddress,LastReportTime,Category,Severity,Source,Component,ErrorCode,ErrorMessage,Status
FIN-PC-014,10.20.4.114,2026-09-22 06:10:00,FixletFailure,High,Demo,BigFix Patch Deployment,,Failed to download prefetch block - relay unavailable,Failed
HR-LT-207,10.20.7.207,2026-09-22 08:45:00,Offline,Critical,Demo,BESClient,,Client not reporting - last report more than threshold ago,Failed
ENG-SRV-03,10.20.12.3,2026-09-21 18:02:00,FixletFailure,High,Demo,BigFix Content Download,,Download Failed - content hash mismatch from relay,Failed
ENG-SRV-03,10.20.12.3,2026-09-21 18:02:00,FixletFailure,High,Demo,7-Zip Install,,Access is denied when writing to C:\Program Files\7-Zip,Failed
OPS-LT-119,10.20.9.119,2026-09-22 07:30:00,Offline,Critical,Demo,BESClient,,Client not reporting - BESClient service stopped,Failed
MFG-WS-441,10.20.15.41,2026-09-22 05:15:00,ComplianceFailure,High,Demo,Baseline: Windows 11 CIS L1,,Not compliant - Secure Boot disabled and BitLocker off,Failed
MFG-WS-441,10.20.15.41,2026-09-22 05:15:00,SoftwareDeployment,High,Demo,CrowdStrike Sensor,,Application deployment failed - exit code 1603 (insufficient disk space),Failed
LAB-VM-009,10.20.30.9,2026-09-22 09:01:00,FixletFailure,Medium,Demo,BigFix Baseline: TLS Hardening,,Relevance evaluation error - singular statement expected,Failed
LAB-VM-009,10.20.30.9,2026-09-22 09:01:00,SoftwareDeployment,High,Demo,BigFix Inventory Scanner,,MaxArchiveSize exceeded - scan results not uploaded,Failed
EXEC-LT-002,10.20.2.50,2026-09-19 11:20:00,Offline,Critical,Demo,BESClient,,Client not reporting - device offline or BESClient stopped,Failed
SALES-LT-088,10.20.6.88,2026-09-22 04:40:00,FixletFailure,High,Demo,BigFix Action,,Action status: Download Failed - pending downloads stuck,Failed
SALES-LT-088,10.20.6.88,2026-09-22 04:40:00,SoftwareDeployment,Medium,Demo,Zoom Client,,The system cannot find the file specified. (0x80070002),Failed
IT-ADMIN-01,10.20.1.10,2026-09-22 08:58:00,ComplianceFailure,High,Demo,Baseline: Password Policy,,Compliance check failed - minimum password length below policy,Failed
'@
}

function Get-DemoFailures {
    $rows = Get-DemoCsv | ConvertFrom-Csv
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) {
        $list.Add((New-FailureRecord `
            -DeviceName $r.DeviceName `
            -IPAddress $r.IPAddress `
            -LastReportTime $r.LastReportTime `
            -Category $r.Category `
            -Severity $r.Severity `
            -Source $r.Source `
            -Component $r.Component `
            -ErrorCode $r.ErrorCode `
            -ErrorMessage $r.ErrorMessage `
            -Status $r.Status))
    }
    return $list
}

# ---------------------------------------------------------------------------
# Import: CSV / XML (fallback when API unavailable)
# ---------------------------------------------------------------------------
function Resolve-ColumnMap {
    param($headers)
    $map = @{}
    foreach ($h in $headers) {
        $n = ("$h" -replace '[^a-zA-Z0-9]', '').ToLower()
        switch -Regex ($n) {
            '^(devicename|computername|computer|device|hostname|host|name)$' { $map.DeviceName = $h; break }
            '^(ip|ipaddress|ipv4)$' { $map.IPAddress = $h; break }
            '^(lastreporttime|lastreport|lastheartbeat|lastseen|lastcontact)$' { $map.LastReportTime = $h; break }
            '^(category|failuretype|type|failurecategory)$' { $map.Category = $h; break }
            '^(severity|priority|risk)$' { $map.Severity = $h; break }
            '^(source|datasource|origin)$' { $map.Source = $h; break }
            '^(component|fixlet|fixletname|action|actionname|software|application|task)$' { $map.Component = $h; break }
            '^(errorcode|code|hresult|resultcode|exitcode)$' { $map.ErrorCode = $h; break }
            '^(errormessage|error|message|details|description|statusdetails)$' { $map.ErrorMessage = $h; break }
            '^(status|state|result)$' { $map.Status = $h; break }
        }
    }
    return $map
}

function Import-FailureCsv {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $rows = @(Import-Csv -LiteralPath $Path)
    if (-not $rows -or $rows.Count -eq 0) { throw "CSV has no data rows: $Path" }

    $headers = $rows[0].PSObject.Properties.Name
    $map = Resolve-ColumnMap -headers $headers
    if (-not $map.DeviceName) {
        throw "CSV must include a device column (DeviceName/Computer/Name). Found: $($headers -join ', ')"
    }

    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) {
        $get = {
            param($key)
            if ($map[$key]) { [string]$r.($map[$key]) } else { '' }
        }
        $device = & $get 'DeviceName'
        if (-not $device) { continue }
        $list.Add((New-FailureRecord `
            -DeviceName $device `
            -IPAddress (& $get 'IPAddress') `
            -LastReportTime (& $get 'LastReportTime') `
            -Category (& $get 'Category') `
            -Severity (& $get 'Severity') `
            -Source $(if ($map.Source) { (& $get 'Source') } else { "Import: $(Split-Path $Path -Leaf)" }) `
            -Component (& $get 'Component') `
            -ErrorCode (& $get 'ErrorCode') `
            -ErrorMessage (& $get 'ErrorMessage') `
            -Status $(if ($map.Status) { (& $get 'Status') } else { 'Failed' })))
    }
    if ($list.Count -eq 0) { throw "No usable rows parsed from $Path" }
    return $list
}

function Import-FailureXml {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    [xml]$xml = Get-Content -LiteralPath $Path -Raw
    $list = New-Object System.Collections.Generic.List[object]

    # Format A: <Failures><Failure>...</Failure></Failures>
    $nodes = $xml.SelectNodes('//Failure')
    foreach ($n in $nodes) {
        $val = {
            param($tag)
            $e = $n.SelectSingleNode($tag)
            if ($e) { [string]$e.InnerText } else { '' }
        }
        $device = & $val 'DeviceName'
        if (-not $device) { $device = & $val 'Name' }
        if (-not $device) { continue }
        $list.Add((New-FailureRecord `
            -DeviceName $device `
            -IPAddress (& $val 'IPAddress') `
            -LastReportTime (& $val 'LastReportTime') `
            -Category (& $val 'Category') `
            -Severity (& $val 'Severity') `
            -Source 'Import XML' `
            -Component (& $val 'Component') `
            -ErrorCode (& $val 'ErrorCode') `
            -ErrorMessage (& $val 'ErrorMessage') `
            -Status $(if (& $val 'Status') { (& $val 'Status') } else { 'Failed' })))
    }
    if ($list.Count -gt 0) { return $list }

    # Format B: BES <Client> / <Computer> with last report time (offline detection)
    $clients = $xml.SelectNodes('//Client|//Computer|//BESClient')
    $threshold = $script:OfflineThresholdHours
    foreach ($c in $clients) {
        $name = ''
        $id = ''
        $ip = ''
        $last = ''
        foreach ($child in $c.ChildNodes) {
            $n = $child.Name
            if ($n -match '^(Name|ComputerName|NetBIOSName)$') { $name = $child.InnerText }
            elseif ($n -eq 'ID') { $id = $child.InnerText }
            elseif ($n -match '^(IP|IPAddress)$') { $ip = $child.InnerText }
            elseif ($n -match 'LastReport|LastHeartbeat|LastContact') { $last = $child.InnerText }
            elseif ($child.Name -eq 'Property' -and $child.GetAttribute('Name') -match 'Last Report') { $last = $child.InnerText }
        }
        if (-not $name -and $id) { $name = "ID:$id" }
        if (-not $name) { continue }
        if ($last) {
            $dt = [datetime]::MinValue
            if ([datetime]::TryParse($last, [ref]$dt) -and $dt -ne [datetime]::MinValue) {
                $hours = ((Get-Date) - $dt).TotalHours
                if ($hours -ge $threshold) {
                    $list.Add((New-FailureRecord `
                        -DeviceName $name -IPAddress $ip -LastReportTime $last `
                        -Category 'Offline' -Source 'Import XML' `
                        -Component 'BESClient' `
                        -ErrorMessage "Client not reporting - last report $([math]::Round($hours,1)) hours ago (threshold ${threshold}h)" `
                        -Status 'Offline'))
                }
            }
        }
    }
    if ($list.Count -eq 0) {
        throw "No Failure nodes and no offline BES clients found in XML: $Path"
    }
    return $list
}

# ---------------------------------------------------------------------------
# BigFix REST API
# ---------------------------------------------------------------------------
$script:BigFixCredential = $null
$script:BigFixSessionToken = $null
$script:BigFixUsername = ''
$script:BigFixPassword = ''
$script:BigFixAuthMode = ''

function Get-HttpStatusCode {
    param($ErrorRecord)
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($resp -and $resp.StatusCode) { return [int]$resp.StatusCode }
    } catch { }
    if ("$($ErrorRecord.Exception.Message)" -match '\b(401|403|404|500|502|503)\b') {
        return [int]$Matches[1]
    }
    return 0
}

function Format-BigFixApiError {
    param($ErrorRecord, [string]$Uri)
    $status = Get-HttpStatusCode -ErrorRecord $ErrorRecord
    $raw = $ErrorRecord.Exception.Message
    $www = ''
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($resp) {
            try { $www = [string]$resp.Headers['WWW-Authenticate'] } catch { }
        }
    } catch { }

    if ($status -eq 401 -or $raw -match '401|Unauthorized') {
        $lines = @(
            'HTTP 401 Unauthorized from BigFix REST API',
            "URL: $Uri",
            ''
            'How to fix:',
            ' 1. Credentials must be a BigFix Console operator (not only Windows SSO).',
            ' 2. That operator must have permission: Can use REST API',
            '    Console > Site > Master > Allow REST API / operator properties.',
            ' 3. Local BigFix user  -> Username without domain (e.g. apiuser).',
            ' 4. LDAP/AD operator   -> try DOMAIN\user  OR  user@domain.com',
            '    (use the exact login name configured in BigFix).',
            ' 5. SAML/SSO-only accounts often get 401 on REST API.',
            '    Create or use a local BigFix API/service account.',
                ' 6. Quick browser test (enter same user/pass when prompted):',
                "    $((($Uri -split '/api')[0]) + '/api/help')",
                ' 6b. Operator needs BOTH Can use REST API and Custom Content = Yes.',
            ' 7. Prefer FQDN in Server if short name fails.',
            ' 8. Re-type password carefully (special characters matter).',
            ' 9. Confirm the account can log into the BigFix Console.',
            '10. Use the exact same Server URL that worked in the browser',
            '    (https://hostname:52311 - prefer FQDN, no extra path).',
            ''
            "Server said: $raw"
        )
        if ($www) { $lines += "WWW-Authenticate: $www" }
        return ($lines -join [Environment]::NewLine)
    }
    if ($status -eq 403) {
        return "HTTP 403 Forbidden - authenticated but not allowed for this API call.`r`nURL: $Uri`r`nCheck operator site/object permissions.`r`n$raw"
    }
    if ($status -eq 0 -and $raw -match 'SSL|certificate|TLS') {
        return "TLS/certificate error (not auth). Enable 'Skip TLS cert check' or install the BigFix cert.`r`n$raw"
    }
    if ($status -gt 0) {
        return "BigFix API HTTP $status`r`nURL: $Uri`r`n$raw"
    }
    return "Cannot reach BigFix API at $Uri`r`n$raw"
}

function New-BigFixAuth {
    param([string]$Username, [string]$Password, [switch]$KeepSession)
    # KeepSession: phase-2 failure scan reuses sticky auth mode + SessionToken
    if (-not $KeepSession) {
        $script:BigFixSessionToken = $null
        $script:BigFixAuthMode = ''
    }
    if (-not $Username) { return @{} }

    $sec = ConvertTo-SecureString -String $Password -AsPlainText -Force
    $script:BigFixCredential = New-Object System.Management.Automation.PSCredential($Username, $sec)
    $script:BigFixUsername = $Username
    $script:BigFixPassword = $Password

    # Standard BigFix REST auth = HTTP Basic with console operator credentials
    $pair = "${Username}:${Password}"
    return @{
        Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
    }
}

function ConvertFrom-BigFixResponse {
    param($WebResponse)
    $stream = $WebResponse.GetResponseStream()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        $text = $reader.ReadToEnd()
    } finally {
        if ($reader) { $reader.Close() }
        $WebResponse.Close()
    }

    try {
        $token = $WebResponse.Headers['SessionToken']
        if ($token) { $script:BigFixSessionToken = [string]$token }
    } catch { }

    if ([string]::IsNullOrWhiteSpace($text)) { return '' }

    $ctype = ''
    try { $ctype = [string]$WebResponse.ContentType } catch { }
    if ($ctype -match 'json' -or ($text.TrimStart().StartsWith('{') -or $text.TrimStart().StartsWith('['))) {
        try { return ($text | ConvertFrom-Json) } catch { return $text }
    }
    if ($ctype -match 'xml' -or $text.TrimStart().StartsWith('<')) {
        try { return [xml]$text } catch { return $text }
    }
    return $text
}

function Invoke-BigFixWebRequest {
    param(
        [string]$Uri,
        [int]$TimeoutSec,
        [ValidateSet('Credential', 'BasicUtf8', 'BasicLatin1')]
        [string]$Mode = 'Credential'
    )
    # Sanitize URI (no CR/LF/spaces that can corrupt the HTTP request line)
    $Uri = ("$Uri" -replace '[\r\n]+', '').Trim()
    if ($Uri -match '\s') { throw "Invalid URL (contains spaces): $Uri" }

    # BigFix REST is GET-oriented; always use an explicit GET verb (never OPTIONS/HEAD/POST).
    $req = [System.Net.HttpWebRequest]::Create($Uri)
    $req.Method = 'GET'
    $req.Timeout = $TimeoutSec * 1000
    $req.ReadWriteTimeout = $TimeoutSec * 1000
    $req.AllowAutoRedirect = $false
    $req.UserAgent = "BigFix-FailureDashboard/$script:AppVersion"
    $req.Accept = 'application/json, application/xml, text/xml, */*'
    $req.KeepAlive = $true
    $req.ProtocolVersion = [System.Version]'1.1'

    if ($script:BigFixSessionToken) {
        $req.Headers['SessionToken'] = $script:BigFixSessionToken
    }

    switch ($Mode) {
        'Credential' {
            if ($script:BigFixCredential) {
                $req.PreAuthenticate = $true
                $req.Credentials = $script:BigFixCredential
                # Also set explicit Basic so servers that ignore Credential still get auth
                $pair = "${script:BigFixUsername}:${script:BigFixPassword}"
                $req.Headers['Authorization'] = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
            }
        }
        'BasicUtf8' {
            if ($script:BigFixUsername) {
                $pair = "${script:BigFixUsername}:${script:BigFixPassword}"
                $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
                $req.Headers['Authorization'] = "Basic $b64"
            }
        }
        'BasicLatin1' {
            if ($script:BigFixUsername) {
                $pair = "${script:BigFixUsername}:${script:BigFixPassword}"
                $enc = [System.Text.Encoding]::GetEncoding('ISO-8859-1')
                $b64 = [Convert]::ToBase64String($enc.GetBytes($pair))
                $req.Headers['Authorization'] = "Basic $b64"
            }
        }
    }

    try {
        $resp = [System.Net.HttpWebResponse]$req.GetResponse()
    } catch [System.Net.WebException] {
        $status = 0
        $body = ''
        $www = ''
        $token = $null
        $resp = $null
        try { $resp = $_.Exception.Response } catch { }
        if ($resp) {
            try { $status = [int]$resp.StatusCode } catch { }
            try { $token = $resp.Headers['SessionToken'] } catch { }
            try { $www = [string]$resp.Headers['WWW-Authenticate'] } catch { }
            try {
                $s = $resp.GetResponseStream()
                if ($s) {
                    $sr = New-Object System.IO.StreamReader($s, [Text.Encoding]::UTF8)
                    $body = $sr.ReadToEnd()
                    $sr.Close()
                }
            } catch { }
            try { $resp.Close() } catch { }
        }
        if ($token) { $script:BigFixSessionToken = [string]$token }

        $msg = $_.Exception.Message
        if ($status -gt 0) { $msg = "HTTP $status $msg" }
        if ($www) { $msg = "$msg | WWW-Authenticate: $www" }
        if ($body) { $msg = "$msg | $body" }

        $wrapper = New-Object System.Exception($msg, $_.Exception)
        if ($status -gt 0) {
            $fakeResp = [pscustomobject]@{ StatusCode = [System.Net.HttpStatusCode]$status; Headers = $null }
            $wrapper | Add-Member -NotePropertyName Response -NotePropertyValue $fakeResp -Force
        }
        throw $wrapper
    }

    try {
        $token = $resp.Headers['SessionToken']
        if ($token) { $script:BigFixSessionToken = [string]$token }
    } catch { }

    $text = ''
    try {
        $s = $resp.GetResponseStream()
        if ($s) {
            $sr = New-Object System.IO.StreamReader($s, [Text.Encoding]::UTF8)
            $text = $sr.ReadToEnd()
            $sr.Close()
        }
    } catch { }
    $ctype = ''
    try { $ctype = [string]$resp.ContentType } catch { }
    try { $resp.Close() } catch { }

    if ([string]::IsNullOrWhiteSpace($text)) { return '' }

    if ($ctype -match 'json' -or $text.TrimStart().StartsWith('{') -or $text.TrimStart().StartsWith('[')) {
        try { return ($text | ConvertFrom-Json) } catch { return $text }
    }
    if ($ctype -match 'xml' -or $text.TrimStart().StartsWith('<')) {
        try { return [xml]$text } catch { return $text }
    }
    return $text
}

function Get-BigFixBaseUri {
    param([string]$Url)
    $u = "$Url".Trim().TrimEnd('/')
    if (-not $u) { return '' }
    if ($u -notmatch '^https?://') { $u = 'https://' + $u }

    # Keep only scheme://host[:port] — strips /api/help, /api, etc.
    try {
        $uri = [System.Uri]$u
        if (-not $uri.Host) { throw 'Invalid URL' }
        $defaultPort = if ($uri.Scheme -eq 'https') { 443 } else { 80 }
        if ($uri.Port -eq $defaultPort) {
            return ('{0}://{1}' -f $uri.Scheme, $uri.Host)
        }
        return ('{0}://{1}:{2}' -f $uri.Scheme, $uri.Host, $uri.Port)
    } catch {
        if ($u -match '^(https?://[^/]+)') { return $Matches[1] }
        return $u
    }
}

function Invoke-BigFixRest {
    param(
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$Path,
        [int]$TimeoutSec = 30
    )
    $base = Get-BigFixBaseUri -Url $BaseUrl
    if (-not $base) { throw 'Server URL is empty.' }
    if (-not $Path.StartsWith('/')) { $Path = '/' + $Path }
    $uri = $base + $Path

    # Auth: prefer mode that already worked; fall back to full sequence on 401
    $modes = @('Credential', 'BasicUtf8', 'BasicLatin1')
    if ($script:BigFixAuthMode -and $modes -contains $script:BigFixAuthMode) {
        $modes = @($script:BigFixAuthMode) + @($modes | Where-Object { $_ -ne $script:BigFixAuthMode })
    }
    $lastError = $null

    foreach ($mode in $modes) {
        if (($mode -ne 'Credential') -and -not $script:BigFixUsername) { continue }
        try {
            $result = Invoke-BigFixWebRequest -Uri $uri -TimeoutSec $TimeoutSec -Mode $mode
            $script:BigFixAuthMode = $mode
            return $result
        } catch {
            $lastError = $_
            $status = 0
            try {
                if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            } catch { }
            if ($status -eq 401 -and $script:BigFixAuthMode -eq $mode) {
                # known-good mode failed - clear so next call re-probes
                $script:BigFixAuthMode = ''
            }
            if ($status -eq 401) { continue }   # try next auth style
            break                                 # non-auth failure: stop
        }
    }

    if ($lastError) {
        $msg = ''
        try { $msg = $lastError.Exception.Message } catch { $msg = "$lastError" }
        $status = 0
        try {
            if ($lastError.Exception.Response) { $status = [int]$lastError.Exception.Response.StatusCode }
        } catch { }
        if (-not $status -and $msg -match '\b(401|403|404|400|500|502|503)\b') { $status = [int]$Matches[1] }

        if ($status -eq 401 -or $msg -match '401|Unauthorized') {
            $lines = @(
                'HTTP 401 Unauthorized from BigFix REST API',
                "URL: $uri",
                'Tried: browser-style login, Basic UTF-8, Basic Latin-1.',
                '',
                'Browser /api/help works -> account is valid. Check:',
                ' 1. Same Server URL as browser (https://host:52311).',
                ' 2. Same username format as console login.',
                ' 3. Operator has "Can use REST API".',
                ' 4. If browser used Windows SSO without typing a password,',
                '    create a local BigFix API user for this tool.',
                '',
                "Server said: $msg"
            )
            throw ($lines -join [Environment]::NewLine)
        }
        if ($status -eq 400 -or $msg -match 'unknown request method|Bad Request') {
            $lines = @(
                'BigFix API rejected the request (400 / unknown request method).',
                "URL: $uri",
                '',
                'This tool only sends HTTP GET. "unknown request method" means',
                'something sent POST/PATCH/OPTIONS/HEAD/TRACE (BigFix returns 501).',
                '',
                'Fix:',
                ' - In browser/Postman/curl use GET (not POST) for /api/help',
                ' - Server field: https://bigfix:52311  (path /api/* is added by the tool)',
                ' - Root https://bigfix:52311/ returning 404 is normal',
                ' - Browser login page is https://bigfix:52311/api/help',
                '',
                "Server said: $msg"
            )
            throw ($lines -join [Environment]::NewLine)
        }
        if ($status -eq 501 -or $msg -match 'Not Implemented|unknown request method') {
            $lines = @(
                'HTTP 501 Not Implemented / unknown request method',
                "URL: $uri",
                '',
                'BigFix only allows GET (and some POST for creates).',
                'The tool uses GET - if you see this, a proxy/client changed the verb.',
                'Retry Connect API; if it persists, test:',
                '  GET https://bigfix:52311/api/help',
                '',
                "Server said: $msg"
            )
            throw ($lines -join [Environment]::NewLine)
        }
        if ($status -eq 404) {
            $lines = @(
                'HTTP 404 Not Found',
                "URL: $uri",
                '',
                'Checks:',
                ' - Root https://bigfix:52311/ always 404s (normal) - do not use as Server',
                ' - Server should be: https://bigfix:52311',
                ' - API path used: /api/help or /api/computers (added automatically)',
                ' - If /api/help works in browser, the API path should too after login',
                '',
                "Server said: $msg"
            )
            throw ($lines -join [Environment]::NewLine)
        }
        if ($status -eq 403) {
            throw "HTTP 403 Forbidden - logged in but not allowed.`r`nURL: $uri`r`nCheck site/object permissions for this operator.`r`n$msg"
        }
        if ($msg -match 'SSL|certificate|TLS') {
            throw "TLS/certificate error (not auth). Enable 'Skip TLS cert check'.`r`n$msg"
        }
        if ($status -gt 0) {
            throw "BigFix API HTTP $status`r`nURL: $uri`r`n$msg"
        }
        throw "Cannot reach BigFix API at $uri`r`n$msg"
    }
    throw "Cannot reach BigFix API at $uri (no credentials configured)."
}

function Invoke-BigFixQuery {
    param(
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$Relevance,
        [int]$Top = 5000,
        [int]$TimeoutSec = 25,
        [switch]$AsJson
    )
    $enc = [uri]::EscapeDataString($Relevance)
    if ($AsJson) {
        $path = "/api/query?relevance=$enc&top=$Top&output=json"
    } else {
        $path = "/api/query?relevance=$enc&top=$Top"
    }
    $resp = Invoke-BigFixRest -BaseUrl $BaseUrl -Headers $Headers -Path $path -TimeoutSec $TimeoutSec

    $answers = New-Object System.Collections.Generic.List[string]
    if ($null -eq $resp -or $resp -eq '') { return $answers }

    # JSON: {"result":["a","b"],...} or tuple rows as nested arrays
    if ($resp -isnot [xml] -and $resp.PSObject.Properties['result']) {
        foreach ($row in @($resp.result)) {
            if ($null -eq $row) { continue }
            if ($row -is [string] -or $row -is [ValueType]) {
                $answers.Add([string]$row)
            } elseif ($row -is [System.Collections.IEnumerable] -and $row -isnot [string]) {
                $parts = @()
                foreach ($cell in $row) { $parts += [string]$cell }
                $answers.Add(($parts -join '|'))
            } elseif ($row.PSObject) {
                # object with Answer / Value members
                if ($row.PSObject.Properties['Answer']) { $answers.Add([string]$row.Answer) }
                else { $answers.Add(($row | Out-String).Trim()) }
            }
        }
        return $answers
    }

    if ($resp -is [string]) {
        try { [xml]$resp = $resp } catch { $answers.Add([string]$resp); return $answers }
    }
    if ($resp -isnot [xml]) {
        if ($resp.Answer) {
            foreach ($a in @($resp.Answer)) { $answers.Add([string]$a) }
        } else {
            $answers.Add(($resp | Out-String).Trim())
        }
        return $answers
    }

    # XML: Query/Result/Tuple/Answer  or  Query/Result/Answer
    $tuples = $resp.SelectNodes('//Tuple')
    if ($tuples -and $tuples.Count -gt 0) {
        foreach ($t in $tuples) {
            $vals = @()
            foreach ($ans in $t.SelectNodes('./Answer')) {
                $v = $ans.SelectSingleNode('.//Value')
                if ($v) { $vals += $v.InnerText } else { $vals += $ans.InnerText }
            }
            if ($vals.Count -gt 0) { $answers.Add(($vals -join '|')) }
        }
        return $answers
    }

    $nodes = $resp.SelectNodes('//Answer')
    foreach ($a in $nodes) {
        $vals = $a.SelectNodes('.//Value')
        if ($vals -and $vals.Count -gt 0) {
            foreach ($v in $vals) { $answers.Add($v.InnerText) }
        } else {
            $t = $a.InnerText
            if ($t) { $answers.Add($t) }
        }
    }
    return $answers
}

function Invoke-BigFixQueryRetry {
    param(
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$Relevance,
        [int]$Top = 5000,
        [int]$TimeoutSec = 25,
        [switch]$AsJson,
        [int]$Retries = 2
    )
    $attempt = 0
    while ($true) {
        try {
            return Invoke-BigFixQuery -BaseUrl $BaseUrl -Headers $Headers -Relevance $Relevance -Top $Top -TimeoutSec $TimeoutSec -AsJson:$AsJson
        } catch {
            $attempt++
            $msg = "$($_.Exception.Message)"
            $transient = ($msg -match 'timed? ?out|503|502|500|reset|closed by|temporarily')
            if ($attempt -gt $Retries -or -not $transient) { throw }
            Start-Sleep -Seconds ([math]::Min(8, [math]::Pow(2, $attempt - 1)))
        }
    }
}

function Invoke-ParallelRelevanceQueries {
    <#
      Runs multiple relevance queries concurrently in a runspace pool (throttled).
      Each item: @{ Name; Relevance; Top; TimeoutSec }
      Returns hashtable Name -> string[] answers (or Name -> $null on error with _Error key).
    #>
    param(
        [array]$Queries,
        [string]$BaseUrl,
        [int]$MaxRunspaces = 3
    )
    $results = @{}
    if (-not $Queries -or $Queries.Count -eq 0) { return $results }
    if ($Queries.Count -eq 1) {
        $q = $Queries[0]
        try {
            $results[$q.Name] = @(Invoke-BigFixQueryRetry -BaseUrl $BaseUrl -Headers $null -Relevance $q.Relevance -Top $q.Top -TimeoutSec $q.TimeoutSec -AsJson)
        } catch {
            $results[$q.Name] = $null
            $results["$($q.Name)_Error"] = "$($_.Exception.Message)"
        }
        return $results
    }

    $pool = [runspacefactory]::CreateRunspacePool(1, [Math]::Max(1, $MaxRunspaces))
    $pool.Open()
    $jobs = @()
    try {
        foreach ($q in $Queries) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool
            # Inject dependency functions + auth session state into the worker runspace
            foreach ($fnName in @('Set-BigFixTls', 'Get-HttpStatusCode', 'Format-BigFixApiError', 'New-BigFixAuth', 'Invoke-BigFixWebRequest', 'Get-BigFixBaseUri', 'Invoke-BigFixRest', 'Invoke-BigFixQuery', 'Invoke-BigFixQueryRetry')) {
                $cmd = Get-Command -Name $fnName -CommandType Function -ErrorAction SilentlyContinue
                if ($cmd) {
                    $null = $ps.AddScript("function $($cmd.Name) { $($cmd.Definition) }")
                }
            }
            $null = $ps.AddScript({
                param($BaseUrl, $Relevance, $Top, $TimeoutSec, $User, $Pass, $SessionToken, $AuthMode, $Cred)
                $script:BigFixSessionToken = $SessionToken
                $script:BigFixAuthMode = $AuthMode
                $script:BigFixUsername = $User
                $script:BigFixPassword = $Pass
                $script:BigFixCredential = $Cred
                try {
                    $ans = Invoke-BigFixQueryRetry -BaseUrl $BaseUrl -Headers $null -Relevance $Relevance -Top $Top -TimeoutSec $TimeoutSec -AsJson
                    return ,@($ans)
                } catch {
                    return ,@('__ERROR__', "$($_.Exception.Message)")
                }
            })
            $null = $ps.AddArgument($BaseUrl)
            $null = $ps.AddArgument($q.Relevance)
            $null = $ps.AddArgument([int]$q.Top)
            $null = $ps.AddArgument([int]$q.TimeoutSec)
            $null = $ps.AddArgument([string]$script:BigFixUsername)
            $null = $ps.AddArgument([string]$script:BigFixPassword)
            $null = $ps.AddArgument([string]$script:BigFixSessionToken)
            $null = $ps.AddArgument([string]$script:BigFixAuthMode)
            $null = $ps.AddArgument($script:BigFixCredential)
            $jobs += [pscustomobject]@{ Name = $q.Name; Ps = $ps; Handle = $ps.BeginInvoke() }
        }
        foreach ($j in $jobs) {
            try {
                $out = $j.Ps.EndInvoke($j.Handle)
                $flat = @()
                foreach ($item in @($out)) {
                    if ($null -eq $item) { continue }
                    if ($item -is [string]) { $flat += $item }
                    elseif ($item -is [System.Collections.IEnumerable] -and $item -isnot [string]) {
                        foreach ($x in $item) {
                            if ($null -ne $x) { $flat += [string]$x }
                        }
                    } else { $flat += [string]$item }
                }
                if ($flat.Count -ge 2 -and $flat[0] -eq '__ERROR__') {
                    $results[$j.Name] = $null
                    $results["$($j.Name)_Error"] = $flat[1]
                } else {
                    $results[$j.Name] = $flat
                }
            } catch {
                $results[$j.Name] = $null
                $results["$($j.Name)_Error"] = "$($_.Exception.Message)"
            } finally {
                try { $j.Ps.Dispose() } catch { }
            }
        }
    } finally {
        try { $pool.Close() } catch { }
        try { $pool.Dispose() } catch { }
    }
    return $results
}

function ConvertFrom-BigFixGmtTime {
    param([string]$Text)
    $t = "$Text".Trim()
    if (-not $t) { return $null }
    # BigFix API times are GMT/UTC (e.g. 2024-01-15T12:34:56+00:00 or ...Z)
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($t, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$dt)) { return $dt }
    $dt2 = [datetime]::MinValue
    if ([datetime]::TryParse($t, [ref]$dt2)) { return $dt2 }
    return $null
}

# ---------------------------------------------------------------------------
# Session cache (offline fallback after a successful API pull)
# ---------------------------------------------------------------------------
function ConvertTo-ObjectArray {
    # WinPS 5.1: @($List[object]) hits PSToObjectArrayBinder -> "Argument types do not match".
    # Enumerate with foreach and return a real object[] / string[] instead.
    param($Value, [switch]$AsString)
    $out = New-Object System.Collections.Generic.List[object]
    if ($null -ne $Value) {
        foreach ($item in $Value) {
            if ($null -eq $item) { continue }
            if ($AsString) { $out.Add([string]$item) } else { $out.Add($item) }
        }
    }
    if ($AsString) { return ,([string[]]$out.ToArray()) }
    return ,([object[]]$out.ToArray())
}

function Get-ApiResultCollection {
    # Safe Count/enumerate for Failures/Devices/Log (may be List[object] or object[]).
    param($Result, [ValidateSet('Failures','Devices','Log')][string]$Name = 'Failures')
    if (-not $Result) { return ,([object[]]@()) }
    $prop = $Result.PSObject.Properties[$Name]
    if (-not $prop) { return ,([object[]]@()) }
    return ,(ConvertTo-ObjectArray -Value $prop.Value -AsString:($Name -eq 'Log'))
}

function ConvertTo-ApiResultJson {
    param($Result)
    $logArr = ConvertTo-ObjectArray -Value $Result.Log -AsString
    $devArr = ConvertTo-ObjectArray -Value $Result.Devices
    $failArr = ConvertTo-ObjectArray -Value $Result.Failures
    [pscustomobject]@{
        SavedAt   = (Get-Date).ToString('o')
        Version   = [string]$Result.Version
        ElapsedSec = $Result.ElapsedSec
        Log       = $logArr
        Devices   = $devArr
        Failures  = $failArr
    } | ConvertTo-Json -Depth 12
}

function ConvertFrom-ApiResultJson {
    param([string]$Json)
    $o = $Json | ConvertFrom-Json
    $failures = New-Object System.Collections.Generic.List[object]
    foreach ($f in (ConvertTo-ObjectArray -Value $o.Failures)) { if ($null -ne $f) { $failures.Add($f) } }
    $devices = New-Object System.Collections.Generic.List[object]
    foreach ($d in (ConvertTo-ObjectArray -Value $o.Devices)) { if ($null -ne $d) { $devices.Add($d) } }
    $log = New-Object System.Collections.Generic.List[string]
    foreach ($l in (ConvertTo-ObjectArray -Value $o.Log -AsString)) { if ($null -ne $l) { $log.Add([string]$l) } }
    [pscustomobject]@{
        Failures   = $failures
        Devices    = $devices
        Log        = $log
        Version    = [string]$o.Version
        ElapsedSec = $o.ElapsedSec
        SavedAt    = [string]$o.SavedAt
    }
}

function Save-SessionCache {
    param($Result)
    if (-not $Result) { return }
    if (-not (Test-Path -LiteralPath $script:CacheDir)) {
        New-Item -ItemType Directory -Path $script:CacheDir -Force | Out-Null
    }
    $json = ConvertTo-ApiResultJson -Result $Result
    Set-Content -LiteralPath $script:CachePath -Value $json -Encoding UTF8
}

function Get-SessionCache {
    if (-not (Test-Path -LiteralPath $script:CachePath)) { return $null }
    try {
        $json = Get-Content -LiteralPath $script:CachePath -Raw
        if ([string]::IsNullOrWhiteSpace($json)) { return $null }
        return ConvertFrom-ApiResultJson -Json $json
    } catch { return $null }
}

function Get-BigFixApiData {
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [string]$Username = '',
        [string]$Password = '',
        [int]$OfflineThresholdHours = 4,
        [switch]$SkipCertCheck,
        [switch]$InventoryOnly,
        [switch]$FailuresOnly,
        [object]$SeedData,
        [scriptblock]$Progress
    )
    Set-BigFixTls -SkipCertCheck:$SkipCertCheck
    if (-not $Username) {
        throw "Username is required for BigFix REST API.`r`nUse a BigFix Console operator that has 'Can use REST API' permission."
    }
    # FailuresOnly reuses sticky auth + SessionToken from inventory phase
    $headers = New-BigFixAuth -Username $Username -Password $Password -KeepSession:$FailuresOnly

    $log = New-Object System.Collections.Generic.List[string]
    $failures = New-Object System.Collections.Generic.List[object]
    $inventory = New-Object System.Collections.Generic.List[object]
    $version = ''
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $stepSw = [System.Diagnostics.Stopwatch]::StartNew()

    $note = {
        param([string]$Msg)
        $stepSw.Stop()
        $log.Add("[$([math]::Round($stepSw.Elapsed.TotalSeconds,1))s] $Msg")
        if ($Progress) { try { & $Progress $Msg } catch { } }
        $stepSw.Restart()
    }

    # FailuresOnly: seed from phase-1 result; skip help/computers/hostname
    if ($FailuresOnly) {
        if (-not $SeedData) { throw 'FailuresOnly requires -SeedData from inventory phase.' }
        foreach ($f in (ConvertTo-ObjectArray -Value $SeedData.Failures)) { if ($null -ne $f) { $failures.Add($f) } }
        foreach ($d in (ConvertTo-ObjectArray -Value $SeedData.Devices)) { if ($null -ne $d) { $inventory.Add($d) } }
        $version = [string]$SeedData.Version
        if ($SeedData.Log) { foreach ($l in (ConvertTo-ObjectArray -Value $SeedData.Log -AsString)) { if ($null -ne $l) { $log.Add($l) } } }
        & $note "Failure scan (reusing $($inventory.Count) devices / $($failures.Count) seed rows)"
    } else {
        # 1) Connectivity + help
        try {
            $help = Invoke-BigFixRest -BaseUrl $BaseUrl -Headers $headers -Path '/api/help' -TimeoutSec 8
            $mode = if ($script:BigFixAuthMode) { $script:BigFixAuthMode } else { 'n/a' }
            if ($help -is [string] -and $help) { $version = (($help -split "`n")[0]).Trim() }
            elseif ($help -is [xml]) { $version = 'XML help' }
            elseif ($help) { $version = 'OK' }
            $tokNote = if ($script:BigFixSessionToken) { 'SessionToken=y' } else { 'SessionToken=n' }
            & $note "API OK - /api/help (auth: $mode, $tokNote)"
            if (-not $version) { $version = 'help' }
        } catch {
            throw $_.Exception.Message
        }

        # 2) Client inventory via /api/computers
        $nameById = @{}
        $ipById = @{}
        $gotComputers = $false
        try {
            & $note "Loading /api/computers..."
            $comps = Invoke-BigFixRest -BaseUrl $BaseUrl -Headers $headers -Path '/api/computers' -TimeoutSec 20
        $compNodes = @()
        $rawPreview = ''
        if ($comps -is [string]) {
            $rawPreview = if ($comps.Length -gt 400) { $comps.Substring(0, 400) } else { $comps }
            try { [xml]$comps = $comps } catch { }
        }
        if ($comps -is [xml]) {
            $compNodes = @($comps.SelectNodes('//Computer|//computer|//Resource|//resource|//entry|//client|//Client'))
            if ($compNodes.Count -eq 0) {
                $compNodes = @($comps.SelectNodes('//*[local-name()="LastReportTime"]/..'))
            }
            $rootName = $comps.DocumentElement.LocalName
            $log.Add("Computers XML root=<$rootName> nodes=$($compNodes.Count)")
            if ($compNodes.Count -gt 0) {
                try { $rawPreview = $compNodes[0].OuterXml } catch { }
            }
        } elseif ($comps.computers) {
            $compNodes = @($comps.computers)
            if ($comps.computers.computer) { $compNodes = @($comps.computers.computer) }
        } elseif ($comps -is [System.Collections.IEnumerable] -and $comps -isnot [string]) {
            $compNodes = @($comps)
        }
        if ($rawPreview) { $log.Add("Computer node sample: $rawPreview") }

        $evalOk = 0
        $evalNoLast = 0
        $evalParseFail = 0
        $loggedOffline = 0
        foreach ($c in $compNodes) {
            $id = ''; $lastRaw = ''; $nm = ''; $ip = ''
            if ($c -is [System.Xml.XmlElement]) {
                foreach ($ch in $c.ChildNodes) {
                    $n = $ch.LocalName
                    if ($n -match '^(ID|Id|id)$') { $id = $ch.InnerText }
                    elseif ($n -match 'LastReport|ReportTime|LastReportTime|LastReported') { if (-not $lastRaw) { $lastRaw = $ch.InnerText } }
                    elseif ($n -match '^(Name|ComputerName|Hostname)$') { $nm = $ch.InnerText }
                    elseif ($n -match 'IP|IPAddress') { $ip = $ch.InnerText }
                }
                if (-not $id -and $c.HasAttribute('id')) { $id = $c.GetAttribute('id') }
                if (-not $id -and $c.HasAttribute('Resource')) {
                    # e.g. Resource="api/computer/2785212"
                    if ($c.GetAttribute('Resource') -match 'computer/(\d+)') { $id = $Matches[1] }
                }
                if (-not $lastRaw -and $c.HasAttribute('LastReportTime')) { $lastRaw = $c.GetAttribute('LastReportTime') }
                if (-not $nm -and $c.HasAttribute('name')) { $nm = $c.GetAttribute('name') }
            } else {
                if ($c.PSObject.Properties['ID']) { $id = [string]$c.ID }
                elseif ($c.PSObject.Properties['Id']) { $id = [string]$c.Id }
                elseif ($c.PSObject.Properties['id']) { $id = [string]$c.id }
                if ($c.PSObject.Properties['LastReportTime']) { $lastRaw = [string]$c.LastReportTime }
                elseif ($c.PSObject.Properties['lastReportTime']) { $lastRaw = [string]$c.lastReportTime }
                if ($c.PSObject.Properties['Name']) { $nm = [string]$c.Name }
                elseif ($c.PSObject.Properties['name']) { $nm = [string]$c.name }
                if ($c.PSObject.Properties['IPAddress']) { $ip = [string]$c.IPAddress }
                if (-not $id -and $c.PSObject.Properties['Resource'] -and "$($c.Resource)" -match 'computer/(\d+)') {
                    $id = $Matches[1]
                }
            }
            if (-not $id -and -not $lastRaw -and -not $nm) { continue }
            if ($nm) { if ($id) { $nameById[$id] = $nm } }
            if ($ip) { if ($id) { $ipById[$id] = $ip } }
            if (-not $lastRaw) {
                $evalNoLast++
                if ($evalNoLast -le 5) { $log.Add("Computer parse: id=$id name=$nm - NO LastReportTime") }
                continue
            }
            $dt = ConvertFrom-BigFixGmtTime -Text $lastRaw
            if (-not $dt) {
                $evalParseFail++
                if ($evalParseFail -le 5) { $log.Add("Computer parse: id=$id name=$nm last='$lastRaw' - UNPARSEABLE time") }
                continue
            }
            $hours = ((Get-Date).ToUniversalTime() - $dt.ToUniversalTime()).TotalHours
            $devName = if ($nm) { $nm } else { "Computer #$id" }
            $inv = [pscustomobject]@{
                Id = $id; Name = $devName; DnsName = ''; Ip = $ip; Hours = [math]::Round($hours, 1)
                LastLocal = $dt.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss'); LastRaw = $lastRaw
            }
            $inventory.Add($inv)
            $evalOk++
            if ($hours -ge $OfflineThresholdHours) {
                $failures.Add((New-FailureRecord `
                    -DeviceName $devName `
                    -IPAddress $ip `
                    -LastReportTime $dt.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') `
                    -Category 'Offline' `
                    -Severity 'Critical' `
                    -Source 'BigFix API' `
                    -Component 'BESClient' `
                    -ErrorCode "ID=$id" `
                    -ErrorMessage "Client not reporting - last report $([math]::Round($hours,1)) hours ago (threshold ${OfflineThresholdHours}h; LastReportTime treated as GMT)" `
                    -Status 'Offline'))
                if ($loggedOffline -lt 25) {
                    $log.Add("  [OFFLINE] $devName (ID=$id) last=$($dt.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')) ($([math]::Round($hours,1))h)")
                    $loggedOffline++
                }
            }
        }
        $gotComputers = ($compNodes.Count -gt 0)
        $offlineCount = @($inventory | Where-Object { $_.Hours -ge $OfflineThresholdHours }).Count
        & $note "/api/computers: nodes=$($compNodes.Count) timeParsed=$evalOk noTime=$evalNoLast badTime=$evalParseFail offline=$offlineCount threshold=${OfflineThresholdHours}h"
    } catch {
        $log.Add("/api/computers failed: $($_.Exception.Message)")
    }

    # 2d) One hostname query (Computer Name | IP) - single try, short timeout
    if ($inventory.Count -gt 0) {
        $needName = @($inventory | Where-Object { $_.Name -eq '' -or $_.Name -like 'Computer #*' -or -not $_.Ip }).Count
        & $note "Resolving hostnames ($needName of $($inventory.Count) need names/IPs)..."
        $nameRel = '(id of it as string & "|" & name of it & "|" & (if (exists ip address of it) then ip address of it as string else "")) of bes computers'
        $answers = $null
        try {
            $answers = Invoke-BigFixQuery -BaseUrl $BaseUrl -Headers $headers -Relevance $nameRel -Top 10000 -TimeoutSec 15 -AsJson
        } catch {
            $log.Add("Hostname query failed: $($_.Exception.Message)")
        }
        if ($answers -and $answers.Count -gt 0) {
            $mapId = @{}
            $mapName = @{}
            $mapIp = @{}
            $resolved = 0
            foreach ($line in $answers) {
                $parts = "$line" -split '\|', 3
                if ($parts.Count -lt 2) { continue }
                $cid = $parts[0].Trim()
                $cname = $parts[1].Trim()
                $cip = if ($parts.Count -gt 2) { $parts[2].Trim() } else { '' }
                if (-not $cid) { continue }
                $mapId[$cid] = $true
                if ($cname) { $mapName[$cid] = $cname }
                if ($cip) { $mapIp[$cid] = $cip }
            }
            foreach ($inv in $inventory) {
                $cid = [string]$inv.Id
                $old = [string]$inv.Name
                $newName = $mapName[$cid]

                if ($newName -and $newName -match '^([^.]+)\.(.+)$') {
                    $short = $Matches[1]
                    $fqdn = $newName
                    $inv | Add-Member -NotePropertyName ShortName -NotePropertyValue $short -Force
                    $inv | Add-Member -NotePropertyName Fqdn -NotePropertyValue $fqdn -Force
                    $inv | Add-Member -NotePropertyName DnsName -NotePropertyValue $fqdn -Force
                    $newName = $short
                } elseif ($newName) {
                    $inv | Add-Member -NotePropertyName ShortName -NotePropertyValue $newName -Force
                }

                if ($newName) { $inv.Name = $newName }
                if (-not $inv.Ip -and $mapIp[$cid]) { $inv.Ip = $mapIp[$cid] }
                if ($newName -and ($old -ne $newName -or $old -like 'Computer #*')) { $resolved++ }

                foreach ($f in $failures) {
                    if ($f.Category -ne 'Offline') { continue }
                    if ($f.ErrorCode -eq "ID=$cid" -or ($old -and $f.DeviceName -eq $old)) {
                        if ($newName) { $f.DeviceName = $newName }
                        if ($inv.Ip) { $f.IPAddress = $inv.Ip }
                    }
                }
            }
            $sample = if ($inventory.Count -gt 0) { [string]$inventory[0].Name } else { '' }
            & $note "Hostname map: $($mapId.Count) ids, renamed=$resolved sample='$sample'"
        } else {
            $log.Add("Hostname query empty - keeping IDs from /api/computers")
        }
    }

    if (-not $gotComputers) {
        & $note "Computers empty - relevance inventory fallback..."
        $clientRel = '(id of it as string & "|" & name of it & "|" & (if (exists ip address of it) then ip address of it as string else "") & "|" & (last report time of it as string)) of bes computers'
        try {
            $answers = Invoke-BigFixQuery -BaseUrl $BaseUrl -Headers $headers -Relevance $clientRel -Top 10000 -TimeoutSec 20 -AsJson
            $rows = 0
            foreach ($line in $answers) {
                $parts = "$line" -split '\|', 4
                if ($parts.Count -lt 4) { continue }
                $rows++
                $dt = ConvertFrom-BigFixGmtTime -Text $parts[3]
                if (-not $dt) { continue }
                $hours = ((Get-Date).ToUniversalTime() - $dt.ToUniversalTime()).TotalHours
                $devName = $parts[1].Trim()
                if (-not $devName) { $devName = "Computer $($parts[0].Trim())" }
                $inv = [pscustomobject]@{
                    Id = $parts[0].Trim(); Name = $devName; DnsName = ''; Ip = $parts[2].Trim()
                    Hours = [math]::Round($hours, 1)
                    LastLocal = $dt.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss'); LastRaw = $parts[3]
                }
                $inventory.Add($inv)
                if ($hours -ge $OfflineThresholdHours) {
                    $failures.Add((New-FailureRecord `
                        -DeviceName $devName `
                        -IPAddress $inv.Ip `
                        -LastReportTime $inv.LastLocal `
                        -Category 'Offline' `
                        -Severity 'Critical' `
                        -Source 'BigFix API' `
                        -Component 'BESClient' `
                        -ErrorCode "ID=$($inv.Id)" `
                        -ErrorMessage "Client not reporting - last report $([math]::Round($hours,1)) hours ago (threshold ${OfflineThresholdHours}h; LastReportTime treated as GMT)" `
                        -Status 'Offline'))
                }
            }
            $gotComputers = $true
            & $note "Client relevance fallback: $rows rows"
        } catch {
            $log.Add("Client relevance fallback failed: $($_.Exception.Message)")
        }
    }

    # Inventory-only: show devices immediately
    if ($InventoryOnly) {
        Add-HealthyRows -Failures $failures -Inventory $inventory -ThresholdHours $OfflineThresholdHours -Log $log
        $sum = Get-FailureSummary $failures
        & $note "INVENTORY-ONLY done in $([math]::Round($sw.Elapsed.TotalSeconds,1))s total=$($sum.Total) healthy=$($sum.Healthy) offline=$($sum.Offline)"
        return [pscustomobject]@{
            Failures = $failures
            Devices  = $inventory
            Log      = $log
            Version  = $version
            ElapsedSec = [math]::Round($sw.Elapsed.TotalSeconds,1)
        }
    }
    }   # end inventory else-block

    # Failure enrichment (FailuresOnly or full path) - short timeouts
    try {
        & $note "Loading /api/actions (list only)..."
        $actions = Invoke-BigFixRest -BaseUrl $BaseUrl -Headers $headers -Path '/api/actions' -TimeoutSec 12
        if ($actions -is [string]) { try { [xml]$actions = $actions } catch { } }
        $actionNodes = @()
        if ($actions -is [xml]) { $actionNodes = @($actions.SelectNodes('//Action|//action')) }
        elseif ($actions.actions) {
            $actionNodes = @($actions.actions)
            if ($actions.actions.action) { $actionNodes = @($actions.actions.action) }
        } elseif ($actions -is [System.Collections.IEnumerable] -and $actions -isnot [string]) {
            $actionNodes = @($actions)
        }
        $failedActions = 0
        $listed = 0
        foreach ($a in $actionNodes) {
            $title = ''; $id = ''; $state = ''
            if ($a -is [System.Xml.XmlElement]) {
                foreach ($ch in $a.ChildNodes) {
                    switch -Regex ($ch.Name) {
                        '^(Name|Title)$' { $title = $ch.InnerText }
                        '^(ID|Id|id)$' { $id = $ch.InnerText }
                        'Status|State' { $state = $ch.InnerText }
                    }
                }
            } else {
                if ($a.PSObject.Properties['Name']) { $title = [string]$a.Name }
                elseif ($a.PSObject.Properties['Title']) { $title = [string]$a.Title }
                if ($a.PSObject.Properties['ID']) { $id = [string]$a.ID }
                elseif ($a.PSObject.Properties['Id']) { $id = [string]$a.Id }
                if ($a.PSObject.Properties['State']) { $state = [string]$a.State }
                elseif ($a.PSObject.Properties['Status']) { $state = [string]$a.Status }
            }
            if (-not $id) { continue }
            $listed++
            if ($state -match 'Fail|Error') {
                $failedActions++
                $failures.Add((New-FailureRecord `
                    -DeviceName "Action #$id" `
                    -Category 'FixletFailure' `
                    -Source 'BigFix API' `
                    -Component $title `
                    -ErrorCode "ActionID=$id" `
                    -ErrorMessage "BigFix action state: $state" `
                    -Status $state))
            }
        }
        & $note "Actions listed: $listed, failed-in-list: $failedActions"
    } catch {
        $log.Add("Actions query failed: $($_.Exception.Message)")
    }

    # Failure relevance - parallel pass with retry/backoff (fixlets + compliance)
    & $note "Querying failed fixlets / compliance (parallel)..."
    $failedFixletRel = '(name of it & "|" & (concatenation of ", " of (names of bes actions whose (state of it as string as lowercase contains "fail" or status of it as string as lowercase contains "fail") of it))) of bes computers'
    $compRel = '(names of bes computers of bes fixlets whose ((name of it as lowercase contains "baseline" or name of it as lowercase contains "compliance" or name of it as lowercase contains "cis " or name of it as lowercase contains "stig") and (exists bes action whose (state of it as string as lowercase contains "fail" or status of it as string as lowercase contains "fail") of it)))'
    $pq = Invoke-ParallelRelevanceQueries -BaseUrl $BaseUrl -MaxRunspaces 3 -Queries @(
        @{ Name = 'Fixlets'; Relevance = $failedFixletRel; Top = 5000; TimeoutSec = 20 },
        @{ Name = 'Compliance'; Relevance = $compRel; Top = 5000; TimeoutSec = 15 }
    )
    if ($pq['Fixlets_Error']) { $log.Add("Failed-fixlet query failed/timed out: $($pq['Fixlets_Error'])") }
    foreach ($line in @($pq['Fixlets'])) {
        if ($null -eq $line) { continue }
        $idx = "$line".IndexOf('|')
        if ($idx -lt 1) { continue }
        $dev = $line.Substring(0, $idx).Trim()
        $detail = $line.Substring($idx + 1).Trim()
        if (-not $detail -or $detail -eq 'Error: ') { continue }
        if ($detail -match '^Error') { continue }
        $failures.Add((New-FailureRecord `
            -DeviceName $dev `
            -Category 'FixletFailure' `
            -Source 'BigFix API' `
            -Component $detail `
            -ErrorMessage "Failed fixlet/action on device: $detail" `
            -Status 'Failed'))
    }
    $log.Add("Failed-fixlet rows: $(@($pq['Fixlets'] | Where-Object { $_ }).Count)")

    if ($pq['Compliance_Error']) { $log.Add("Compliance relevance query failed: $($pq['Compliance_Error'])") }
    foreach ($dev in @($pq['Compliance'])) {
        if ($null -eq $dev) { continue }
        $d = "$dev".Trim()
        if (-not $d -or $d -match '^Error') { continue }
        $failures.Add((New-FailureRecord `
            -DeviceName $d `
            -Category 'ComplianceFailure' `
            -Source 'BigFix API' `
            -Component 'Compliance baseline' `
            -ErrorMessage 'Compliance/baseline fixlet has failed action results on this device' `
            -Status 'Failed'))
    }
    $log.Add("Compliance failure rows: $(@($pq['Compliance'] | Where-Object { $_ }).Count)")

    Add-HealthyRows -Failures $failures -Inventory $inventory -ThresholdHours $OfflineThresholdHours -Log $log

    $sum = Get-FailureSummary $failures
    $sw.Stop()
    & $note "DONE in $([math]::Round($sw.Elapsed.TotalSeconds,1))s total=$($sum.Total) healthy=$($sum.Healthy) offline=$($sum.Offline) fixlet=$($sum.FixletFailures) compliance=$($sum.Compliance) devices=$($sum.Devices)"

    $result = [pscustomobject]@{
        Failures = $failures
        Devices  = $inventory
        Log      = $log
        Version  = $version
        ElapsedSec = [math]::Round($sw.Elapsed.TotalSeconds,1)
    }
    try { Save-SessionCache -Result $result } catch { }
    return $result
}

function Add-HealthyRows {
    param(
        $Failures,
        $Inventory,
        [int]$ThresholdHours,
        $Log
    )
    if (-not $Inventory -or $Inventory.Count -eq 0) { return }
    $failedDeviceNames = @{}
    foreach ($f in $Failures) {
        if ($f.DeviceName) { $failedDeviceNames[[string]$f.DeviceName] = $true }
    }
    $healthyAdded = 0
    foreach ($inv in $Inventory) {
        if ($failedDeviceNames.ContainsKey([string]$inv.Name)) { continue }
        $dns = ''
        if ($inv.PSObject.Properties['DnsName']) { $dns = [string]$inv.DnsName }
        if (-not $dns -and $inv.PSObject.Properties['Fqdn']) { $dns = [string]$inv.Fqdn }
        $errMsg = "Reporting healthy - last report $($inv.Hours)h ago (threshold ${ThresholdHours}h)"
        if ($dns -and $dns -ne $inv.Name) { $errMsg = "$errMsg | FQDN/DNS: $dns" }
        $Failures.Add((New-FailureRecord `
            -DeviceName $inv.Name `
            -IPAddress $inv.Ip `
            -LastReportTime $inv.LastLocal `
            -Category 'Healthy' `
            -Severity 'Low' `
            -Source 'BigFix API' `
            -Component 'BESClient' `
            -ErrorCode "ID=$($inv.Id)" `
            -ErrorMessage $errMsg `
            -Status 'Healthy'))
        $healthyAdded++
    }
    if ($Inventory.Count -gt 0 -and $Log) {
        $Log.Add("Healthy inventory rows added: $healthyAdded (inventory=$($Inventory.Count))")
    }
}

# ---------------------------------------------------------------------------
# Aggregation / export
# ---------------------------------------------------------------------------
function Get-FailureSummary {
    param([System.Collections.IEnumerable]$Records)
    $all = ConvertTo-ObjectArray -Value $Records
    [pscustomobject]@{
        Total          = $all.Count
        Offline        = @($all | Where-Object { $_.Category -eq 'Offline' }).Count
        FixletFailures = @($all | Where-Object { $_.Category -eq 'FixletFailure' }).Count
        Compliance     = @($all | Where-Object { $_.Category -eq 'ComplianceFailure' }).Count
        SoftwareDeploy = @($all | Where-Object { $_.Category -eq 'SoftwareDeployment' }).Count
        Healthy        = @($all | Where-Object { $_.Category -eq 'Healthy' }).Count
        Critical       = @($all | Where-Object { $_.SeverityRank -ge 4 }).Count
        High           = @($all | Where-Object { $_.SeverityRank -eq 3 }).Count
        Devices        = @($all | Select-Object -ExpandProperty DeviceName -Unique).Count
    }
}

function Export-FailuresCsv {
    param(
        [System.Collections.IEnumerable]$Records,
        [Parameter(Mandatory)][string]$Path
    )
    $recArr = ConvertTo-ObjectArray -Value $Records
    $sorted = @($recArr | Sort-Object -Property @{ Expression = { [int]$_.SeverityRank }; Descending = $true }, DeviceName)
    $rows = $sorted | Select-Object DeviceName, IPAddress, LastReportTime, HoursSinceReport,
        Category, Severity, Source, Component, ErrorCode, ErrorMessage, Status,
        RootCause, Remediation, MatchedRule
    # UTF-8 with BOM so Excel opens international characters correctly
    $csv = @($rows | ConvertTo-Csv -NoTypeInformation)
    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines($Path, $csv, $enc)
}

function Export-FailuresJson {
    param(
        [System.Collections.IEnumerable]$Records,
        [Parameter(Mandatory)][string]$Path
    )
    $recArr = ConvertTo-ObjectArray -Value $Records
    $payload = [pscustomobject]@{
        GeneratedAt = (Get-Date).ToString('o')
        Tool        = "BigFix-FailureDashboard/$script:AppVersion"
        Count       = $recArr.Count
        Records     = $recArr
    }
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, ($payload | ConvertTo-Json -Depth 8), $enc)
}

function Get-HtmlEncode {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function New-FailureHtmlReport {
    param(
        [System.Collections.IEnumerable]$Records,
        [string]$Title = 'BigFix Failed Device Report'
    )
    $all = ConvertTo-ObjectArray -Value $Records
    $sum = Get-FailureSummary $all
    $gen = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="en"><head><meta charset="utf-8" />')
    [void]$sb.AppendLine("<title>$(Get-HtmlEncode $Title)</title>")
    [void]$sb.AppendLine(@'
<style>
  :root { color-scheme: light; }
  body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1a1a1a; background: #f5f7fa; }
  h1 { font-size: 1.5rem; color: #1e385c; margin-bottom: 4px; }
  .meta { color: #667; font-size: 0.9rem; margin-bottom: 16px; }
  .tiles { display: flex; flex-wrap: wrap; gap: 10px; margin: 16px 0; }
  .tile { background: #fff; border: 1px solid #dde3ea; border-radius: 8px; padding: 10px 16px; min-width: 110px; text-align: center; }
  .tile .v { font-size: 1.6rem; font-weight: 700; color: #1e385c; }
  .tile .t { font-size: 0.75rem; color: #667; text-transform: uppercase; letter-spacing: 0.04em; }
  .tile.crit .v { color: #b00020; }
  .tile.ok .v { color: #1b7a3d; }
  table { border-collapse: collapse; width: 100%; background: #fff; font-size: 0.85rem; }
  th, td { border: 1px solid #dde3ea; padding: 6px 8px; text-align: left; vertical-align: top; }
  th { background: #1e385c; color: #fff; position: sticky; top: 0; }
  tr:nth-child(even) { background: #f8f9fb; }
  .sev-Critical { color: #b00020; font-weight: 700; }
  .sev-High { color: #c45c00; font-weight: 600; }
  .sev-Medium { color: #8a6d00; }
  .sev-Low { color: #667; }
  .cat-Offline { color: #b00020; }
  .cat-ComplianceFailure { color: #6a1b9a; }
  .cat-SoftwareDeployment { color: #00695c; }
  .cat-Healthy { color: #1b7a3d; }
  .cat-FixletFailure { color: #c45c00; }
  footer { margin-top: 20px; color: #667; font-size: 0.8rem; }
  @media print { body { background: #fff; margin: 0; } .tile { break-inside: avoid; } table { font-size: 0.7rem; } }
</style>
</head><body>
'@)
    [void]$sb.AppendLine("<h1>$(Get-HtmlEncode $Title)</h1>")
    [void]$sb.AppendLine("<div class='meta'>Generated $gen &mdash; BigFix Failure Dashboard v$(Get-HtmlEncode $script:AppVersion) &mdash; $($all.Count) records</div>")
    [void]$sb.AppendLine('<div class="tiles">')
    foreach ($t in @(
        @{ C = $sum.Total; L = 'Records'; Cl = '' },
        @{ C = $sum.Devices; L = 'Devices'; Cl = '' },
        @{ C = $sum.Offline; L = 'Offline'; Cl = 'crit' },
        @{ C = $sum.FixletFailures; L = 'Fixlet'; Cl = '' },
        @{ C = $sum.Compliance; L = 'Compliance'; Cl = '' },
        @{ C = $sum.SoftwareDeploy; L = 'Software'; Cl = '' },
        @{ C = ($sum.Critical + $sum.High); L = 'Crit+High'; Cl = 'crit' },
        @{ C = $sum.Healthy; L = 'Healthy'; Cl = 'ok' }
    )) {
        [void]$sb.AppendLine("<div class='tile $($t.Cl)'><div class='v'>$($t.C)</div><div class='t'>$($t.L)</div></div>")
    }
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<table><thead><tr>')
    foreach ($h in @('Device', 'IP', 'Category', 'Severity', 'Component', 'Error', 'Root Cause', 'Remediation', 'Rule')) {
        [void]$sb.AppendLine("<th>$h</th>")
    }
    [void]$sb.AppendLine('</tr></thead><tbody>')
    foreach ($r in ($all | Sort-Object -Property @{ Expression = { [int]$_.SeverityRank }; Descending = $true }, DeviceName)) {
        $sevClass = "sev-$($r.Severity)"
        $catClass = "cat-$($r.Category)"
        [void]$sb.AppendLine(("<tr><td>{0}</td><td>{1}</td><td class='{2}'>{3}</td><td class='{4}'>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td></tr>" -f `
            (Get-HtmlEncode $r.DeviceName), (Get-HtmlEncode $r.IPAddress), $catClass, (Get-HtmlEncode $r.Category),
            $sevClass, (Get-HtmlEncode $r.Severity), (Get-HtmlEncode $r.Component), (Get-HtmlEncode $r.ErrorMessage),
            (Get-HtmlEncode $r.RootCause), (Get-HtmlEncode $r.Remediation), (Get-HtmlEncode $r.MatchedRule)))
    }
    [void]$sb.AppendLine('</tbody></table>')
    [void]$sb.AppendLine("<footer>Internal operational data &mdash; handle as confidential. Exported from BigFix-FailureDashboard.</footer>")
    [void]$sb.AppendLine('</body></html>')
    return $sb.ToString()
}

function Export-FailuresHtml {
    param(
        [System.Collections.IEnumerable]$Records,
        [Parameter(Mandatory)][string]$Path,
        [string]$Title = 'BigFix Failed Device Report'
    )
    $html = New-FailureHtmlReport -Records $Records -Title $Title
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $html, $enc)
}

function Send-FailureDigest {
    param(
        [System.Collections.IEnumerable]$Records,
        $Smtp,
        [string]$SubjectPrefix = 'BigFix Failure Digest'
    )
    if (-not $Smtp -or -not [bool]$Smtp.Enable) { throw 'SMTP is not enabled in settings. Open Settings and check Enable email.' }
    if (-not $Smtp.Host) { throw 'SMTP host is not configured. Open Settings.' }
    if (-not $Smtp.From) { throw 'SMTP From address is not configured. Open Settings.' }
    if (-not $Smtp.To) { throw 'SMTP To address is not configured. Open Settings.' }
    if ($Smtp.Username -and -not $script:SmtpPassword -and -not $env:BIGFIX_SMTP_PASSWORD) {
        throw 'SMTP username is set but no password is available for this session. Open Settings and re-enter the password (session-only), or set $env:BIGFIX_SMTP_PASSWORD.'
    }
    if (-not $script:SmtpPassword -and $env:BIGFIX_SMTP_PASSWORD) { $script:SmtpPassword = $env:BIGFIX_SMTP_PASSWORD }
    $recordsSafe = ConvertTo-ObjectArray -Value $Records
    if ($recordsSafe.Count -eq 0) { throw 'No records to email.' }
    $sum = Get-FailureSummary $recordsSafe
    $html = New-FailureHtmlReport -Records $recordsSafe -Title $SubjectPrefix
    $subject = "{0} - {1} records, {2} offline, {3} critical+high - {4}" -f `
        $SubjectPrefix, $sum.Total, $sum.Offline, ($sum.Critical + $sum.High), (Get-Date -Format 'yyyy-MM-dd')
    $mail = @{
        SmtpServer  = [string]$Smtp.Host
        Port        = [int]$Smtp.Port
        UseSsl      = [bool]$Smtp.UseSsl
        From        = [string]$Smtp.From
        To          = @(([string]$Smtp.To) -split ';\s*' | Where-Object { $_ })
        Subject     = $subject
        Body        = $html
        BodyAsHtml  = $true
        ErrorAction = 'Stop'
    }
    if ($Smtp.Username) {
        # Password is not persisted; prompt only when sending from GUI if needed via session var
        if ($script:SmtpPassword) {
            $sec = ConvertTo-SecureString -String $script:SmtpPassword -AsPlainText -Force
            $mail.Credential = New-Object System.Management.Automation.PSCredential([string]$Smtp.Username, $sec)
        }
    }
    Send-MailMessage @mail
}

function Send-TicketWebhook {
    param(
        [System.Collections.IEnumerable]$Records,
        $Ticket,
        [string]$Source = 'BigFix-FailureDashboard'
    )
    if (-not $Ticket -or -not [bool]$Ticket.Enable) { throw 'Ticket webhook is not enabled in settings. Open Settings and check Enable ticket webhook.' }
    if (-not $Ticket.WebhookUrl) { throw 'Ticket webhook URL is not configured. Open Settings.' }
    $all = ConvertTo-ObjectArray -Value $Records
    if ($all.Count -eq 0) { throw 'No records to post to the ticket webhook.' }
    $payload = [ordered]@{
        source      = $Source
        toolVersion = $script:AppVersion
        generatedAt = (Get-Date).ToString('o')
        count       = $all.Count
        failures    = @($all | ForEach-Object {
            [ordered]@{
                device      = $_.DeviceName
                ip          = $_.IPAddress
                category    = $_.Category
                severity    = $_.Severity
                component   = $_.Component
                errorCode   = $_.ErrorCode
                errorMessage= $_.ErrorMessage
                rootCause   = $_.RootCause
                remediation = $_.Remediation
                matchedRule = $_.MatchedRule
                lastReport  = $_.LastReportTime
                status      = $_.Status
            }
        })
    }
    $headers = @{}
    if ($Ticket.HeaderJson) {
        try {
            $h = $Ticket.HeaderJson | ConvertFrom-Json
            foreach ($p in $h.PSObject.Properties) { $headers[$p.Name] = [string]$p.Value }
        } catch { }
    }
    $body = ($payload | ConvertTo-Json -Depth 8)
    $irm = @{
        Uri         = [string]$Ticket.WebhookUrl
        Method      = 'Post'
        Body        = $body
        ContentType = 'application/json'
        ErrorAction = 'Stop'
    }
    if ($headers.Count -gt 0) { $irm.Headers = $headers }
    $resp = Invoke-WebRequest @irm
    return [pscustomobject]@{ StatusCode = [int]$resp.StatusCode; Content = [string]$resp.Content }
}

function Get-DeviceAggregates {
    <#
      One row per device for Device view: worst severity/category + failure count + latest error.
      Timeline is not embedded here - Show-SelectedDetails expands from AllRecords.
    #>
    param([System.Collections.IEnumerable]$Records)
    $recArr = ConvertTo-ObjectArray -Value $Records
    $groups = @($recArr | Group-Object -Property DeviceName)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($g in $groups) {
        $items = ConvertTo-ObjectArray -Value $g.Group
        $failed = @($items | Where-Object { $_.Category -ne 'Healthy' })
        $use = $items
        if ($failed.Count -gt 0) { $use = [object[]]$failed }
        $worst = $use | Sort-Object -Property SeverityRank -Descending | Select-Object -First 1
        $latest = $use | Sort-Object -Property LastReportTime -Descending | Select-Object -First 1
        $cats = @($use | Select-Object -ExpandProperty Category -Unique) -join ', '
        $list.Add([pscustomobject]@{
            DeviceName       = $g.Name
            IPAddress        = ($use | Where-Object { $_.IPAddress } | Select-Object -First 1 -ExpandProperty IPAddress)
            LastReportTime   = $worst.LastReportTime
            HoursSinceReport = $worst.HoursSinceReport
            Category         = if ($failed.Count -gt 0) { $worst.Category } else { 'Healthy' }
            Severity         = $worst.Severity
            SeverityRank     = $worst.SeverityRank
            Component        = "$($use.Count) failure(s): $cats"
            ErrorCode        = ''
            ErrorMessage     = if ($latest) { $latest.ErrorMessage } else { '' }
            Status           = $worst.Status
            Source           = $worst.Source
            RootCause        = $worst.RootCause
            Remediation      = $worst.Remediation
            MatchedRule      = $worst.MatchedRule
            FailureCount     = $use.Count
            _Categories      = $cats
        })
    }
    return @($list | Sort-Object -Property @{ Expression = { [int]$_.SeverityRank }; Descending = $true }, DeviceName)
}

function Get-RunbookTopicText {
    param([hashtable]$Topic)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine($Topic.Title)
    [void]$sb.AppendLine(('=' * [Math]::Min(72, $Topic.Title.Length)))
    [void]$sb.AppendLine('')
    if ($Topic.Tags) { [void]$sb.AppendLine("TAGS: $($Topic.Tags)"); [void]$sb.AppendLine('') }
    [void]$sb.AppendLine('ISSUE')
    [void]$sb.AppendLine("  $($Topic.Summary)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('ROOT CAUSE')
    [void]$sb.AppendLine("  $($Topic.RootCause)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('SOLUTION / STEPS')
    $i = 1
    foreach ($s in @($Topic.Steps)) {
        [void]$sb.AppendLine(("  {0}. {1}" -f $i, $s))
        $i++
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('REFERENCES (HCL BigFix docs / KB / community)')
    foreach ($r in @($Topic.Refs)) {
        [void]$sb.AppendLine("  $r")
    }
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Fix plan / remote execution (WinRM)
# ---------------------------------------------------------------------------
function Get-FixSolutionId {
    param($Record)
    if (-not $Record) { return '' }
    $rule = "$($Record.MatchedRule)".Trim()
    if ($rule -and $rule -ne 'Default' -and $script:FixSolutions.ContainsKey($rule)) { return $rule }
    $cat = "$($Record.Category)"
    if ($script:FixCategoryDefaults.ContainsKey($cat)) { return $script:FixCategoryDefaults[$cat] }
    return 'BF-ACCESS-DEFAULT'
}

function Test-RemotableDeviceName {
    param([string]$Name)
    if (-not $Name) { return $false }
    if ($Name -match '^(Action\s+#|ID:)') { return $false }
    if ($Name -match '[\\\/\s]') { return $false }
    if ($Name -match '[^A-Za-z0-9._-]') { return $false }
    return $true
}

function Get-FixSeverityRank {
    param([string]$Severity)
    switch -Regex ("$Severity") {
        '^Critical' { return 4 }
        '^High'     { return 3 }
        '^Medium'   { return 2 }
        '^Low'      { return 1 }
        default     { return 0 }
    }
}

function Get-FixPlan {
    <#
      Builds the ordered list of solutions to run for the given records.
      Only solutions relevant to each record's MatchedRule/category are included.
      MinSeverity: '' = all; 'High' keeps Critical+High; 'Critical' keeps Critical only.
      Returns: DeviceName, IPAddress, SolutionId, Title, ManualOnly, Category,
               Severity, MatchedRule, ErrorMessage, Script, Remotable
    #>
    param(
        [System.Collections.IEnumerable]$Records,
        [string]$MinSeverity = ''
    )

    $minRank = 0
    if ($MinSeverity) { $minRank = Get-FixSeverityRank -Severity $MinSeverity }

    $plan = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($r in (ConvertTo-ObjectArray -Value $Records)) {
        if (-not $r) { continue }
        if ("$($r.Category)" -eq 'Healthy') { continue }
        if ($minRank -gt 0 -and (Get-FixSeverityRank -Severity "$($r.Severity)") -lt $minRank) { continue }
        $dev = "$($r.DeviceName)".Trim()
        if (-not $dev) { continue }

        $sid = Get-FixSolutionId -Record $r
        if (-not $sid -or -not $script:FixSolutions.ContainsKey($sid)) { continue }
        $sol = $script:FixSolutions[$sid]
        $key = "$dev`|$sid"
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true

        $manual = [bool]$sol.Manual
        $plan.Add([pscustomobject]@{
            DeviceName  = $dev
            IPAddress   = "$($r.IPAddress)"
            SolutionId  = $sid
            Title       = "$($sol.Title)"
            ManualOnly  = $manual
            Category    = "$($r.Category)"
            Severity    = "$($r.Severity)"
            MatchedRule = "$($r.MatchedRule)"
            ErrorMessage= "$($r.ErrorMessage)"
            Script      = if ($manual) { '' } else { "$($sol.Script)" }
            Remotable   = (Test-RemotableDeviceName -Name $dev)
        })
    }
    return @($plan.ToArray())
}

function Format-FixPlanPreview {
    param($Plan)
    $lines = New-Object System.Collections.Generic.List[string]
    $items = ConvertTo-ObjectArray -Value $Plan
    if ($items.Count -eq 0) {
        $lines.Add('No applicable fixes for the current selection (Healthy rows are skipped).')
        return ($lines -join [Environment]::NewLine)
    }
    $run = @($items | Where-Object { -not $_.ManualOnly -and $_.Remotable })
    $manual = @($items | Where-Object { $_.ManualOnly })
    $notRemote = @($items | Where-Object { -not $_.ManualOnly -and -not $_.Remotable })
    $devs = @{}
    foreach ($p in $items) { $devs[[string]$p.DeviceName] = $true }
    $crit = @($items | Where-Object { (Get-FixSeverityRank -Severity "$($_.Severity)") -ge 4 }).Count
    $high = @($items | Where-Object { (Get-FixSeverityRank -Severity "$($_.Severity)") -eq 3 }).Count
    $lines.Add("Plan: $($items.Count) solution(s) on $($devs.Count) host(s) | remote=$($run.Count) | manual=$($manual.Count) | skipped-name=$($notRemote.Count)")
    $lines.Add("Severity: Critical=$crit High=$high")
    $lines.Add('')
    foreach ($p in $items) {
        $sev = ('{0,-8}' -f "$($p.Severity)")
        if ($p.ManualOnly) {
            $lines.Add(("[MANUAL ] {0,-16} {1,-8} {2}  ({3})" -f $p.DeviceName, $sev.Trim(), $p.Title, $p.SolutionId))
        } elseif (-not $p.Remotable) {
            $lines.Add(("[SKIP   ] {0,-16} {1,-8} {2}  (not a valid host name)" -f $p.DeviceName, $sev.Trim(), $p.Title))
        } else {
            $lines.Add(("[WINRM  ] {0,-16} {1,-8} {2}  ({3})" -f $p.DeviceName, $sev.Trim(), $p.Title, $p.SolutionId))
        }
    }
    return ($lines -join [Environment]::NewLine)
}

function Invoke-RemoteFixPlan {
    <#
      Executes remote (WinRM) parts of the plan. Manual entries are not executed.
      $Credential: optional PSCredential; $null = current user.
      $WhatIf: preview only.
      $Log: scriptblock receiving status strings.
      Returns result objects: DeviceName, SolutionId, Title, Status, Output, DurationSec
    #>
    param(
        $Plan,
        [System.Management.Automation.PSCredential]$Credential = $null,
        [switch]$WhatIf,
        [scriptblock]$Log,
        [int]$TimeoutSec = 60
    )
    $note = { param($m) if ($Log) { & $Log $m } }
    $results = New-Object System.Collections.Generic.List[object]
    $items = ConvertTo-ObjectArray -Value $Plan
    if ($items.Count -eq 0) { return @() }

    $byDevice = @{}
    foreach ($p in $items) {
        if ($p.ManualOnly) {
            $results.Add([pscustomobject]@{
                DeviceName=$p.DeviceName; SolutionId=$p.SolutionId; Title=$p.Title
                Status='Manual'; Output='Manual - open BigFix console / server runbook step; no WinRM script.'
                DurationSec=0
            })
            & $note ("MANUAL {0} :: {1}" -f $p.DeviceName, $p.Title)
            continue
        }
        if (-not $p.Remotable) {
            $results.Add([pscustomobject]@{
                DeviceName=$p.DeviceName; SolutionId=$p.SolutionId; Title=$p.Title
                Status='Skipped'; Output='Device name not valid for WinRM (action id / invalid host).'
                DurationSec=0
            })
            & $note ("SKIP {0} :: {1}" -f $p.DeviceName, $p.Title)
            continue
        }
        if (-not $byDevice.ContainsKey([string]$p.DeviceName)) { $byDevice[[string]$p.DeviceName] = New-Object System.Collections.Generic.List[object] }
        [void]$byDevice[[string]$p.DeviceName].Add($p)
    }

    $hostTotal = @($byDevice.Keys).Count
    $hostIdx = 0
    foreach ($dev in ($byDevice.Keys | Sort-Object)) {
        if ($script:FixStopRequested) {
            foreach ($d in ($byDevice.Keys | Where-Object { $_ -ne $dev })) {
                foreach ($j in @($byDevice[[string]$d])) {
                    $results.Add([pscustomobject]@{
                        DeviceName=$d; SolutionId=$j.SolutionId; Title=$j.Title
                        Status='Skipped'; Output='Stopped by user before host ran.'
                        DurationSec=0
                    })
                }
            }
            & $note 'Stopped by user - remaining hosts marked Skipped.'
            break
        }
        $hostIdx++
        $jobs = @($byDevice[[string]$dev].ToArray())
        & $note ("[{0}/{1}] Connecting WinRM -> {2} ({3} script(s))..." -f $hostIdx, $hostTotal, $dev, $jobs.Count)
        if ($WhatIf) {
            foreach ($j in $jobs) {
                $results.Add([pscustomobject]@{
                    DeviceName=$dev; SolutionId=$j.SolutionId; Title=$j.Title
                    Status='WhatIf'; Output='WhatIf - not executed.'; DurationSec=0
                })
                & $note ("WHATIF {0} :: {1}" -f $dev, $j.Title)
            }
            continue
        }

        $payload = @($jobs | ForEach-Object {
            [pscustomobject]@{ SolutionId = $_.SolutionId; Title = $_.Title; Script = $_.Script }
        })

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $icmParams = @{
            ComputerName = $dev
            ErrorAction  = 'Stop'
            ArgumentList = @(,$payload)
            ScriptBlock  = {
                param($items)
                $out = @()
                foreach ($it in $items) {
                    try {
                        $sb = [scriptblock]::Create($it.Script)
                        $r = & $sb
                        $ok = $true
                        $text = ''
                        if ($r -is [hashtable]) {
                            $ok = [bool]$r.Ok
                            $text = "$($r.Out)"
                        } elseif ($null -ne $r) {
                            $text = "$r"
                        } else {
                            $text = 'Completed (no output)'
                        }
                        $out += [pscustomobject]@{
                            SolutionId=$it.SolutionId; Title=$it.Title
                            Status=if ($ok) { 'OK' } else { 'Failed' }
                            Output=$text
                        }
                    } catch {
                        $out += [pscustomobject]@{
                            SolutionId=$it.SolutionId; Title=$it.Title
                            Status='Failed'; Output=$_.Exception.Message
                        }
                    }
                }
                return $out
            }
        }
        if ($Credential) { $icmParams.Credential = $Credential }
        if ($TimeoutSec -gt 0) { $icmParams.SessionOption = (New-PSSessionOption -OperationTimeout ($TimeoutSec * 1000) -OpenTimeout ($TimeoutSec * 1000)) }

        try {
            $remoteOut = Invoke-Command @icmParams
            $sw.Stop()
            $dur = [math]::Round($sw.Elapsed.TotalSeconds, 1)
            foreach ($o in @($remoteOut)) {
                if (-not $o) { continue }
                $results.Add([pscustomobject]@{
                    DeviceName=$dev; SolutionId=[string]$o.SolutionId; Title=[string]$o.Title
                    Status=[string]$o.Status; Output=[string]$o.Output; DurationSec=$dur
                })
                & $note ("{0} {1} :: {2} ({3}s)" -f $o.Status, $dev, $o.Title, $dur)
            }
            if (-not $remoteOut) {
                foreach ($j in $jobs) {
                    $results.Add([pscustomobject]@{
                        DeviceName=$dev; SolutionId=$j.SolutionId; Title=$j.Title
                        Status='OK'; Output='Invoke-Command returned no output (treated as success).'
                        DurationSec=$dur
                    })
                }
            }
        } catch {
            $sw.Stop()
            $dur = [math]::Round($sw.Elapsed.TotalSeconds, 1)
            $msg = $_.Exception.Message
            foreach ($j in $jobs) {
                $results.Add([pscustomobject]@{
                    DeviceName=$dev; SolutionId=$j.SolutionId; Title=$j.Title
                    Status='WinRM-Failed'; Output=$msg; DurationSec=$dur
                })
            }
            & $note ("WinRM-Failed {0}: {1} ({2}s)" -f $dev, $msg, $dur)
        }
    }
    return @($results.ToArray())
}

function Test-FixPreFlight {
    <#
      Tests WinRM (Test-WSMan) for every unique remotable host in the plan.
      Returns: DeviceName, Status (OK/Failed/Skipped), Detail
    #>
    param(
        $Plan,
        [System.Management.Automation.PSCredential]$Credential = $null,
        [int]$TimeoutSec = 5,
        [scriptblock]$Log
    )
    $note = { param($m) if ($Log) { & $Log $m } }
    $items = ConvertTo-ObjectArray -Value $Plan
    $hosts = @{}
    foreach ($p in $items) {
        if ($p.ManualOnly) { continue }
        if (-not $p.Remotable) { continue }
        $hosts[[string]$p.DeviceName] = $true
    }
    $names = @($hosts.Keys | Sort-Object)
    if ($names.Count -eq 0) {
        & $note 'Pre-flight: no remotable hosts in plan.'
        return @()
    }
    & $note ("Pre-flight: testing WinRM on {0} host(s) (timeout {1}s)..." -f $names.Count, $TimeoutSec)
    $out = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($h in $names) {
        if ($script:FixStopRequested) {
            & $note 'Pre-flight stopped by user.'
            break
        }
        $i++
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $tp = @{ ComputerName = $h; ErrorAction = 'Stop' }
            if ($Credential) { $tp.Credential = $Credential }
            if ($TimeoutSec -gt 0) { $tp.OperationTimeoutSec = $TimeoutSec }
            $null = Test-WSMan @tp
            $sw.Stop()
            $detail = "WSMan OK ($([math]::Round($sw.Elapsed.TotalSeconds,1))s)"
            $st = 'OK'
        } catch {
            $sw.Stop()
            $detail = $_.Exception.Message
            $st = 'Failed'
        }
        $out.Add([pscustomobject]@{ DeviceName = $h; Status = $st; Detail = $detail; DurationSec = [math]::Round($sw.Elapsed.TotalSeconds, 1) })
        $col = if ($st -eq 'OK') { 'LightGreen' } else { 'Salmon' }
        & $note ("[{0}/{1}] {2} {3} - {4}" -f $i, $names.Count, $st, $h, $detail)
        [void]$col
    }
    $okN = @($out | Where-Object { $_.Status -eq 'OK' }).Count
    $failN = @($out | Where-Object { $_.Status -ne 'OK' }).Count
    & $note ("Pre-flight done: OK={0} Failed={1}" -f $okN, $failN)
    return @($out.ToArray())
}

function Save-FixRunHistory {
    param(
        [System.Collections.IEnumerable]$Results,
        [string]$Mode = 'Execute'
    )
    try {
        $rows = ConvertTo-ObjectArray -Value $Results
        if ($rows.Count -eq 0) { return $null }
        $dir = Join-Path $PSScriptRoot 'reports\history'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $path = Join-Path $dir ("fix-run-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
        $csv = foreach ($r in $rows) {
            [pscustomobject]@{
                Timestamp  = (Get-Date -Format 'o')
                Mode       = $Mode
                DeviceName = "$($r.DeviceName)"
                SolutionId = "$($r.SolutionId)"
                Title      = "$($r.Title)"
                Status     = "$($r.Status)"
                DurationSec= "$($r.DurationSec)"
                Output     = "$($r.Output)"
            }
        }
        $csv | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
        return $path
    } catch {
        return $null
    }
}

function Show-RunbookDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "BigFix Server & Client Runbook v$script:AppVersion"
    $dlg.Size = New-Object System.Drawing.Size(1040, 720)
    $dlg.StartPosition = 'CenterParent'
    $dlg.MinimumSize = New-Object System.Drawing.Size(760, 500)
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $dlg.ShowInTaskbar = $false

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text = "Day-to-day maintenance & troubleshooting - $($script:RunbookTopics.Count) topics (HCL BigFix docs, KBs & community)"
    $lblHead.Dock = 'Top'
    $lblHead.Height = 32
    $lblHead.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $lblHead.ForeColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
    $lblHead.Padding = New-Object System.Windows.Forms.Padding(10, 8, 0, 0)
    $dlg.Controls.Add($lblHead)

    # Search bar
    $pnlSearch = New-Object System.Windows.Forms.Panel
    $pnlSearch.Dock = 'Top'
    $pnlSearch.Height = 34
    $pnlSearch.BackColor = [System.Drawing.Color]::FromArgb(236, 240, 244)
    $dlg.Controls.Add($pnlSearch)
    $lblS = New-Object System.Windows.Forms.Label
    $lblS.Text = 'Search:'
    $lblS.Location = New-Object System.Drawing.Point(10, 8)
    $lblS.Size = New-Object System.Drawing.Size(50, 18)
    $pnlSearch.Controls.Add($lblS)
    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(62, 5)
    $txtSearch.Size = New-Object System.Drawing.Size(360, 22)
    $pnlSearch.Controls.Add($txtSearch)
    $lblHits = New-Object System.Windows.Forms.Label
    $lblHits.Text = "$($script:RunbookTopics.Count) topics"
    $lblHits.Location = New-Object System.Drawing.Point(432, 8)
    $lblHits.Size = New-Object System.Drawing.Size(200, 18)
    $lblHits.ForeColor = [System.Drawing.Color]::FromArgb(90, 100, 110)
    $pnlSearch.Controls.Add($lblHits)

    $pnlBody = New-Object System.Windows.Forms.SplitContainer
    $pnlBody.Dock = 'Fill'
    $pnlBody.Orientation = 'Vertical'
    $dlg.Controls.Add($pnlBody)
    $pnlBody.BringToFront()

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Dock = 'Fill'
    $lst.IntegralHeight = $false
    $lst.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
    $lst.BackColor = [System.Drawing.Color]::White
    foreach ($t in @($script:RunbookTopics)) { [void]$lst.Items.Add($t) }
    if ($lst.Items.Count -gt 0) { $lst.SelectedIndex = 0 }
    $pnlBody.Panel1.Controls.Add($lst)
    $lst.DisplayMember = 'Title'

    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Dock = 'Fill'
    $rtb.ReadOnly = $true
    $rtb.WordWrap = $true
    $rtb.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 252)
    $rtb.Font = New-Object System.Drawing.Font('Consolas', 9.5)
    $rtb.BorderStyle = 'None'
    $rtb.Text = "Select a topic on the left or search above.`r`n`r`nCovers: offline clients, relay health, console cache/FillDB, Inventory upload (UUID/MaxArchiveSize), WebUI/DB credentials, server maintenance, action statuses, slow clients, client install errors, and Web Reports performance."
    $pnlBody.Panel2.Controls.Add($rtb)

    $render = {
        if ($lst.SelectedItem -is [hashtable]) {
            $rtb.Text = Get-RunbookTopicText -Topic $lst.SelectedItem
            $rtb.SelectionStart = 0
            $rtb.SelectionLength = 0
        }
    }.GetNewClosure()
    $lst.Add_SelectedIndexChanged($render)

    $applyFilter = {
        $q = ($txtSearch.Text -replace '[\[\]\*#]', '').Trim()
        $prev = if ($lst.SelectedItem) { $lst.SelectedItem.Id } else { $null }
        $lst.BeginUpdate()
        $lst.Items.Clear()
        $shown = 0
        foreach ($t in @($script:RunbookTopics)) {
            $hit = -not $q
            if ($q) {
                $eq = [regex]::Escape($q)
                $blob = ("$($t.Title) $($t.Tags) $($t.Summary) $($t.RootCause) $($t.Steps -join ' ') $($t.Refs -join ' ')")
                if ($blob -match "(?i)$eq") { $hit = $true }
            }
            if ($hit) { [void]$lst.Items.Add($t); $shown++ }
        }
        $lst.EndUpdate()
        $lblHits.Text = "$shown of $($script:RunbookTopics.Count) topics"
        if ($lst.Items.Count -gt 0) {
            $sel = $null
            if ($prev) { foreach ($i in 0..($lst.Items.Count-1)) { if ($lst.Items[$i].Id -eq $prev) { $sel = $i; break } } }
            $lst.SelectedIndex = if ($sel -ne $null) { $sel } else { 0 }
        } else {
            $rtb.Text = "No topics match '$q'."
        }
    }.GetNewClosure()
    $txtSearch.Add_TextChanged($applyFilter)

    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Dock = 'Bottom'
    $pnlBottom.Height = 44
    $pnlBottom.BackColor = [System.Drawing.Color]::FromArgb(236, 240, 244)
    $dlg.Controls.Add($pnlBottom)

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy topic'
    $btnCopy.Location = New-Object System.Drawing.Point(12, 8)
    $btnCopy.Size = New-Object System.Drawing.Size(110, 28)
    $btnCopy.FlatStyle = 'Flat'
    $pnlBottom.Controls.Add($btnCopy)
    $btnCopy.Add_Click({
        if ($lst.SelectedItem -is [hashtable]) {
            try { [System.Windows.Forms.Clipboard]::SetText((Get-RunbookTopicText -Topic $lst.SelectedItem)) } catch { }
        }
    })

    $btnExportAll = New-Object System.Windows.Forms.Button
    $btnExportAll.Text = 'Export all...'
    $btnExportAll.Location = New-Object System.Drawing.Point(130, 8)
    $btnExportAll.Size = New-Object System.Drawing.Size(110, 28)
    $btnExportAll.FlatStyle = 'Flat'
    $pnlBottom.Controls.Add($btnExportAll)
    $btnExportAll.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = 'Export full runbook'
        $sfd.Filter = 'Text files (*.txt)|*.txt|Markdown (*.md)|*.md'
        $sfd.FileName = "BigFix-Runbook-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
        if ($sfd.ShowDialog() -ne 'OK') { return }
        try {
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine("BigFix Day-to-Day Runbook v$script:AppVersion  ($(Get-Date -Format 'yyyy-MM-dd HH:mm'))")
            [void]$sb.AppendLine('=' * 72)
            foreach ($t in @($script:RunbookTopics)) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine((Get-RunbookTopicText -Topic $t))
                [void]$sb.AppendLine('-' * 72)
            }
            [System.IO.File]::WriteAllText($sfd.FileName, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
            Set-Status "Runbook exported: $($sfd.FileName)"
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Export failed', 'OK', 'Error')
        }
    })

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(920, 8)
    $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Anchor = 'Top,Right'
    $btnClose.FlatStyle = 'Flat'
    $pnlBottom.Controls.Add($btnClose)
    $btnClose.Add_Click({ $dlg.Close() })

    $pnlBottom.Add_Resize({
        $btnClose.Left = $pnlBottom.ClientSize.Width - $btnClose.Width - 12
    })

    $dlg.Add_Shown({
        try {
            $w = $pnlBody.ClientSize.Width
            if ($w -le 0) { $w = $dlg.ClientSize.Width }
            if ($w -le 0) { return }

            $d = [Math]::Min([Math]::Max([int]($w * 0.32), 240), [Math]::Max(240, $w - 320))
            $pnlBody.SplitterDistance = $d

            $p1 = [Math]::Min(200, [Math]::Max(100, $d - 120))
            $p2 = [Math]::Min(300, [Math]::Max(120, $w - $d - 80))
            $pnlBody.Panel1MinSize = $p1
            $pnlBody.Panel2MinSize = $p2
        } catch { }
    })

    $null = $dlg.ShowDialog($script:MainForm)
    $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# Self-test (no GUI)
# ---------------------------------------------------------------------------
# session Remote Fix state (not persisted)
$script:LastFixResults = $null
$script:LastFixPlan = $null
$script:FixRunning = $false
$script:FixStopRequested = $false

if ($SelfTest) {
    Write-Host "BigFix Failure Dashboard v$script:AppVersion - SelfTest" -ForegroundColor Cyan
    $demo = Get-DemoFailures
    $sum = Get-FailureSummary $demo
    Write-Host "Demo records : $($sum.Total) across $($sum.Devices) devices"
    Write-Host "  Offline    : $($sum.Offline)"
    Write-Host "  Fixlet     : $($sum.FixletFailures)"
    Write-Host "  Compliance : $($sum.Compliance)"
    Write-Host "  Software   : $($sum.SoftwareDeploy)"
    Write-Host "  Healthy    : $($sum.Healthy)"
    Write-Host "  Critical   : $($sum.Critical) / High: $($sum.High)"

    $sample = Join-Path $script:ScriptRoot 'sample-data\sample-failures.csv'
    if (Test-Path -LiteralPath $sample) {
        $imported = Import-FailureCsv -Path $sample
        Write-Host "Sample CSV  : imported $($imported.Count) rows OK"
    } else {
        Write-Warning "Sample CSV not found at $sample"
    }

    # Root-cause spot checks
    $checks = @(
        @{ Err='Failed to download prefetch block - relay unavailable'; ExpectCat='FixletFailure' },
        @{ Err='Not compliant - BitLocker off'; ExpectCat='ComplianceFailure' },
        @{ Err='Client not reporting - last report more than threshold ago'; ExpectCat='Offline' },
        @{ Err='Action status: Download Failed'; ExpectCat='FixletFailure' },
        @{ Err='Hash Mismatch on download'; ExpectCat='FixletFailure' },
        @{ Err='MaxArchiveSize exceeded - scan not uploaded'; ExpectCat='SoftwareDeployment' },
        @{ Err='Duplicate VM UUID detected'; ExpectCat='SoftwareDeployment' },
        @{ Err='BESClient install failed error 1920'; ExpectCat='FixletFailure' },
        @{ Err='FillDB log stopped - console stale'; ExpectCat='FixletFailure' }
    )
    $pass = 0
    foreach ($c in $checks) {
        $a = Get-FailureAnalysis -ErrorMessage $c.Err
        $ok = ($a.Category -eq $c.ExpectCat)
        if ($ok) { $pass++ } else { Write-Warning "Analysis mismatch for '$($c.Err)' -> got $($a.Category), expected $($c.ExpectCat)" }
    }
    Write-Host "Root-cause checks: $pass/$($checks.Count) passed"

    # Runbook checks
    $rbPass = 0
    $rbTotal = 0
    $rbTotal++; if (@($script:RunbookTopics).Count -ge 10) { $rbPass++ } else { Write-Warning "Runbook has only $(@($script:RunbookTopics).Count) topics (need >=10)" }
    $rbTotal++; $offTopic = @($script:RunbookTopics | Where-Object { $_.Id -eq 'clients-offline' }) | Select-Object -First 1; if ($offTopic -and (($offTopic.Steps -join ' ') -match '52311') -and (($offTopic.Steps -join ' ') -match 'ForceRefresh|forcerefresh')) { $rbPass++ } else { Write-Warning 'Offline topic missing 52311/ForceRefresh steps' }
    $rbTotal++; $invTopic = @($script:RunbookTopics | Where-Object { $_.Id -eq 'inventory-upload' }) | Select-Object -First 1; if ($invTopic -and (($invTopic.Steps -join ' ') -match 'UUID') -and (($invTopic.Steps -join ' ') -match 'MaxArchiveSize|Computer Support Data')) { $rbPass++ } else { Write-Warning 'Inventory topic missing UUID/MaxArchiveSize steps' }
    $rbTotal++; if (@($script:RunbookTopics | Where-Object { $_.Id -eq 'action-status' }) ) { $rbPass++ } else { Write-Warning 'Action-status topic missing' }
    $rbTotal++; if (@($script:RunbookTopics | Where-Object { $_.Id -eq 'client-slow' }) ) { $rbPass++ } else { Write-Warning 'Slow-client topic missing' }
    foreach ($t in @($script:RunbookTopics)) {
        if (-not $t.Steps -or @($t.Steps).Count -lt 3) { Write-Warning "Topic '$($t.Title)' has too few steps"; $rbTotal++ } else { $rbTotal++; $rbPass++ }
        if (-not $t.Refs -or @($t.Refs).Count -lt 1) { Write-Warning "Topic '$($t.Title)' has no refs"; $rbTotal++ } else { $rbTotal++; $rbPass++ }
    }
    Write-Host "Runbook checks: $rbPass/$rbTotal passed"

    # Fix-solution coverage + plan relevance checks
    $fxPass = 0
    $fxTotal = 0
    $fxTotal++; if ($script:FixSolutions.Count -ge 20) { $fxPass++ } else { Write-Warning "FixSolutions count only $($script:FixSolutions.Count)" }
    foreach ($r in @($script:RootCauseRules)) {
        $fxTotal++
        $has = $script:FixSolutions.ContainsKey($r.Id) -or $script:FixCategoryDefaults.ContainsKey($r.Category)
        if ($has) { $fxPass++ } else { Write-Warning "No fix solution for rule $($r.Id) / cat $($r.Category)" }
    }
    $fxTotal++
    $planAll = @(Get-FixPlan -Records $demo)
    if ($planAll.Count -gt 0) { $fxPass++ } else { Write-Warning 'Get-FixPlan returned empty for demo' }
    $fxTotal++
    $planAllOk = $true
    foreach ($p in $planAll) {
        if (-not $p.ManualOnly -and -not $p.Script) { $planAllOk = $false; Write-Warning "Plan missing script for $($p.SolutionId)" }
        if ($p.ManualOnly -and $p.Script) { $planAllOk = $false; Write-Warning "Manual plan should have empty script: $($p.SolutionId)" }
    }
    if ($planAllOk -and $planAll.Count -gt 0) { $fxPass++ }
    # Severity filter: Critical-only plan must be a subset with only Critical rows
    $fxTotal++
    $planCrit = @(Get-FixPlan -Records $demo -MinSeverity 'Critical')
    $planCritOk = $true
    if ($planCrit.Count -gt $planAll.Count) { $planCritOk = $false; Write-Warning 'Critical plan larger than full plan' }
    foreach ($p in $planCrit) {
        if ((Get-FixSeverityRank -Severity "$($p.Severity)") -lt 4) { $planCritOk = $false; Write-Warning "Critical plan has non-Critical $($p.Severity) $($p.SolutionId)" }
    }
    if ($planCrit.Count -ge 0 -and $planCritOk) { $fxPass++ } else { Write-Warning 'MinSeverity Critical filter failed' }
    # Relevant-only: offline records must not pull download/hash/etc solutions
    $fxTotal++
    $offRecs = @($demo | Where-Object { $_.Category -eq 'Offline' })
    $offPlan = @(Get-FixPlan -Records $offRecs)
    $offBad = @($offPlan | Where-Object { $_.SolutionId -notin @('BF-RELAY-OFFLINE', 'BF-ACCESS-DEFAULT', 'BF-NET-PORT', 'BF-DNS-RELAY', 'BF-CACHE-CORRUPT') -and -not $_.ManualOnly })
    if ($offRecs.Count -gt 0 -and $offPlan.Count -gt 0 -and $offBad.Count -eq 0) { $fxPass++ }
    else { Write-Warning "Offline plan not relevant: bad=$($offBad.Count) plan=$($offPlan.Count)" }
    # WhatIf dry-run returns status without needing network when all manual/skip - run full WhatIf on plan
    $fxTotal++
    $whatIfOut = @(Invoke-RemoteFixPlan -Plan $planAll -WhatIf -Log { })
    $okWhatIf = ($whatIfOut.Count -ge @($planAll | Where-Object { -not $_.ManualOnly -and $_.Remotable }).Count)
    if ($okWhatIf) { $fxPass++ } else { Write-Warning "WhatIf results short: $($whatIfOut.Count)" }
    Write-Host "Fix-solution checks: $fxPass/$fxTotal passed"

    # v2 feature checks: exports, cache, device aggregates, settings
    $v2Pass = 0
    $v2Total = 0
    $tmpDir = Join-Path $env:TEMP ("bfd-selftest-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
    try {
        $v2Total++
        $htmlPath = Join-Path $tmpDir 'report.html'
        Export-FailuresHtml -Records $demo -Path $htmlPath -Title 'SelfTest Report'
        $htmlRaw = Get-Content -LiteralPath $htmlPath -Raw
        if ((Test-Path $htmlPath) -and $htmlRaw -match '<table' -and $htmlRaw -match 'ROOT CAUSE|Root Cause|rootCause|Root cause') { $v2Pass++ }
        else { Write-Warning 'HTML export missing table/root-cause content' }

        $v2Total++
        $jsonPath = Join-Path $tmpDir 'report.json'
        Export-FailuresJson -Records $demo -Path $jsonPath
        $jraw = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
        if ($jraw.Count -eq $demo.Count -and $jraw.Records.Count -eq $demo.Count) { $v2Pass++ }
        else { Write-Warning "JSON export count mismatch: $($jraw.Count)" }

        $v2Total++
        $csvPath = Join-Path $tmpDir 'report.csv'
        Export-FailuresCsv -Records $demo -Path $csvPath
        $csvRows = @(Import-Csv -LiteralPath $csvPath)
        if ($csvRows.Count -eq $demo.Count -and $csvRows[0].PSObject.Properties.Name -contains 'MatchedRule') { $v2Pass++ }
        else { Write-Warning "CSV export rows=$($csvRows.Count)" }

        $v2Total++
        $devAgg = @(Get-DeviceAggregates -Records $demo)
        if ($devAgg.Count -gt 0 -and $devAgg.Count -le $demo.Count -and ($devAgg | Where-Object { $_.FailureCount -ge 1 })) { $v2Pass++ }
        else { Write-Warning "Device aggregates bad: $($devAgg.Count)" }

        $v2Total++
        $cacheJson = ConvertTo-ApiResultJson -Result ([pscustomobject]@{
            Failures = $demo; Devices = @(); Log = @('cache-test'); Version = 'selftest'; ElapsedSec = 1
        })
        $cached = ConvertFrom-ApiResultJson -Json $cacheJson
        if ($cached.Failures.Count -eq $demo.Count) { $v2Pass++ } else { Write-Warning "Cache roundtrip failures=$($cached.Failures.Count)" }

        $v2Total++
        $setPath = Join-Path $tmpDir 'settings.json'
        $oldSettingsPath = $script:SettingsPath
        $script:SettingsPath = $setPath
        $script:AppSettings = Get-DefaultAppSettings
        $script:AppSettings.Server = 'https://unit-test:52311'
        $script:AppSettings.AutoRefreshMinutes = 15
        $null = Save-AppSettings
        $script:AppSettings = $null
        $loaded = Get-AppSettings
        $script:SettingsPath = $oldSettingsPath
        $null = Get-AppSettings
        if ($loaded.Server -eq 'https://unit-test:52311' -and [int]$loaded.AutoRefreshMinutes -eq 15) { $v2Pass++ }
        else { Write-Warning 'Settings roundtrip failed' }

        $v2Total++
        if ($script:FixSolutions.Count -ge 30 -and @($script:RootCauseRules).Count -ge 35) { $v2Pass++ }
        else { Write-Warning "Rule/solution counts low: rules=$(@($script:RootCauseRules).Count) solutions=$($script:FixSolutions.Count)" }
    } finally {
        try { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    }
    Write-Host "v2 feature checks: $v2Pass/$v2Total passed"

    if ($pass -eq $checks.Count -and $rbPass -eq $rbTotal -and $fxPass -eq $fxTotal -and $v2Pass -eq $v2Total -and $sum.Total -gt 0) {
        Write-Host "SELFTEST PASSED" -ForegroundColor Green
        exit 0
    } else {
        Write-Host "SELFTEST FAILED" -ForegroundColor Red
        exit 1
    }
}

# ===========================================================================
# ApiWorker (background Connect API) - writes progress + JSON result files
# ===========================================================================
if ($ApiWorker) {
    $ErrorActionPreference = 'Stop'
    $cfgPath = $ApiWorker
    $cfg = $null
    try {
        $cfg = Get-Content -LiteralPath $cfgPath -Raw | ConvertFrom-Json
        $progressFile = [string]$cfg.ProgressFile
        $resultFile = [string]$cfg.ResultFile
        $note = {
            param($m)
            try { Set-Content -LiteralPath $progressFile -Value $m -Encoding UTF8 } catch { }
        }
        & $note "Starting API worker (phase=$($cfg.Phase))..."
        $seed = $null
        if ($cfg.Phase -eq 'Failures' -and $cfg.SeedJson) {
            $seed = ConvertFrom-ApiResultJson -Json ([string]$cfg.SeedJson)
        }
        $params = @{
            BaseUrl             = [string]$cfg.BaseUrl
            Username            = [string]$cfg.Username
            Password            = [string]$cfg.Password
            OfflineThresholdHours = [int]$cfg.OfflineThresholdHours
            Progress            = $note
        }
        if ($cfg.SkipCertCheck) { $params.SkipCertCheck = $true }
        if ($cfg.Phase -eq 'Inventory') { $params.InventoryOnly = $true }
        if ($cfg.Phase -eq 'Failures') {
            $params.FailuresOnly = $true
            $params.SeedData = $seed
        }
        $result = Get-BigFixApiData @params
        $json = ConvertTo-ApiResultJson -Result $result
        Set-Content -LiteralPath $resultFile -Value $json -Encoding UTF8
        & $note "Worker complete."
        exit 0
    } catch {
        $errJson = ([pscustomobject]@{ Error = "$($_.Exception.Message)"; ResultFile = $(if ($cfg) { [string]$cfg.ResultFile } else { '' }) }) | ConvertTo-Json -Depth 4
        try {
            if ($cfg -and $cfg.ResultFile) { Set-Content -LiteralPath ([string]$cfg.ResultFile) -Value $errJson -Encoding UTF8 }
        } catch { }
        try { if ($cfg -and $cfg.ProgressFile) { Set-Content -LiteralPath ([string]$cfg.ProgressFile) -Value "ERROR: $($_.Exception.Message)" -Encoding UTF8 } } catch { }
        exit 1
    } finally {
        try { if ($cfgPath -and (Test-Path -LiteralPath $cfgPath)) { Remove-Item -LiteralPath $cfgPath -Force -ErrorAction SilentlyContinue } } catch { }
    }
}

# ===========================================================================
# Headless export / email mode (no GUI)
# ===========================================================================
$script:IsHeadless = ($ExportCsv -or $ExportHtml -or $ExportJson -or $Email -or $ImportPath -or ($UseCache -and -not $SelfTest))
if ($script:IsHeadless -and -not $SelfTest -and -not $ApiWorker) {
    $ErrorActionPreference = 'Stop'
    try {
        $records = $null
        $sourceLabel = ''
        if ($ImportPath) {
            $ext = [System.IO.Path]::GetExtension($ImportPath).ToLower()
            if ($ext -eq '.xml') { $records = Import-FailureXml -Path $ImportPath } else { $records = Import-FailureCsv -Path $ImportPath }
            $sourceLabel = "Import: $ImportPath"
        } elseif ($UseCache -or -not $Password) {
            $cached = Get-SessionCache
            if ($cached) {
                $records = ConvertTo-ObjectArray -Value $cached.Failures
                $sourceLabel = "Cache: $($cached.SavedAt)"
            } elseif (-not $ImportPath) {
                if ($UseCache) { throw "No session cache found at $script:CachePath. Connect API once in GUI, or pass -ImportPath." }
                # No password and no cache: fall back to demo for report smoke paths
                $records = Get-DemoFailures
                $sourceLabel = 'Demo data (no password/cache)'
            }
        }
        if (-not $records -and $Password) {
            $baseUrl = if ($Server) { Get-BigFixBaseUri -Url $Server } else { throw 'Provide -Server with password (env BIGFIX_DASH_PASSWORD or -Password), -ImportPath, or -UseCache.' }
            $user = if ($Username) { $Username } else { throw '-Username is required for API headless mode.' }
            $result = Get-BigFixApiData -BaseUrl $baseUrl -Username $user -Password $Password -OfflineThresholdHours 4 -Progress { param($m) Write-Host $m }
            $records = ConvertTo-ObjectArray -Value $result.Failures
            $sourceLabel = "BigFix API ($baseUrl)"
        }
        $records = ConvertTo-ObjectArray -Value $records
        if (-not $records -or $records.Count -eq 0) { throw 'No records loaded for headless export.' }
        Write-Host "Headless source: $sourceLabel ($($records.Count) rows)"

        if ($ExportCsv) {
            Export-FailuresCsv -Records $records -Path $ExportCsv
            Write-Host "CSV  -> $ExportCsv"
        }
        if ($ExportHtml) {
            Export-FailuresHtml -Records $records -Path $ExportHtml -Title 'BigFix Failed Device Report'
            Write-Host "HTML -> $ExportHtml"
        }
        if ($ExportJson) {
            Export-FailuresJson -Records $records -Path $ExportJson
            Write-Host "JSON -> $ExportJson"
        }
        if ($Email) {
            $smtp = $script:AppSettings.Smtp
            Send-FailureDigest -Records $records -Smtp $smtp
            Write-Host 'Email digest sent.'
        }
        exit 0
    } catch {
        Write-Host "HEADLESS FAILED: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# ===========================================================================
# GUI
# ===========================================================================
$script:AllRecords = New-Object System.Collections.Generic.List[object]
$script:FilteredRecords = New-Object System.Collections.Generic.List[object]
$script:OfflineThresholdHours = 4
$script:LastSource = 'None'
$script:DataTable = $null
$script:DataView = $null

# ---- Form ----
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Failed Device Dashboard v$script:AppVersion"
$form.Size = New-Object System.Drawing.Size(1360, 860)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1100, 700)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# ---- Top connection panel ----
$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Dock = 'Top'
$pnlTop.Height = 92
$pnlTop.BackColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
$form.Controls.Add($pnlTop)

function New-Label {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 80, [string]$Color = 'White')
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, 18)
    $l.ForeColor = [System.Drawing.Color]::FromName($Color)
    $Parent.Controls.Add($l)
    return $l
}
function New-TextBox {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 220, [switch]$Password)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Text = $Text; $t.Location = New-Object System.Drawing.Point($X, $Y)
    $t.Size = New-Object System.Drawing.Size($W, 24)
    if ($Password) { $t.UseSystemPasswordChar = $true }
    $Parent.Controls.Add($t)
    return $t
}
function New-Button {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 110, [int]$H = 28, [string]$Tag = '')
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H)
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 210, 220)
    if ($Tag) { $b.Name = $Tag }
    $Parent.Controls.Add($b)
    return $b
}

$lblTitle = New-Label $pnlTop 'BigFix Failed Device Dashboard' 14 8 360 'White'
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)

$null = New-Label $pnlTop 'Server:' 14 42 50
$txtServer = New-TextBox $pnlTop $(if ($Server) { $Server } else { 'https://localhost:52311' }) 68 39 250

$null = New-Label $pnlTop 'User:' 328 42 36
$txtUser = New-TextBox $pnlTop $Username 366 39 120

$null = New-Label $pnlTop 'Password:' 496 42 60
$txtPass = New-TextBox $pnlTop '' 560 39 120 -Password

$uiTip = New-Object System.Windows.Forms.ToolTip
$uiTip.SetToolTip($txtServer, 'https://bigfix:52311  (you may paste .../api/help - path is stripped automatically)')
$uiTip.SetToolTip($txtUser, 'BigFix console operator with "Can use REST API". Local: apiuser | LDAP: DOMAIN\user or user@domain')
$uiTip.SetToolTip($txtPass, 'Password for the BigFix console operator (not saved to disk)')

$null = New-Label $pnlTop 'Offline >' 696 42 60
$numOffline = New-Object System.Windows.Forms.NumericUpDown
$numOffline.Location = New-Object System.Drawing.Point(758, 39)
$numOffline.Size = New-Object System.Drawing.Size(56, 24)
$numOffline.Minimum = 1; $numOffline.Maximum = 720; $numOffline.Value = 4
$numOffline.TextAlign = 'Right'
$pnlTop.Controls.Add($numOffline)
$null = New-Label $pnlTop 'hrs' 818 42 30

$chkSkipCert = New-Object System.Windows.Forms.CheckBox
$chkSkipCert.Text = 'Skip TLS cert check'
$chkSkipCert.Location = New-Object System.Drawing.Point(854, 40)
$chkSkipCert.Size = New-Object System.Drawing.Size(140, 22)
$chkSkipCert.ForeColor = [System.Drawing.Color]::White
$chkSkipCert.Checked = $false
$pnlTop.Controls.Add($chkSkipCert)

$null = New-Label $pnlTop 'Auto-refresh' 1010 42 80
$numAutoRefresh = New-Object System.Windows.Forms.NumericUpDown
$numAutoRefresh.Location = New-Object System.Drawing.Point(1095, 39)
$numAutoRefresh.Size = New-Object System.Drawing.Size(56, 24)
$numAutoRefresh.Minimum = 0; $numAutoRefresh.Maximum = 1440; $numAutoRefresh.Value = 0
$numAutoRefresh.TextAlign = 'Right'
$pnlTop.Controls.Add($numAutoRefresh)
$null = New-Label $pnlTop 'min (0=off)' 1155 42 80

$btnConnect = New-Button $pnlTop 'Connect API' 14 62 100 24
$btnImport  = New-Button $pnlTop 'Import...' 120 62 90 24
$btnDemo    = New-Button $pnlTop 'Demo' 216 62 70 24
$btnRefresh = New-Button $pnlTop 'Refresh' 292 62 80 24
$btnExport  = New-Button $pnlTop 'CSV...' 378 62 70 24
$btnHtml    = New-Button $pnlTop 'HTML...' 454 62 70 24
$btnEmail   = New-Button $pnlTop 'Email' 530 62 70 24
$btnTicket  = New-Button $pnlTop 'Ticket' 606 62 70 24
$btnRunbook = New-Button $pnlTop 'Runbook' 682 62 80 24
$btnSettings= New-Button $pnlTop 'Settings' 768 62 80 24
$btnAbout   = New-Button $pnlTop 'About' 854 62 70 24
$script:MainForm = $form

$allToolbarBtns = @($btnConnect, $btnImport, $btnDemo, $btnRefresh, $btnExport, $btnHtml, $btnEmail, $btnTicket, $btnRunbook, $btnSettings, $btnAbout)
foreach ($b in $allToolbarBtns) {
    $b.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 252)
    $b.ForeColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
}
$uiTip.SetToolTip($btnRunbook, 'Day-to-day server/client troubleshooting runbook (HCL Maintenance & Troubleshooting)')
$uiTip.SetToolTip($btnHtml, 'Export executive HTML report (filtered rows)')
$uiTip.SetToolTip($btnEmail, 'Send SMTP digest of current records (configure in Settings)')
$uiTip.SetToolTip($btnTicket, 'POST selected/filtered failures to ticket webhook (configure in Settings)')
$uiTip.SetToolTip($btnSettings, 'Persisted settings: filters, auto-refresh, SMTP, ticket webhook')
$uiTip.SetToolTip($numAutoRefresh, 'Minutes between automatic refreshes (0 disables)')

# ---- Summary tiles ----
$pnlSummary = New-Object System.Windows.Forms.Panel
$pnlSummary.Dock = 'Top'
$pnlSummary.Height = 64
$pnlSummary.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($pnlSummary)

function New-Tile {
    param($Parent, [string]$Title, [int]$X, [string]$ValueColor)
    $p = New-Object System.Windows.Forms.Panel
    $p.Location = New-Object System.Drawing.Point($X, 8)
    $p.Size = New-Object System.Drawing.Size(150, 48)
    $p.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 247)
    $lblV = New-Object System.Windows.Forms.Label
    $lblV.Text = '0'
    $lblV.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
    $lblV.ForeColor = [System.Drawing.Color]::FromName($ValueColor)
    $lblV.Location = New-Object System.Drawing.Point(8, 4)
    $lblV.Size = New-Object System.Drawing.Size(134, 24)
    $lblV.TextAlign = 'MiddleCenter'
    $lblT = New-Object System.Windows.Forms.Label
    $lblT.Text = $Title
    $lblT.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lblT.ForeColor = [System.Drawing.Color]::FromArgb(90, 100, 110)
    $lblT.Location = New-Object System.Drawing.Point(8, 28)
    $lblT.Size = New-Object System.Drawing.Size(134, 16)
    $lblT.TextAlign = 'MiddleCenter'
    $p.Controls.Add($lblV); $p.Controls.Add($lblT)
    $Parent.Controls.Add($p)
    return $lblV
}

$tileTotal  = New-Tile $pnlSummary 'Failed Records' 12 'DarkBlue'
$tileDev    = New-Tile $pnlSummary 'Devices' 172 'DarkBlue'
$tileOff    = New-Tile $pnlSummary 'Offline' 332 'DarkRed'
$tileFix    = New-Tile $pnlSummary 'Fixlet Failures' 492 'DarkOrange'
$tileComp   = New-Tile $pnlSummary 'Compliance' 652 'DarkViolet'
$tileSw     = New-Tile $pnlSummary 'Software Deploy' 812 'Teal'
$tileCrit   = New-Tile $pnlSummary 'Critical / High' 972 'Red'
$tileHealthy= New-Tile $pnlSummary 'Healthy' 1132 'ForestGreen'

# ---- Filter bar ----
$pnlFilter = New-Object System.Windows.Forms.Panel
$pnlFilter.Dock = 'Top'
$pnlFilter.Height = 40
$pnlFilter.BackColor = [System.Drawing.Color]::FromArgb(236, 240, 244)
$form.Controls.Add($pnlFilter)

$null = New-Label $pnlFilter 'Category:' 12 11 60 'Black'
$cmbCategory = New-Object System.Windows.Forms.ComboBox
$cmbCategory.DropDownStyle = 'DropDownList'
$cmbCategory.Location = New-Object System.Drawing.Point(74, 8)
$cmbCategory.Size = New-Object System.Drawing.Size(160, 24)
[void]$cmbCategory.Items.AddRange(@('All', 'Offline', 'FixletFailure', 'ComplianceFailure', 'SoftwareDeployment', 'Healthy'))
$cmbCategory.SelectedItem = 'All'
$pnlFilter.Controls.Add($cmbCategory)

$null = New-Label $pnlFilter 'Severity:' 248 11 56 'Black'
$cmbSeverity = New-Object System.Windows.Forms.ComboBox
$cmbSeverity.DropDownStyle = 'DropDownList'
$cmbSeverity.Location = New-Object System.Drawing.Point(308, 8)
$cmbSeverity.Size = New-Object System.Drawing.Size(110, 24)
[void]$cmbSeverity.Items.AddRange(@('All', 'Critical', 'High', 'Medium', 'Low'))
$cmbSeverity.SelectedItem = 'All'
$pnlFilter.Controls.Add($cmbSeverity)

$null = New-Label $pnlFilter 'Search:' 436 11 50 'Black'
$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(490, 8)
$txtSearch.Size = New-Object System.Drawing.Size(260, 24)
$pnlFilter.Controls.Add($txtSearch)

$btnClearFilter = New-Button $pnlFilter 'Clear Filters' 760 6 100 26
$btnClearFilter.BackColor = [System.Drawing.Color]::White
$btnClearFilter.ForeColor = [System.Drawing.Color]::FromArgb(30, 56, 92)

$chkDeviceView = New-Object System.Windows.Forms.CheckBox
$chkDeviceView.Text = 'Device view'
$chkDeviceView.Location = New-Object System.Drawing.Point(870, 10)
$chkDeviceView.Size = New-Object System.Drawing.Size(100, 22)
$chkDeviceView.Checked = [bool]$script:AppSettings.DeviceGroupView
$pnlFilter.Controls.Add($chkDeviceView)
$uiTip.SetToolTip($chkDeviceView, 'Group rows by device (one row per host; detail pane shows full failure timeline)')

$lblFilterCount = New-Label $pnlFilter '' 980 11 200 'Black'
$lblFilterCount.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Italic)

# ---- Main tab control (Dashboard + Charts + last Remote Fix tab) ----
$tabMain = New-Object System.Windows.Forms.TabControl
$tabMain.Dock = 'Fill'
$tabMain.Padding = New-Object System.Drawing.Point(10, 6)
$tabMain.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.Controls.Add($tabMain)

$tabDashboard = New-Object System.Windows.Forms.TabPage
$tabDashboard.Text = 'Dashboard'
$tabDashboard.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$tabMain.TabPages.Add($tabDashboard)

$tabCharts = New-Object System.Windows.Forms.TabPage
$tabCharts.Text = 'Charts'
$tabCharts.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$tabMain.TabPages.Add($tabCharts)

$tabRemoteFix = New-Object System.Windows.Forms.TabPage
$tabRemoteFix.Text = 'Remote Fix'
$tabRemoteFix.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$tabMain.TabPages.Add($tabRemoteFix)   # last tab - Fix button lives here
$tabMain.BringToFront()

# ---- Charts tab content ----
$script:ChartsAvailable = $false
try {
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction Stop
    $script:ChartsAvailable = $true
} catch { }

if ($script:ChartsAvailable) {
    $splitCharts = New-Object System.Windows.Forms.SplitContainer
    $splitCharts.Dock = 'Fill'
    $splitCharts.Orientation = 'Vertical'
    $splitCharts.SplitterDistance = [int]($tabCharts.Width / 2)
    $tabCharts.Controls.Add($splitCharts)
    $splitCharts.BringToFront()

    function New-DashChart {
        param($Parent, [string]$ChartTitle, [string]$ChartType)
        $ch = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
        $ch.Dock = 'Fill'
        $ch.BackColor = [System.Drawing.Color]::White
        $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
        $area.Name = 'Main'
        [void]$ch.ChartAreas.Add($area)
        $chTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
        $chTitle.Text = [string]$ChartTitle
        $chTitle.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $chTitle.ForeColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
        [void]$ch.Titles.Add($chTitle)
        if ($ch.Legends.Count -gt 0) {
            $ch.Legends[0].Enabled = $true
            $ch.Legends[0].Docking = 'Bottom'
        }
        $Parent.Controls.Add($ch)
        return $ch
    }

    $rowTop = New-Object System.Windows.Forms.SplitContainer
    $rowTop.Dock = 'Fill'
    $rowTop.Orientation = 'Vertical'
    $splitCharts.Panel1.Controls.Add($rowTop)
    $rowTop.BringToFront()
    $rowBot = New-Object System.Windows.Forms.SplitContainer
    $rowBot.Dock = 'Fill'
    $rowBot.Orientation = 'Vertical'
    $splitCharts.Panel2.Controls.Add($rowBot)
    $rowBot.BringToFront()

    $chartCategory = New-DashChart $rowTop.Panel1 'Failures by category' 'Pie'
    $chartSeverity = New-DashChart $rowTop.Panel2 'Failures by severity' 'Column'
    $chartTopDev    = New-DashChart $rowBot.Panel1 'Top devices by failure count' 'Bar'
    $chartRules     = New-DashChart $rowBot.Panel2 'Top matched rules' 'Bar'

    $script:ChartCategory = $chartCategory
    $script:ChartSeverity = $chartSeverity
    $script:ChartTopDev = $chartTopDev
    $script:ChartRules = $chartRules
} else {
    $lblNoCharts = New-Object System.Windows.Forms.Label
    $lblNoCharts.Text = 'Charting assembly (System.Windows.Forms.DataVisualization) is not available on this system.'
    $lblNoCharts.Dock = 'Fill'
    $lblNoCharts.TextAlign = 'MiddleCenter'
    $tabCharts.Controls.Add($lblNoCharts)
}

# ---- Main split: grid + details (Dashboard tab) ----
$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock = 'Fill'
$splitMain.Orientation = 'Horizontal'
$splitMain.Panel1MinSize = 180
$tabDashboard.Controls.Add($splitMain)
$splitMain.BringToFront()

# Grid
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToOrderColumns = $true
$grid.ReadOnly = $true
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = 'None'
$grid.RowHeadersVisible = $false
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = 'Fixed3D'
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)
# Large-grid optimizations
$grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$grid.ColumnHeadersHeight = 30
$grid.RowTemplate.Height = 22
$grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
$grid.VirtualMode = $false
$splitMain.Panel1.Controls.Add($grid)

# Details (right/bottom)
$pnlDetails = New-Object System.Windows.Forms.Panel
$pnlDetails.Dock = 'Fill'
$pnlDetails.BackColor = [System.Drawing.Color]::White
$pnlDetails.Padding = New-Object System.Windows.Forms.Padding(8)
$splitMain.Panel2.Controls.Add($pnlDetails)

$lblDetailsTitle = New-Object System.Windows.Forms.Label
$lblDetailsTitle.Text = 'Device details / root cause'
$lblDetailsTitle.Dock = 'Top'
$lblDetailsTitle.Height = 28
$lblDetailsTitle.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$lblDetailsTitle.ForeColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
$lblDetailsTitle.Padding = New-Object System.Windows.Forms.Padding(4, 6, 0, 0)
$pnlDetails.Controls.Add($lblDetailsTitle)

$rtbDetails = New-Object System.Windows.Forms.RichTextBox
$rtbDetails.Dock = 'Fill'
$rtbDetails.ReadOnly = $true
$rtbDetails.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 252)
$rtbDetails.Font = New-Object System.Drawing.Font('Consolas', 9.5)
$rtbDetails.WordWrap = $true
$rtbDetails.BorderStyle = 'None'
$pnlDetails.Controls.Add($rtbDetails)
$rtbDetails.BringToFront()

# ---- Remote Fix tab UI (last tab) ----
$pnlFixTop = New-Object System.Windows.Forms.Panel
$pnlFixTop.Dock = 'Top'
$pnlFixTop.Height = 140
$pnlFixTop.BackColor = [System.Drawing.Color]::White
$tabRemoteFix.Controls.Add($pnlFixTop)

$null = New-Label $pnlFixTop 'Scope:' 12 10 50 'Black'
$radScopeSel = New-Object System.Windows.Forms.RadioButton
$radScopeSel.Text = 'Selected row'
$radScopeSel.Location = New-Object System.Drawing.Point(64, 8)
$radScopeSel.Size = New-Object System.Drawing.Size(120, 22)
$radScopeSel.Checked = $true
$pnlFixTop.Controls.Add($radScopeSel)

$radScopeFiltered = New-Object System.Windows.Forms.RadioButton
$radScopeFiltered.Text = 'Filtered rows'
$radScopeFiltered.Location = New-Object System.Drawing.Point(190, 8)
$radScopeFiltered.Size = New-Object System.Drawing.Size(130, 22)
$pnlFixTop.Controls.Add($radScopeFiltered)

$radScopeAll = New-Object System.Windows.Forms.RadioButton
$radScopeAll.Text = 'All failed (not Healthy)'
$radScopeAll.Location = New-Object System.Drawing.Point(326, 8)
$radScopeAll.Size = New-Object System.Drawing.Size(180, 22)
$pnlFixTop.Controls.Add($radScopeAll)

$null = New-Label $pnlFixTop 'Severity:' 520 10 60 'Black'
$cmbFixSeverity = New-Object System.Windows.Forms.ComboBox
$cmbFixSeverity.DropDownStyle = 'DropDownList'
$cmbFixSeverity.Location = New-Object System.Drawing.Point(584, 8)
$cmbFixSeverity.Size = New-Object System.Drawing.Size(140, 24)
foreach ($sv in @('All severities', 'Critical only', 'Critical + High')) { [void]$cmbFixSeverity.Items.Add($sv) }
$cmbFixSeverity.SelectedIndex = 0
$pnlFixTop.Controls.Add($cmbFixSeverity)

$null = New-Label $pnlFixTop 'WinRM user:' 12 42 80 'Black'
$txtFixUser = New-TextBox $pnlFixTop '' 96 39 180
$null = New-Label $pnlFixTop 'Password:' 286 42 66 'Black'
$txtFixPass = New-TextBox $pnlFixTop '' 356 39 150 -Password

$chkFixCurrentCred = New-Object System.Windows.Forms.CheckBox
$chkFixCurrentCred.Text = 'Use current Windows credentials'
$chkFixCurrentCred.Location = New-Object System.Drawing.Point(520, 40)
$chkFixCurrentCred.Size = New-Object System.Drawing.Size(210, 22)
$chkFixCurrentCred.Checked = $true
$pnlFixTop.Controls.Add($chkFixCurrentCred)

$chkFixWhatIf = New-Object System.Windows.Forms.CheckBox
$chkFixWhatIf.Text = 'Dry run (do not execute)'
$chkFixWhatIf.Location = New-Object System.Drawing.Point(740, 40)
$chkFixWhatIf.Size = New-Object System.Drawing.Size(170, 22)
$pnlFixTop.Controls.Add($chkFixWhatIf)

$null = New-Label $pnlFixTop 'Timeout:' 740 10 55 'Black'
$numFixTimeout = New-Object System.Windows.Forms.NumericUpDown
$numFixTimeout.Location = New-Object System.Drawing.Point(798, 8)
$numFixTimeout.Size = New-Object System.Drawing.Size(60, 24)
$numFixTimeout.Minimum = 5
$numFixTimeout.Maximum = 300
$numFixTimeout.Value = 60
$pnlFixTop.Controls.Add($numFixTimeout)
$null = New-Label $pnlFixTop 'sec' 862 10 30 'Black'

$chkFixLocalhost = New-Object System.Windows.Forms.CheckBox
$chkFixLocalhost.Text = 'Treat demo hosts as localhost (lab only)'
$chkFixLocalhost.Location = New-Object System.Drawing.Point(12, 74)
$chkFixLocalhost.Size = New-Object System.Drawing.Size(300, 22)
$chkFixLocalhost.Checked = $false
$pnlFixTop.Controls.Add($chkFixLocalhost)

$btnFixPreview = New-Button $pnlFixTop 'Preview plan' 330 72 120 26
$btnFix = New-Button $pnlFixTop 'Fix (WinRM)' 460 72 130 26
$btnFixClearLog = New-Button $pnlFixTop 'Clear log' 600 72 100 26
$btnFixPreflight = New-Button $pnlFixTop 'Test WinRM' 710 72 110 26
$btnFixExportLog = New-Button $pnlFixTop 'Export log' 830 72 100 26
$btnFixRetry = New-Button $pnlFixTop 'Retry failed' 12 106 110 26
$btnFixStop = New-Button $pnlFixTop 'Stop' 130 106 80 26

foreach ($b in @($btnFixPreview, $btnFix, $btnFixClearLog, $btnFixPreflight, $btnFixExportLog, $btnFixRetry, $btnFixStop)) {
    $b.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 252)
    $b.ForeColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
}
$btnFix.BackColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
$btnFix.ForeColor = [System.Drawing.Color]::White
$btnFixStop.BackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
$btnFixStop.ForeColor = [System.Drawing.Color]::White
$btnFixStop.Enabled = $false
$null = New-Label $pnlFixTop 'Workflow: scope + severity -> Test WinRM (optional) -> Preview plan -> Fix. Retry re-runs last Failed/WinRM-Failed rows.' 220 110 680 'Gray'

$uiTip.SetToolTip($btnFix, 'Run only the WinRM solutions relevant to the selected/filtered/failed errors')
$uiTip.SetToolTip($chkFixWhatIf, 'Build and show the plan without calling Invoke-Command')
$uiTip.SetToolTip($chkFixLocalhost, 'Maps every host to localhost - use only when testing against this machine')
$uiTip.SetToolTip($chkFixCurrentCred, 'Use the account running this app for WinRM; uncheck to enter explicit user/password')
$uiTip.SetToolTip($cmbFixSeverity, 'Limit the plan to Critical only or Critical+High failures')
$uiTip.SetToolTip($btnFixPreflight, 'Test-WSMan every remotable host in the current scope before Fix')
$uiTip.SetToolTip($btnFixExportLog, 'Save the Remote Fix log to a text file')
$uiTip.SetToolTip($btnFixRetry, 'Rebuild plan from the last run and re-execute only Failed / WinRM-Failed entries')
$uiTip.SetToolTip($numFixTimeout, 'WinRM open/operation timeout in seconds per host')
$uiTip.SetToolTip($btnFixStop, 'Stop after the current host finishes (cooperative cancel)')

$pnlFixLog = New-Object System.Windows.Forms.Panel
$pnlFixLog.Dock = 'Fill'
$pnlFixLog.Padding = New-Object System.Windows.Forms.Padding(6)
$tabRemoteFix.Controls.Add($pnlFixLog)
$pnlFixLog.BringToFront()

$rtbFixLog = New-Object System.Windows.Forms.RichTextBox
$rtbFixLog.Dock = 'Fill'
$rtbFixLog.ReadOnly = $true
$rtbFixLog.WordWrap = $true
$rtbFixLog.BackColor = [System.Drawing.Color]::FromArgb(20, 24, 30)
$rtbFixLog.ForeColor = [System.Drawing.Color]::Gainsboro
$rtbFixLog.Font = New-Object System.Drawing.Font('Consolas', 9.5)
$rtbFixLog.BorderStyle = 'None'
$rtbFixLog.Text = "Remote Fix tab - WinRM / PS remoting.`r`nLoad data on Dashboard, choose scope + severity, optional Test WinRM, Preview plan, then Fix.`r`nOnly solutions that match each error's rule/category are executed. Retry failed re-runs the last failures.`r`n`r`n"
$pnlFixLog.Controls.Add($rtbFixLog)

# Status strip
$status = New-Object System.Windows.Forms.StatusStrip
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = 'Ready - connect to BigFix API, import a report, or load demo data.'
$lblStatus.Spring = $true
$lblStatus.TextAlign = 'MiddleLeft'
[void]$status.Items.Add($lblStatus)
$form.Controls.Add($status)
$status.SendToBack()

# ---- Data table ----
function Initialize-DataTable {
    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add('DeviceName', [string])
    [void]$dt.Columns.Add('IPAddress', [string])
    [void]$dt.Columns.Add('LastReportTime', [string])
    [void]$dt.Columns.Add('HoursSinceReport', [string])
    [void]$dt.Columns.Add('Category', [string])
    [void]$dt.Columns.Add('Severity', [string])
    [void]$dt.Columns.Add('SeverityRank', [int])
    [void]$dt.Columns.Add('Component', [string])
    [void]$dt.Columns.Add('ErrorCode', [string])
    [void]$dt.Columns.Add('ErrorMessage', [string])
    [void]$dt.Columns.Add('Status', [string])
    [void]$dt.Columns.Add('Source', [string])
    [void]$dt.Columns.Add('RootCause', [string])
    [void]$dt.Columns.Add('Remediation', [string])
    [void]$dt.Columns.Add('MatchedRule', [string])
    $script:DataTable = $dt
    $script:DataView = New-Object System.Data.DataView($dt)
    $grid.DataSource = $script:DataView
    Configure-GridColumns
}

function Configure-GridColumns {
    $grid.Columns['SeverityRank'].Visible = $false
    $widths = @{
        DeviceName = 130; IPAddress = 110; LastReportTime = 130; HoursSinceReport = 70
        Category = 120; Severity = 80; Component = 170; ErrorCode = 90
        ErrorMessage = 280; Status = 80; Source = 90
    }
    foreach ($kv in $widths.GetEnumerator()) {
        if ($grid.Columns[$kv.Key]) { $grid.Columns[$kv.Key].Width = $kv.Value }
    }
    foreach ($name in @('RootCause', 'Remediation', 'MatchedRule', 'SeverityRank')) {
        if ($grid.Columns[$name]) { $grid.Columns[$name].Visible = $false }
    }
    if ($grid.Columns['ErrorMessage']) { $grid.Columns['ErrorMessage'].AutoSizeMode = 'Fill' }
}

function Update-GridFromRecords {
    try {
        $script:DataTable.Rows.Clear()
        if ($chkDeviceView.Checked) {
            $rows = ConvertTo-ObjectArray -Value (Get-DeviceAggregates -Records $script:AllRecords)
        } else {
            $rows = ConvertTo-ObjectArray -Value $script:AllRecords
        }
        $script:GridRowCount = $rows.Count
        $script:DataTable.BeginLoadData()
        try {
            foreach ($r in $rows) {
                $rank = 0
                $rankTmp = 0
                if ([int]::TryParse("$($r.SeverityRank)", [ref]$rankTmp)) { $rank = $rankTmp }
                [void]$script:DataTable.Rows.Add(
                    [object[]](
                        [string]$r.DeviceName, [string]$r.IPAddress, [string]$r.LastReportTime, [string]"$($r.HoursSinceReport)",
                        [string]$r.Category, [string]$r.Severity, $rank, [string]$r.Component, [string]$r.ErrorCode,
                        [string]$r.ErrorMessage, [string]$r.Status, [string]$r.Source, [string]$r.RootCause, [string]$r.Remediation, [string]$r.MatchedRule
                    )
                )
            }
        } finally {
            $script:DataTable.EndLoadData()
        }
        Apply-Filters
        Update-Charts
    } catch {
        $err = $_
        try { [System.IO.File]::AppendAllText((Join-Path $env:TEMP 'BigFixDashboard-UIErrors.log'), "$(Get-Date -Format o) Update-GridFromRecords: $($err.Exception.Message)`r`n$($err.InvocationInfo.PositionMessage)`r`n$($err.ScriptStackTrace)`r`n`r`n") } catch { }
        throw
    }
}

function Apply-Filters {
    if (-not $script:DataView) { return }
    $cat = [string]$cmbCategory.SelectedItem
    $sev = [string]$cmbSeverity.SelectedItem
    $q = ($txtSearch.Text -replace '[\[\]\*#]', '').Trim()

    $clauses = @()
    if ($cat -and $cat -ne 'All') {
        $catEsc = $cat.Replace("'", "''")
        $clauses += "Category = '" + $catEsc + "'"
    }
    if ($sev -and $sev -ne 'All') {
        $sevEsc = $sev.Replace("'", "''")
        $clauses += "Severity = '" + $sevEsc + "'"
    }
    if ($q) {
        $clauses += "(DeviceName LIKE '%$q%' OR Component LIKE '%$q%' OR ErrorMessage LIKE '%$q%' OR ErrorCode LIKE '%$q%' OR RootCause LIKE '%$q%' OR IPAddress LIKE '%$q%')"
    }
    $script:DataView.RowFilter = ($clauses -join ' AND ')

    $count = $script:DataView.Count
    $lblFilterCount.Text = "Showing $count of $($script:DataTable.Rows.Count)"
    # Cache summary so filter keystrokes do not re-aggregate full set every time
    if ($null -eq $script:CachedSummary -or $script:SummaryDirty) {
        $script:CachedSummary = Get-FailureSummary $script:AllRecords
        $script:SummaryDirty = $false
    }
    $sum = $script:CachedSummary
    $tileTotal.Text = "$($sum.Total)"
    $tileDev.Text = "$($sum.Devices)"
    $tileOff.Text = "$($sum.Offline)"
    $tileFix.Text = "$($sum.FixletFailures)"
    $tileComp.Text = "$($sum.Compliance)"
    $tileSw.Text = "$($sum.SoftwareDeploy)"
    $tileCrit.Text = "$($sum.Critical + $sum.High)"
    $tileHealthy.Text = "$($sum.Healthy)"

    if ($grid.Rows.Count -gt 0) {
        if (-not $grid.CurrentRow -or $grid.CurrentRow.Index -lt 0) {
            $grid.ClearSelection()
            $grid.Rows[0].Selected = $true
        }
        Show-SelectedDetails
    } else {
        $rtbDetails.Clear()
        $rtbDetails.AppendText("No devices match the current filters.`r`n")
        $rtbDetails.AppendText("Connect to BigFix API to load devices, or Load Demo / Import Report.`r`n")
    }
}

function Show-SelectedDetails {
    $rtbDetails.Clear()
    if ($grid.CurrentRow -or ($grid.Rows.Count -gt 0 -and -not $grid.CurrentRow)) {
        $rowIdx = -1
        if ($grid.CurrentRow) { $rowIdx = $grid.CurrentRow.Index }
        elseif ($grid.Rows.Count -gt 0) { $rowIdx = 0 }
        if ($rowIdx -lt 0) { return }

        $viewRow = $script:DataView[$rowIdx]
        $dev = $viewRow['DeviceName']
        $err = $viewRow['ErrorMessage']
        $comp = $viewRow['Component']

        $rtbDetails.SelectionStart = $rtbDetails.TextLength
        $rtbDetails.SelectionColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
        $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 11, [System.Drawing.FontStyle]::Bold)
        $rtbDetails.AppendText("$dev`r`n")
        $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 9.5)

        # Device view: show full failure timeline for this host
        if ($chkDeviceView.Checked) {
            $deviceRecs = @($script:AllRecords | Where-Object { $_.DeviceName -eq $dev })
            $rtbDetails.SelectionColor = [System.Drawing.Color]::Black
            $failedN = @($deviceRecs | Where-Object { $_.Category -ne 'Healthy' }).Count
            $ipA = [string]$viewRow['IPAddress']
            $lrt = [string]$viewRow['LastReportTime']
            $hrs = [string]$viewRow['HoursSinceReport']
            $worstSev = [string]$viewRow['Severity']
            $worstCat = [string]$viewRow['Category']
            $rtbDetails.AppendText(("Device view  : {0} record(s), {1} failure(s)`r`n" -f $deviceRecs.Count, $failedN))
            $rtbDetails.AppendText(("IP Address   : {0}`r`n" -f $ipA))
            $rtbDetails.AppendText(("Last Report  : {0}  ({1} h ago)`r`n" -f $lrt, $hrs))
            $rtbDetails.AppendText(("Worst        : {0} / {1}`r`n" -f $worstSev, $worstCat))
            $rtbDetails.AppendText("`r`n")
            $rtbDetails.SelectionColor = [System.Drawing.Color]::DarkOrange
            $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 10, [System.Drawing.FontStyle]::Bold)
            $rtbDetails.AppendText("FAILURE TIMELINE`r`n")
            $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 9.5)
            $i = 1
            foreach ($fr in ($deviceRecs | Sort-Object -Property SeverityRank -Descending)) {
                $rtbDetails.SelectionColor = switch ([string]$fr.Severity) {
                    'Critical' { [System.Drawing.Color]::DarkRed }
                    'High'     { [System.Drawing.Color]::FromArgb(180, 60, 0) }
                    'Medium'   { [System.Drawing.Color]::DarkGoldenrod }
                    default    { [System.Drawing.Color]::Gray }
                }
                $rtbDetails.AppendText(("[{0}] {1} | {2} | {3}`r`n" -f $i, $fr.Severity, $fr.Category, $fr.Component))
                $rtbDetails.SelectionColor = [System.Drawing.Color]::FromArgb(120, 20, 20)
                $rtbDetails.AppendText("     $($fr.ErrorMessage)`r`n")
                if ($fr.RootCause) {
                    $rtbDetails.SelectionColor = [System.Drawing.Color]::FromArgb(20, 70, 20)
                    $rtbDetails.AppendText("     Root: $($fr.RootCause)`r`n")
                }
                $i++
                if ($i -gt 30) {
                    $rtbDetails.SelectionColor = [System.Drawing.Color]::Gray
                    $rtbDetails.AppendText("     ... $(@($deviceRecs).Count - 30) more`r`n")
                    break
                }
            }
            $rtbDetails.SelectionColor = [System.Drawing.Color]::Black
            $rtbDetails.SelectionStart = 0
            $rtbDetails.SelectionLength = 0
            return
        }

        $rec = $null
        foreach ($candidate in $script:AllRecords) {
            if ($candidate.DeviceName -eq $dev -and $candidate.ErrorMessage -eq $err -and $candidate.Component -eq $comp) {
                $rec = $candidate
                break
            }
        }

        $rtbDetails.SelectionColor = [System.Drawing.Color]::Black
        $lines = @(
            ("Category      : {0}" -f [string]$viewRow['Category']),
            ("Severity      : {0}" -f [string]$viewRow['Severity']),
            ("Status        : {0}" -f [string]$viewRow['Status']),
            ("IP Address    : {0}" -f [string]$viewRow['IPAddress']),
            ("Last Report   : {0}  ({1} h ago)" -f [string]$viewRow['LastReportTime'], [string]$viewRow['HoursSinceReport']),
            ("Component     : {0}" -f [string]$viewRow['Component']),
            ("Error Code    : {0}" -f [string]$viewRow['ErrorCode']),
            ("Source        : {0}" -f [string]$viewRow['Source']),
            ("Matched Rule  : {0}" -f [string]$viewRow['MatchedRule'])
        )
        foreach ($l in $lines) { $rtbDetails.AppendText("$l`r`n") }

        $rtbDetails.AppendText("`r`n")
        $rtbDetails.SelectionColor = [System.Drawing.Color]::DarkRed
        $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 10, [System.Drawing.FontStyle]::Bold)
        if ([string]$viewRow['Category'] -eq 'Healthy') {
            $rtbDetails.SelectionColor = [System.Drawing.Color]::ForestGreen
            $rtbDetails.AppendText("STATUS`r`n")
        } else {
            $rtbDetails.AppendText("RAW ERROR`r`n")
        }
        $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 9.5)
        $rtbDetails.SelectionColor = if ([string]$viewRow['Category'] -eq 'Healthy') { [System.Drawing.Color]::FromArgb(20, 90, 20) } else { [System.Drawing.Color]::FromArgb(120, 20, 20) }
        $rtbDetails.AppendText("$($viewRow['ErrorMessage'])`r`n")

        $root = [string]$viewRow['RootCause']
        $rem = [string]$viewRow['Remediation']
        if ($rec) { if ($rec.RootCause) { $root = $rec.RootCause }; if ($rec.Remediation) { $rem = $rec.Remediation } }

        if ($root -and $root -notmatch 'Unclassified') {
            $rtbDetails.AppendText("`r`n")
            $rtbDetails.SelectionColor = [System.Drawing.Color]::DarkGreen
            $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 10, [System.Drawing.FontStyle]::Bold)
            $rtbDetails.AppendText("ROOT CAUSE`r`n")
            $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 9.5)
            $rtbDetails.SelectionColor = [System.Drawing.Color]::FromArgb(20, 70, 20)
            $rtbDetails.AppendText("$root`r`n")
        }

        if ($rem) {
            $rtbDetails.AppendText("`r`n")
            $rtbDetails.SelectionColor = [System.Drawing.Color]::FromArgb(0, 70, 130)
            $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 10, [System.Drawing.FontStyle]::Bold)
            $rtbDetails.AppendText("REMEDIATION`r`n")
            $rtbDetails.SelectionFont = New-Object System.Drawing.Font('Consolas', 9.5)
            $rtbDetails.SelectionColor = [System.Drawing.Color]::FromArgb(0, 50, 100)
            $rtbDetails.AppendText("$rem`r`n")
        }

        $rtbDetails.SelectionColor = [System.Drawing.Color]::Black
        $rtbDetails.SelectionStart = 0
        $rtbDetails.SelectionLength = 0
    }
}

function Set-Status {
    param([string]$Text)
    $lblStatus.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-ButtonsEnabled {
    param([bool]$Enabled)
    foreach ($b in @($btnConnect, $btnImport, $btnDemo, $btnRefresh, $btnExport, $btnHtml, $btnEmail, $btnTicket)) {
        if ($b) { $b.Enabled = $Enabled }
    }
    $btnFix.Enabled = $Enabled
    $btnFixPreview.Enabled = $Enabled
    if ($btnFixPreflight) { $btnFixPreflight.Enabled = $Enabled }
    if ($btnFixExportLog) { $btnFixExportLog.Enabled = $Enabled }
    if ($btnFixRetry) { $btnFixRetry.Enabled = $Enabled }
    if ($Enabled) {
        if ($btnFixStop) { $btnFixStop.Enabled = [bool]$script:FixRunning }
        if ($btnFixRetry -and -not $script:LastFixResults) { $btnFixRetry.Enabled = $false }
    } else {
        if ($btnFixStop) { $btnFixStop.Enabled = $false }
    }
}

function Write-FixLog {
    param([string]$Text, [string]$Color = 'Gainsboro')
    $rtbFixLog.SelectionStart = $rtbFixLog.TextLength
    $rtbFixLog.SelectionLength = 0
    $rtbFixLog.SelectionColor = [System.Drawing.Color]::FromName($Color)
    $rtbFixLog.AppendText("$Text`r`n")
    $rtbFixLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-FixScopeRecords {
    $scope = 'Selected'
    if ($radScopeFiltered.Checked) { $scope = 'Filtered' }
    elseif ($radScopeAll.Checked) { $scope = 'All' }

    switch ($scope) {
        'Selected' {
            $rows = @()
            if ($grid.CurrentRow -and $script:DataView) {
                $idx = $grid.CurrentRow.Index
                if ($idx -ge 0 -and $idx -lt $script:DataView.Count) {
                    $viewRow = $script:DataView[$idx]
                    $dev = $viewRow['DeviceName']; $err = $viewRow['ErrorMessage']; $comp = $viewRow['Component']
                    if ($chkDeviceView.Checked) {
                        $rows = @($script:AllRecords | Where-Object { $_.DeviceName -eq $dev })
                    } else {
                        $rows = @($script:AllRecords | Where-Object {
                            $_.DeviceName -eq $dev -and $_.ErrorMessage -eq $err -and $_.Component -eq $comp
                        })
                    }
                }
            }
            if ($rows.Count -eq 0) {
                [void][System.Windows.Forms.MessageBox]::Show('Select a row on the Dashboard grid first.', 'Remote Fix', 'OK', 'Information')
                return @()
            }
            return @($rows | Where-Object { $_.Category -ne 'Healthy' })
        }
        'Filtered' {
            $rows = New-Object System.Collections.Generic.List[object]
            if ($script:DataView) {
                if ($chkDeviceView.Checked) {
                    $devNames = @{}
                    foreach ($viewRow in $script:DataView) {
                        $devNames[[string]$viewRow['DeviceName']] = $true
                    }
                    foreach ($rec in $script:AllRecords) {
                        if ($devNames.ContainsKey([string]$rec.DeviceName) -and $rec.Category -ne 'Healthy') { $rows.Add($rec) }
                    }
                } else {
                    foreach ($viewRow in $script:DataView) {
                        $dev = $viewRow['DeviceName']; $err = $viewRow['ErrorMessage']; $comp = $viewRow['Component']
                        $rec = $null
                        foreach ($candidate in $script:AllRecords) {
                            if ($candidate.DeviceName -eq $dev -and $candidate.ErrorMessage -eq $err -and $candidate.Component -eq $comp) {
                                $rec = $candidate
                                break
                            }
                        }
                        if ($rec -and $rec.Category -ne 'Healthy') { $rows.Add($rec) }
                    }
                }
            }
            if ($rows.Count -eq 0) {
                [void][System.Windows.Forms.MessageBox]::Show('No non-Healthy rows match current filters.', 'Remote Fix', 'OK', 'Information')
            }
            return @($rows)
        }
        default {
            $rows = @($script:AllRecords | Where-Object { $_.Category -ne 'Healthy' })
            if ($rows.Count -eq 0) {
                [void][System.Windows.Forms.MessageBox]::Show('No failed records loaded. Connect / Import / Load Demo first.', 'Remote Fix', 'OK', 'Information')
            }
            return $rows
        }
    }
}

function Resolve-FixHostTarget {
    param([string]$DeviceName)
    if ($chkFixLocalhost.Checked) { return 'localhost' }
    return $DeviceName
}

function Get-FixMinSeverityChoice {
    if (-not $cmbFixSeverity) { return '' }
    switch ("$($cmbFixSeverity.SelectedItem)") {
        'Critical only'  { return 'Critical' }
        'Critical + High'{ return 'High' }
        default          { return '' }
    }
}

function Get-FixScopedPlan {
    param([switch]$FromRetry)
    if ($FromRetry) {
        $failed = @($script:LastFixResults | Where-Object { $_.Status -in @('Failed', 'WinRM-Failed') })
        if ($failed.Count -eq 0) { return @() }
        $plan = New-Object System.Collections.Generic.List[object]
        $seen = @{}
        foreach ($fr in $failed) {
            $frDev = [string]$fr.DeviceName
            $frSol = [string]$fr.SolutionId
            $key = $frDev + '|' + $frSol
            if ($seen.ContainsKey($key)) { continue }
            $src = $null
            foreach ($p in @($script:LastFixPlan)) {
                if (([string]$p.DeviceName -eq $frDev) -and ([string]$p.SolutionId -eq $frSol)) { $src = $p; break }
            }
            if (-not $src) { continue }
            $seen[$key] = $true
            $clone = $src | Select-Object *
            $plan.Add($clone)
        }
        return @($plan.ToArray())
    }
    $records = ConvertTo-ObjectArray -Value (Get-FixScopeRecords)
    if (-not $records -or $records.Count -eq 0) { return @() }
    $minSev = Get-FixMinSeverityChoice
    return @(Get-FixPlan -Records $records -MinSeverity $minSev)
}

function Resolve-FixCredential {
    if ($chkFixCurrentCred.Checked) { return $null }
    $u = $txtFixUser.Text.Trim()
    if (-not $u) { throw 'Enter WinRM username or enable "Use current Windows credentials".' }
    $sec = ConvertTo-SecureString -String $txtFixPass.Text -AsPlainText -Force
    return (New-Object System.Management.Automation.PSCredential($u, $sec))
}

function Complete-FixRun {
    param(
        [System.Collections.IEnumerable]$Results,
        [string]$Mode = 'Execute'
    )
    $script:LastFixResults = ConvertTo-ObjectArray -Value $Results
    $results = $script:LastFixResults

    $okN = @($results | Where-Object { $_.Status -eq 'OK' }).Count
    $failN = @($results | Where-Object { $_.Status -in @('Failed', 'WinRM-Failed') }).Count
    $manN = @($results | Where-Object { $_.Status -eq 'Manual' }).Count
    $skipN = @($results | Where-Object { $_.Status -in @('Skipped', 'WhatIf') }).Count

    foreach ($r in @($results)) {
        $col = switch ($r.Status) {
            'OK'           { 'LightGreen' }
            'Failed'       { 'Salmon' }
            'WinRM-Failed' { 'OrangeRed' }
            'Manual'       { 'Khaki' }
            'WhatIf'       { 'LightSkyBlue' }
            default        { 'Gray' }
        }
        $dur = ''
        if ($null -ne $r.PSObject.Properties['DurationSec'] -and $r.DurationSec) { $dur = " ($($r.DurationSec)s)" }
        Write-FixLog ("[{0}] {1} :: {2}{3} -> {4}" -f $r.Status, $r.DeviceName, $r.Title, $dur, $r.Output) $col
    }
    Write-FixLog "Done: OK=$okN Failed=$failN Manual=$manN Skip/WhatIf=$skipN" 'Yellow'
    if ($failN -gt 0) { Write-FixLog 'Tip: use Retry failed to re-run only Failed / WinRM-Failed entries.' 'Khaki' }

    $hist = Save-FixRunHistory -Results $results -Mode $Mode
    if ($hist) { Write-FixLog "History: $hist" 'Gray' }

    if ($btnFixRetry) { $btnFixRetry.Enabled = ($failN -gt 0) }
    Set-Status "Remote Fix complete - OK=$okN Failed=$failN Manual=$manN"
}

function Show-FixPlanPreview {
    $records = ConvertTo-ObjectArray -Value (Get-FixScopeRecords)
    if (-not $records -or $records.Count -eq 0) { return $null }
    $minSev = Get-FixMinSeverityChoice
    $plan = Get-FixPlan -Records $records -MinSeverity $minSev
    $text = Format-FixPlanPreview -Plan $plan
    $sevLabel = if ($minSev) { $minSev } else { 'all' }
    Write-FixLog ("---- Plan {0} | severity={1} ----" -f (Get-Date -Format 'HH:mm:ss'), $sevLabel) 'Yellow'
    foreach ($line in ($text -split "`r?`n")) {
        $col = 'Gainsboro'
        if ($line -match '^\[MANUAL') { $col = 'Khaki' }
        elseif ($line -match '^\[WINRM') { $col = 'LightSkyBlue' }
        elseif ($line -match '^\[SKIP') { $col = 'Gray' }
        elseif ($line -match '^Plan:') { $col = 'Yellow' }
        elseif ($line -match '^Severity:') { $col = 'Yellow' }
        Write-FixLog $line $col
    }
    return $plan
}

function Load-Records {
    param(
        [System.Collections.IEnumerable]$Records,
        [string]$SourceLabel
    )
    $script:AllRecords.Clear()
    foreach ($r in (ConvertTo-ObjectArray -Value $Records)) { if ($null -ne $r) { $script:AllRecords.Add($r) } }
    $script:LastSource = $SourceLabel
    $script:CachedSummary = Get-FailureSummary $script:AllRecords
    $script:SummaryDirty = $false
    if (-not $script:DataTable) { Initialize-DataTable }
    try {
        Update-GridFromRecords
    } catch {
        $err = $_
        try { [System.IO.File]::AppendAllText((Join-Path $env:TEMP 'BigFixDashboard-UIErrors.log'), "$(Get-Date -Format o) Update-GridFromRecords: $($err.Exception.Message)`r`n$($err.InvocationInfo.PositionMessage)`r`n`r`n") } catch { }
        throw
    }
    $sum = $script:CachedSummary
    Set-Status "$SourceLabel - $($sum.Total) rows, $($sum.Devices) devices (Healthy: $($sum.Healthy), Offline: $($sum.Offline), Fixlet: $($sum.FixletFailures), Compliance: $($sum.Compliance))"
    Test-NewCriticalAlerts
}

# ---- Charts / alerts / settings / async helpers ----
function Update-Charts {
    if (-not $script:ChartsAvailable) { return }
    try {
        $records = ConvertTo-ObjectArray -Value $script:AllRecords
        $failed = @($records | Where-Object { $_.Category -ne 'Healthy' })
        if ($failed.Count -gt 0) { $source = [object[]]$failed } else { $source = [object[]]$records }

        # Category pie
        $ch = $script:ChartCategory
        $ch.Series.Clear()
        $s = $ch.Series.Add('Category')
        $s.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
        $catColors = @{
            'Offline' = '#b00020'; 'FixletFailure' = '#e65100'
            'ComplianceFailure' = '#6a1b9a'; 'SoftwareDeployment' = '#00695c'
            'Healthy' = '#1b7a3d'
        }
        foreach ($g in ($source | Group-Object Category)) {
            $idx = $s.Points.AddXY($g.Name, $g.Count)
            if ($catColors.ContainsKey($g.Name)) {
                $s.Points[$idx].Color = [System.Drawing.ColorTranslator]::FromHtml($catColors[$g.Name])
            }
        }

        # Severity column
        $ch = $script:ChartSeverity
        $ch.Series.Clear()
        $s = $ch.Series.Add('Severity')
        $s.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
        $s.IsValueShownAsLabel = $true
        foreach ($name in @('Critical', 'High', 'Medium', 'Low')) {
            $n = @($source | Where-Object { $_.Severity -eq $name }).Count
            [void]$s.Points.AddXY($name, $n)
        }

        # Top devices
        $ch = $script:ChartTopDev
        $ch.Series.Clear()
        $s = $ch.Series.Add('Failures')
        $s.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Bar
        $s.IsValueShownAsLabel = $true
        foreach ($g in ($failed | Group-Object DeviceName | Sort-Object Count -Descending | Select-Object -First 10)) {
            [void]$s.Points.AddXY($g.Name, $g.Count)
        }

        # Top rules
        $ch = $script:ChartRules
        $ch.Series.Clear()
        $s = $ch.Series.Add('Rules')
        $s.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Bar
        $s.IsValueShownAsLabel = $true
        foreach ($g in ($source | Group-Object MatchedRule | Sort-Object Count -Descending | Select-Object -First 10)) {
            [void]$s.Points.AddXY($g.Name, $g.Count)
        }
    } catch { }
}

function Test-NewCriticalAlerts {
    if (-not $script:AppSettings.AlertOnNewCritical) { return }
    if (-not $script:NotifyIcon) { return }
    $crit = @($script:AllRecords | Where-Object { $_.SeverityRank -ge 4 -and $_.Category -ne 'Healthy' }).Count
    $prev = [int]$script:PrevCriticalCount
    $script:PrevCriticalCount = $crit
    if ($crit -gt $prev -and $crit -gt 0) {
        try {
            $script:NotifyIcon.Visible = $true
            $script:NotifyIcon.ShowBalloonTip(5000, 'BigFix Failure Dashboard', "$crit critical/high failures (was $prev). Open the dashboard for details.", [System.Windows.Forms.ToolTipIcon]::Warning)
        } catch { }
    }
}

function Persist-UiSettings {
    $s = $script:AppSettings
    if (-not $s) { return }
    try {
        $s.Server = $txtServer.Text.Trim()
        $s.Username = $txtUser.Text.Trim()
        $s.OfflineThresholdHours = [int]$numOffline.Value
        $s.SkipTlsCertCheck = [bool]$chkSkipCert.Checked
        $s.AutoRefreshMinutes = [int]$numAutoRefresh.Value
        $s.CategoryFilter = [string]$cmbCategory.SelectedItem
        $s.SeverityFilter = [string]$cmbSeverity.SelectedItem
        $s.DeviceGroupView = [bool]$chkDeviceView.Checked
        $null = Save-AppSettings -Settings $s
    } catch { }
}

function Show-SettingsDialog {
    param([switch]$PassThru)
    $s = $script:AppSettings
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Settings - BigFix Failure Dashboard v$script:AppVersion"
    $dlg.Size = New-Object System.Drawing.Size(540, 720)
    $dlg.MinimumSize = New-Object System.Drawing.Size(540, 560)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.AutoScroll = $true
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $y = 12
    function Add-SetLabel($text, [ref]$y) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text
        $l.Location = New-Object System.Drawing.Point(16, $y.Value)
        $l.Size = New-Object System.Drawing.Size(490, 18)
        $l.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $l.ForeColor = [System.Drawing.Color]::FromArgb(30, 56, 92)
        $dlg.Controls.Add($l)
        $y.Value += 22
    }
    function Add-SetBox($text, $value, [ref]$y, [switch]$Multiline) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text
        $l.Location = New-Object System.Drawing.Point(16, ($y.Value + 3))
        $l.Size = New-Object System.Drawing.Size(130, 18)
        $dlg.Controls.Add($l)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(150, $y.Value)
        $t.Size = New-Object System.Drawing.Size(350, $(if ($Multiline) { 56 } else { 24 }))
        if ($Multiline) { $t.Multiline = $true; $t.ScrollBars = 'Vertical' }
        $t.Text = "$value"
        $dlg.Controls.Add($t)
        $y.Value += $(if ($Multiline) { 64 } else { 30 })
        return $t
    }

    Add-SetLabel 'General' ([ref]$y)
    $txtSetServer = Add-SetBox 'Server' $s.Server ([ref]$y)
    $txtSetUser = Add-SetBox 'API username' $s.Username ([ref]$y)
    $numSetOffline = Add-SetBox 'Offline threshold (h)' $s.OfflineThresholdHours ([ref]$y)
    $chkSetSkip = New-Object System.Windows.Forms.CheckBox
    $chkSetSkip.Text = 'Skip TLS cert check (lab only)'
    $chkSetSkip.Location = New-Object System.Drawing.Point(150, $y)
    $chkSetSkip.Size = New-Object System.Drawing.Size(350, 22)
    $chkSetSkip.Checked = [bool]$s.SkipTlsCertCheck
    $dlg.Controls.Add($chkSetSkip)
    $y += 30
    $numSetAuto = Add-SetBox 'Auto-refresh (min, 0=off)' $s.AutoRefreshMinutes ([ref]$y)
    $chkSetAlert = New-Object System.Windows.Forms.CheckBox
    $chkSetAlert.Text = 'Balloon alert when new critical/high failures appear'
    $chkSetAlert.Location = New-Object System.Drawing.Point(150, $y)
    $chkSetAlert.Size = New-Object System.Drawing.Size(350, 22)
    $chkSetAlert.Checked = [bool]$s.AlertOnNewCritical
    $dlg.Controls.Add($chkSetAlert)
    $y += 34

    Add-SetLabel 'SMTP email digest (password not saved)' ([ref]$y)
    $chkSmtp = New-Object System.Windows.Forms.CheckBox
    $chkSmtp.Text = 'Enable email'
    $chkSmtp.Location = New-Object System.Drawing.Point(150, $y)
    $chkSmtp.Size = New-Object System.Drawing.Size(200, 22)
    $chkSmtp.Checked = [bool]$s.Smtp.Enable
    $dlg.Controls.Add($chkSmtp)
    $y += 28
    $txtSmtpHost = Add-SetBox 'SMTP host' $s.Smtp.Host ([ref]$y)
    $txtSmtpPort = Add-SetBox 'Port' $s.Smtp.Port ([ref]$y)
    $chkSmtpSsl = New-Object System.Windows.Forms.CheckBox
    $chkSmtpSsl.Text = 'Use SSL/TLS'
    $chkSmtpSsl.Location = New-Object System.Drawing.Point(150, $y)
    $chkSmtpSsl.Size = New-Object System.Drawing.Size(200, 22)
    $chkSmtpSsl.Checked = [bool]$s.Smtp.UseSsl
    $dlg.Controls.Add($chkSmtpSsl)
    $y += 28
    $txtSmtpFrom = Add-SetBox 'From' $s.Smtp.From ([ref]$y)
    $txtSmtpTo = Add-SetBox 'To (semicolon list)' $s.Smtp.To ([ref]$y)
    $txtSmtpUser = Add-SetBox 'SMTP username' $s.Smtp.Username ([ref]$y)
    $txtSmtpPass = Add-SetBox 'SMTP password (session only)' '' ([ref]$y)
    $txtSmtpPass.UseSystemPasswordChar = $true

    Add-SetLabel 'Ticket webhook (JSON POST)' ([ref]$y)
    $chkTkt = New-Object System.Windows.Forms.CheckBox
    $chkTkt.Text = 'Enable ticket webhook'
    $chkTkt.Location = New-Object System.Drawing.Point(150, $y)
    $chkTkt.Size = New-Object System.Drawing.Size(250, 22)
    $chkTkt.Checked = [bool]$s.Ticket.Enable
    $dlg.Controls.Add($chkTkt)
    $y += 28
    $txtTktUrl = Add-SetBox 'Webhook URL' $s.Ticket.WebhookUrl ([ref]$y)
    $txtTktHdr = Add-SetBox 'Headers JSON' $s.Ticket.HeaderJson ([ref]$y) -Multiline

    # Buttons sit below all fields (content ends near $y; form is 720px tall)
    $btnY = [Math]::Max($y + 12, 660)
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Save'
    $btnOk.Location = New-Object System.Drawing.Point(300, $btnY)
    $btnOk.Size = New-Object System.Drawing.Size(90, 30)
    $btnOk.Anchor = 'Bottom,Right'
    $dlg.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(400, $btnY)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Anchor = 'Bottom,Right'
    $btnCancel.DialogResult = 'Cancel'
    $dlg.Controls.Add($btnCancel)
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel
    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "Password fields are session-only (never written to settings.json)."
    $lblHint.Location = New-Object System.Drawing.Point(16, ($btnY + 8))
    $lblHint.Size = New-Object System.Drawing.Size(480, 16)
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $dlg.Controls.Add($lblHint)

    $script:SettingsSaved = $false
    # GetNewClosure: capture dialog locals for the click handler (modal ShowDialog).
    $btnOk.Add_Click({
        if ($chkSmtp.Checked) {
            if (-not $txtSmtpHost.Text.Trim() -or -not $txtSmtpFrom.Text.Trim() -or -not $txtSmtpTo.Text.Trim()) {
                [void][System.Windows.Forms.MessageBox]::Show('Enable email requires SMTP host, From, and To.', 'Settings', 'OK', 'Warning')
                return
            }
        }
        if ($chkTkt.Checked -and -not $txtTktUrl.Text.Trim()) {
            [void][System.Windows.Forms.MessageBox]::Show('Enable ticket webhook requires a Webhook URL.', 'Settings', 'OK', 'Warning')
            return
        }
        $s.Server = $txtSetServer.Text.Trim()
        $s.Username = $txtSetUser.Text.Trim()
        $iv = 4
        if ([int]::TryParse($numSetOffline.Text, [ref]$iv)) { $s.OfflineThresholdHours = [Math]::Max(1, $iv) }
        $s.SkipTlsCertCheck = [bool]$chkSetSkip.Checked
        $av = 0
        if ([int]::TryParse($numSetAuto.Text, [ref]$av)) { $s.AutoRefreshMinutes = [Math]::Max(0, $av) }
        $s.AlertOnNewCritical = [bool]$chkSetAlert.Checked
        $s.Smtp.Enable = [bool]$chkSmtp.Checked
        $s.Smtp.Host = $txtSmtpHost.Text.Trim()
        $sp = 587
        if ([int]::TryParse($txtSmtpPort.Text, [ref]$sp)) { $s.Smtp.Port = $sp }
        $s.Smtp.UseSsl = [bool]$chkSmtpSsl.Checked
        $s.Smtp.From = $txtSmtpFrom.Text.Trim()
        $s.Smtp.To = $txtSmtpTo.Text.Trim()
        $s.Smtp.Username = $txtSmtpUser.Text.Trim()
        if ($txtSmtpPass.Text) { $script:SmtpPassword = $txtSmtpPass.Text }
        $s.Ticket.Enable = [bool]$chkTkt.Checked
        $s.Ticket.WebhookUrl = $txtTktUrl.Text.Trim()
        $s.Ticket.HeaderJson = $txtTktHdr.Text.Trim()
        if (-not (Save-AppSettings -Settings $s)) {
            [void][System.Windows.Forms.MessageBox]::Show("Could not write settings to:`r`n$script:SettingsPath`r`n`r`nCheck folder permissions.", 'Settings', 'OK', 'Error')
            return
        }
        $script:SettingsSaved = $true
        try {
            $txtServer.Text = $s.Server
            $txtUser.Text = $s.Username
            $numOffline.Value = [Math]::Min($numOffline.Maximum, [Math]::Max($numOffline.Minimum, [int]$s.OfflineThresholdHours))
            $chkSkipCert.Checked = [bool]$s.SkipTlsCertCheck
            $numAutoRefresh.Value = [Math]::Min($numAutoRefresh.Maximum, [Math]::Max(0, [int]$s.AutoRefreshMinutes))
        } catch { }
        $dlg.DialogResult = 'OK'
        $dlg.Close()
    }.GetNewClosure())

    $null = $dlg.ShowDialog($script:MainForm)
    $dlg.Dispose()
    Apply-UiSettingsFromAppSettings
    if ($PassThru) { return [bool]$script:SettingsSaved }
}

function Apply-UiSettingsFromAppSettings {
    $s = $script:AppSettings
    if (-not $s) { return }
    if ($s.Server) { $txtServer.Text = $s.Server }
    if ($s.Username) { $txtUser.Text = $s.Username }
    try { $numOffline.Value = [Math]::Min($numOffline.Maximum, [Math]::Max($numOffline.Minimum, [int]$s.OfflineThresholdHours)) } catch { }
    $chkSkipCert.Checked = [bool]$s.SkipTlsCertCheck
    try { $numAutoRefresh.Value = [Math]::Min($numAutoRefresh.Maximum, [Math]::Max(0, [int]$s.AutoRefreshMinutes)) } catch { }
    if ($s.CategoryFilter -and $cmbCategory.Items -contains $s.CategoryFilter) { $cmbCategory.SelectedItem = $s.CategoryFilter }
    if ($s.SeverityFilter -and $cmbSeverity.Items -contains $s.SeverityFilter) { $cmbSeverity.SelectedItem = $s.SeverityFilter }
    $chkDeviceView.Checked = [bool]$s.DeviceGroupView
    Update-AutoRefreshTimer
}

function Update-AutoRefreshTimer {
    $mins = [int]$numAutoRefresh.Value
    if ($mins -le 0) {
        $script:AutoRefreshTimer.Stop()
    } else {
        $script:AutoRefreshTimer.Interval = $mins * 60 * 1000
        $script:AutoRefreshTimer.Start()
    }
}

function Invoke-CurrentSourceRefresh {
    if ($script:LastSource -like 'BigFix API*') {
        Start-AsyncApiConnect -IsRefresh
    } elseif ($script:LastSource -like 'Demo*' -or $script:LastSource -like 'Cache*') {
        if ($script:LastSource -like 'Cache*') {
            $cached = Get-SessionCache
            if ($cached) { Load-Records -Records (ConvertTo-ObjectArray -Value $cached.Failures) -SourceLabel "Cache: $($cached.SavedAt)" }
        } else {
            Load-Records -Records (Get-DemoFailures) -SourceLabel 'Demo data'
        }
        Set-Status 'Refreshed.'
    }
}

function Get-VisibleRecords {
    $visible = New-Object System.Collections.Generic.List[object]
    if (-not $script:DataView) { return (ConvertTo-ObjectArray -Value $script:AllRecords) }
    if ($chkDeviceView.Checked) {
        $devNames = @{}
        foreach ($row in $script:DataView) { $devNames[[string]$row['DeviceName']] = $true }
        foreach ($rec in $script:AllRecords) {
            if ($devNames.ContainsKey([string]$rec.DeviceName)) { $visible.Add($rec) }
        }
    } else {
        foreach ($row in $script:DataView) {
            $dev = $row['DeviceName']; $err = $row['ErrorMessage']; $comp = $row['Component']
            foreach ($candidate in $script:AllRecords) {
                if ($candidate.DeviceName -eq $dev -and $candidate.ErrorMessage -eq $err -and $candidate.Component -eq $comp) {
                    $visible.Add($candidate)
                    break
                }
            }
        }
    }
    return (ConvertTo-ObjectArray -Value $visible)
}

function Start-AsyncApiConnect {
    param([switch]$IsRefresh)
    if ($script:BgJob) {
        Set-Status 'API load already in progress...'
        return
    }
    $script:OfflineThresholdHours = [int]$numOffline.Value
    $url = $txtServer.Text.Trim()
    if (-not $url) { throw 'Enter BigFix server URL (https://server:52311).' }
    $url = Get-BigFixBaseUri -Url $url
    $txtServer.Text = $url
    if (-not $txtUser.Text.Trim()) {
        throw "Enter BigFix Console username (operator with 'Can use REST API' permission)."
    }
    Persist-UiSettings
    Start-AsyncApiPhase -Phase 'Inventory' -BaseUrl $url -IsRefresh:$IsRefresh
}

function Start-AsyncApiPhase {
    param(
        [ValidateSet('Inventory', 'Failures')][string]$Phase,
        [string]$BaseUrl,
        [switch]$IsRefresh,
        $Seed = $null
    )
    $tmp = Join-Path $env:TEMP ("bfd-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    $cfgPath = Join-Path $tmp 'cfg.json'
    $progressFile = Join-Path $tmp 'progress.txt'
    $resultFile = Join-Path $tmp 'result.json'
    $cfg = [pscustomobject]@{
        Phase = $Phase
        BaseUrl = $BaseUrl
        Username = $txtUser.Text.Trim()
        Password = $txtPass.Text
        OfflineThresholdHours = [int]$numOffline.Value
        SkipCertCheck = [bool]$chkSkipCert.Checked
        ProgressFile = $progressFile
        ResultFile = $resultFile
        SeedJson = ''
        IsRefresh = [bool]$IsRefresh
        TmpDir = $tmp
    }
    if ($Seed) { $cfg.SeedJson = ConvertTo-ApiResultJson -Result $Seed }
    $cfg | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $cfgPath -Encoding UTF8

    Set-ButtonsEnabled $false
    Set-Status "API $Phase starting (background)..."
    $job = Start-Job -ScriptBlock {
        param($scriptPath, $cfgPath)
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -ApiWorker $cfgPath -NoRelaunch
    } -ArgumentList $PSCommandPath, $cfgPath

    $script:BgJob = $job
    $script:BgProgressFile = $progressFile
    $script:BgResultFile = $resultFile
    $script:BgTmpDir = $tmp
    $script:BgPhase = $Phase
    $script:BgBaseUrl = $BaseUrl
    $script:BgIsRefresh = [bool]$IsRefresh
    $script:BgSeed = $Seed
    $script:BgTimer.Start()
}

function Complete-AsyncApiPhase {
    $script:BgTimer.Stop()
    $job = $script:BgJob
    $progressFile = $script:BgProgressFile
    $resultFile = $script:BgResultFile
    $tmp = $script:BgTmpDir
    $phase = $script:BgPhase
    $baseUrl = $script:BgBaseUrl
    $isRefresh = [bool]$script:BgIsRefresh
    $seed = $script:BgSeed
    $script:BgJob = $null

    try { if ($job) { Wait-Job -Job $job -Timeout 5 | Out-Null; Receive-Job -Job $job -ErrorAction SilentlyContinue | Out-Null; Remove-Job -Job $job -Force -ErrorAction SilentlyContinue } } catch { }
    $script:CursorDefault = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    $err = $null
    $result = $null
    try {
        if (Test-Path -LiteralPath $resultFile) {
            $json = Get-Content -LiteralPath $resultFile -Raw
            if (-not [string]::IsNullOrWhiteSpace($json)) {
                $obj = $json | ConvertFrom-Json
                if ($obj.PSObject.Properties['Error'] -and $obj.Error) {
                    $err = [string]$obj.Error
                } else {
                    $result = ConvertFrom-ApiResultJson -Json $json
                }
            } else { $err = 'Worker produced empty result.' }
        } else { $err = 'Worker produced no result file.' }
    } catch {
        $err = "Failed to read worker result: $($_.Exception.Message)"
    }

    try { if ($tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue } } catch { }

    if ($err) {
        Set-ButtonsEnabled $true
        $msg = $err
        if ($msg -notmatch '401|Unauthorized|HTTP ') {
            $msg = "Failed to connect to BigFix API.`r`n`r`n$msg`r`n`r`nUse Import as fallback, Load Demo, or open Settings."
        }
        # Offer session cache fallback
        $cached = Get-SessionCache
        if ($cached) {
            $answer = [System.Windows.Forms.MessageBox]::Show("$msg`r`n`r`nLoad last successful pull from cache ($($cached.SavedAt))?", 'API failed', 'YesNo', 'Warning')
            if ($answer -eq 'Yes') {
                Load-Records -Records (ConvertTo-ObjectArray -Value $cached.Failures) -SourceLabel "Cache: $($cached.SavedAt)"
                Set-Status "Loaded session cache after API error."
                return
            }
        }
        [void][System.Windows.Forms.MessageBox]::Show($msg, 'BigFix API connection failed', 'OK', 'Warning')
        Set-Status (($err -split "`r`n")[0])
        return
    }

    if ($phase -eq 'Inventory') {
        $failArr = ConvertTo-ObjectArray -Value $result.Failures
        $devArr = ConvertTo-ObjectArray -Value $result.Devices
        if ($devArr.Count -eq 0 -and $failArr.Count -eq 0) {
            Load-Records -Records @() -SourceLabel "BigFix API ($baseUrl) - no devices returned"
            Set-Status "Connected to $baseUrl but /api/computers returned 0 nodes."
            Set-ButtonsEnabled $true
            return
        }
        Load-Records -Records $failArr -SourceLabel "BigFix API ($baseUrl)"
        $sumI = Get-FailureSummary $failArr
        Set-Status "Devices loaded in $($result.ElapsedSec)s ($($sumI.Devices) devices) - scanning failures..."
        Start-AsyncApiPhase -Phase 'Failures' -BaseUrl $baseUrl -IsRefresh:$isRefresh -Seed $result
        return
    }

    # Failures phase complete
    $failArr2 = ConvertTo-ObjectArray -Value $result.Failures
    $logArr2 = ConvertTo-ObjectArray -Value $result.Log -AsString
    Load-Records -Records $failArr2 -SourceLabel "BigFix API ($baseUrl)"
    $sum = Get-FailureSummary $failArr2
    $tail = @($logArr2 | Select-Object -Last 5) -join ' | '
    Set-Status "API $($result.Version) in $($result.ElapsedSec)s - $($failArr2.Count) rows / $($sum.Devices) devices (healthy=$($sum.Healthy), offline=$($sum.Offline), fixlet=$($sum.FixletFailures)). $tail"
    Set-ButtonsEnabled $true
}

# ---- Events ----
$btnDemo.Add_Click({
    try {
        Set-ButtonsEnabled $false
        Set-Status 'Loading demo data...'
        $records = Get-DemoFailures
        Load-Records -Records $records -SourceLabel 'Demo data'
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Demo load failed', 'OK', 'Error')
        Set-Status "Demo load failed: $($_.Exception.Message)"
    } finally {
        Set-ButtonsEnabled $true
    }
})

$btnConnect.Add_Click({
    try {
        Start-AsyncApiConnect
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'BigFix API connection failed', 'OK', 'Warning')
        Set-Status "$($_.Exception.Message)"
    }
})

$btnImport.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = 'Import BigFix failure report'
    $ofd.Filter = 'Report files (*.csv;*.xml)|*.csv;*.xml|CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml|All files (*.*)|*.*'
    $ofd.InitialDirectory = Join-Path $script:ScriptRoot 'sample-data'
    if ($ofd.ShowDialog() -ne 'OK') { return }
    try {
        Set-ButtonsEnabled $false
        Set-Status "Importing $($ofd.FileName) ..."
        $ext = [System.IO.Path]::GetExtension($ofd.FileName).ToLower()
        if ($ext -eq '.xml') { $records = Import-FailureXml -Path $ofd.FileName } else { $records = Import-FailureCsv -Path $ofd.FileName }
        Load-Records -Records $records -SourceLabel "Import: $([System.IO.Path]::GetFileName($ofd.FileName))"
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Import failed', 'OK', 'Error')
        Set-Status "Import failed: $($_.Exception.Message)"
    } finally {
        Set-ButtonsEnabled $true
    }
})

$btnRefresh.Add_Click({
    try {
        if ($script:BgJob) { Set-Status 'Refresh already in progress...'; return }
        if ($script:LastSource -like 'BigFix API*') {
            Set-Status 'Refreshing from BigFix API (background)...'
            Start-AsyncApiConnect -IsRefresh
        } else {
            Invoke-CurrentSourceRefresh
        }
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Refresh failed', 'OK', 'Error')
        Set-Status "Refresh failed: $($_.Exception.Message)"
    }
})

$btnExport.Add_Click({
    if ($script:AllRecords.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No records to export.', 'Export', 'OK', 'Information')
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title = 'Export failed devices'
    $sfd.Filter = 'CSV files (*.csv)|*.csv'
    $sfd.FileName = "BigFix-FailedDevices-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    if ($script:AppSettings.ExportDirectory -and (Test-Path -LiteralPath $script:AppSettings.ExportDirectory)) {
        $sfd.InitialDirectory = $script:AppSettings.ExportDirectory
    }
    if ($sfd.ShowDialog() -ne 'OK') { return }
    try {
        $toExport = ConvertTo-ObjectArray -Value (Get-VisibleRecords)
        if ($toExport.Count -eq 0) { $toExport = ConvertTo-ObjectArray -Value $script:AllRecords }
        Export-FailuresCsv -Records $toExport -Path $sfd.FileName
        try {
            $script:AppSettings.ExportDirectory = Split-Path $sfd.FileName -Parent
            $null = Save-AppSettings
        } catch { }
        Set-Status "Exported $($toExport.Count) rows to $($sfd.FileName)"
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Export failed', 'OK', 'Error')
    }
})

$btnHtml.Add_Click({
    if ($script:AllRecords.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No records to export.', 'HTML Report', 'OK', 'Information')
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title = 'Export HTML report'
    $sfd.Filter = 'HTML files (*.html)|*.html'
    $sfd.FileName = "BigFix-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    if ($script:AppSettings.ExportDirectory -and (Test-Path -LiteralPath $script:AppSettings.ExportDirectory)) {
        $sfd.InitialDirectory = $script:AppSettings.ExportDirectory
    }
    if ($sfd.ShowDialog() -ne 'OK') { return }
    try {
        $toExport = ConvertTo-ObjectArray -Value (Get-VisibleRecords)
        if ($toExport.Count -eq 0) { $toExport = ConvertTo-ObjectArray -Value $script:AllRecords }
        Export-FailuresHtml -Records $toExport -Path $sfd.FileName -Title 'BigFix Failed Device Report'
        Set-Status "HTML report written: $($sfd.FileName)"
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'HTML export failed', 'OK', 'Error')
    }
})

$btnEmail.Add_Click({
    if ($script:AllRecords.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No records to email.', 'Email', 'OK', 'Information')
        return
    }
    $smtp = $script:AppSettings.Smtp
    if (-not $smtp -or -not [bool]$smtp.Enable) {
        $ans = [System.Windows.Forms.MessageBox]::Show("Email is not enabled.`r`n`r`nOpen Settings and check 'Enable email' (host/From/To)?", 'Email', 'YesNo', 'Question')
        if ($ans -eq 'Yes') { Show-SettingsDialog }
        return
    }
    try {
        Set-ButtonsEnabled $false
        $recs = ConvertTo-ObjectArray -Value (Get-VisibleRecords)
        if ($recs.Count -eq 0) { $recs = ConvertTo-ObjectArray -Value $script:AllRecords }
        Send-FailureDigest -Records $recs -Smtp $smtp
        [void][System.Windows.Forms.MessageBox]::Show('Digest email sent.', 'Email', 'OK', 'Information')
        Set-Status 'Email digest sent.'
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show("$($_.Exception.Message)`r`n`r`nConfigure SMTP under Settings (password is session-only).", 'Email failed', 'OK', 'Error')
        Set-Status "Email failed: $($_.Exception.Message)"
    } finally {
        Set-ButtonsEnabled $true
    }
})

$btnTicket.Add_Click({
    if ($script:AllRecords.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No records to post.', 'Ticket', 'OK', 'Information')
        return
    }
    $tkt = $script:AppSettings.Ticket
    if (-not $tkt -or -not [bool]$tkt.Enable) {
        $ans = [System.Windows.Forms.MessageBox]::Show("Ticket webhook is not enabled.`r`n`r`nOpen Settings and check 'Enable ticket webhook'?", 'Ticket', 'YesNo', 'Question')
        if ($ans -eq 'Yes') { Show-SettingsDialog }
        return
    }
    try {
        Set-ButtonsEnabled $false
        $scopeRecords = ConvertTo-ObjectArray -Value (Get-VisibleRecords)
        if ($scopeRecords.Count -eq 0) { $scopeRecords = ConvertTo-ObjectArray -Value $script:AllRecords }
        $failed = @($scopeRecords | Where-Object { "$($_.Category)" -ne 'Healthy' })
        if ($failed.Count -gt 0) { $scopeRecords = [object[]]$failed }
        if ($scopeRecords.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No failure rows to post.', 'Ticket', 'OK', 'Information')
            return
        }
        $resp = Send-TicketWebhook -Records $scopeRecords -Ticket $tkt
        [void][System.Windows.Forms.MessageBox]::Show("Webhook accepted ($($resp.StatusCode)).`r`n$($scopeRecords.Count) failure(s) posted.", 'Ticket', 'OK', 'Information')
        Set-Status "Ticket webhook HTTP $($resp.StatusCode) - $($scopeRecords.Count) records"
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show("$($_.Exception.Message)`r`n`r`nConfigure the webhook under Settings.", 'Ticket failed', 'OK', 'Error')
        Set-Status "Ticket failed: $($_.Exception.Message)"
    } finally {
        Set-ButtonsEnabled $true
    }
})

$btnSettings.Add_Click({
    try {
        $saved = Show-SettingsDialog -PassThru
        if ($saved) { Set-Status "Settings saved to $script:SettingsPath" }
        else { Set-Status 'Settings closed (not saved).' }
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show("$($_.Exception.Message)`r`n$($_.InvocationInfo.PositionMessage)", 'Settings failed', 'OK', 'Error')
        Set-Status "Settings failed: $($_.Exception.Message)"
    }
})

$btnRunbook.Add_Click({
    try {
        Show-RunbookDialog
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Runbook failed', 'OK', 'Error')
    }
})

$btnFixPreview.Add_Click({
    try {
        Set-ButtonsEnabled $false
        $null = Show-FixPlanPreview
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Preview failed', 'OK', 'Error')
        Write-FixLog "Preview error: $($_.Exception.Message)" 'Salmon'
    } finally {
        Set-ButtonsEnabled $true
    }
})

$btnFixClearLog.Add_Click({
    $rtbFixLog.Clear()
    $rtbFixLog.Text = "Remote Fix log cleared.`r`n`r`n"
})

$btnFixPreflight.Add_Click({
    try {
        Set-ButtonsEnabled $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:FixStopRequested = $true
        $records = ConvertTo-ObjectArray -Value (Get-FixScopeRecords)
        if (-not $records -or $records.Count -eq 0) { return }
        $minSev = Get-FixMinSeverityChoice
        $plan = Get-FixPlan -Records $records -MinSeverity $minSev
        if ($plan.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No applicable solutions for the current scope/severity.', 'Remote Fix', 'OK', 'Information')
            return
        }
        $cred = Resolve-FixCredential
        $script:FixRunning = $true
        $btnFixStop.Enabled = $true
        $script:FixStopRequested = $false
        Write-FixLog "---- Pre-flight $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ----" 'Yellow'
        $pf = Test-FixPreFlight -Plan $plan -Credential $cred -TimeoutSec ([int]$numFixTimeout.Value) -Log { param($m) Write-FixLog $m 'DarkGray' }
        $okN = @($pf | Where-Object { $_.Status -eq 'OK' }).Count
        $failN = @($pf | Where-Object { $_.Status -ne 'OK' }).Count
        Set-Status "WinRM pre-flight: OK=$okN Failed=$failN"
        if ($failN -gt 0 -and $okN -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show("All $failN host(s) failed WinRM test.`r`nCheck firewall, WinRM, and credentials.", 'Pre-flight', 'OK', 'Warning')
        }
    } catch {
        Write-FixLog "Pre-flight error: $($_.Exception.Message)" 'Salmon'
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Pre-flight failed', 'OK', 'Error')
    } finally {
        $script:FixRunning = $false
        $script:FixStopRequested = $false
        $btnFixStop.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        Set-ButtonsEnabled $true
    }
})

$btnFixExportLog.Add_Click({
    try {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
        $sfd.FileName = ("fix-log-{0:yyyyMMdd-HHmmss}.txt" -f (Get-Date))
        $sfd.InitialDirectory = if ($script:AppSettings.ExportDirectory -and (Test-Path -LiteralPath $script:AppSettings.ExportDirectory)) { $script:AppSettings.ExportDirectory } else { $env:USERPROFILE }
        if ($sfd.ShowDialog() -ne 'OK') { return }
        Set-Content -LiteralPath $sfd.FileName -Value $rtbFixLog.Text -Encoding UTF8
        Set-Status "Fix log saved: $($sfd.FileName)"
        Write-FixLog "Log exported: $($sfd.FileName)" 'LightGreen'
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Export log failed', 'OK', 'Error')
    }
})

$btnFixStop.Add_Click({
    $script:FixStopRequested = $true
    Write-FixLog 'Stop requested - finishing current host then stopping...' 'Khaki'
    $btnFixStop.Enabled = $false
    [System.Windows.Forms.Application]::DoEvents()
})

$btnFixRetry.Add_Click({
    try {
        if (-not $script:LastFixResults -or @($script:LastFixResults).Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No previous fix run to retry. Run Fix first.', 'Remote Fix', 'OK', 'Information')
            return
        }
        $failed = @($script:LastFixResults | Where-Object { $_.Status -in @('Failed', 'WinRM-Failed') })
        if ($failed.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('Last run had no Failed / WinRM-Failed entries to retry.', 'Remote Fix', 'OK', 'Information')
            return
        }
        Set-ButtonsEnabled $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:FixStopRequested = $false
        $plan = Get-FixScopedPlan -FromRetry
        if ($plan.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('Could not rebuild plan from last run (plan cache missing). Run Fix again.', 'Remote Fix', 'OK', 'Warning')
            return
        }
        $whatIf = [bool]$chkFixWhatIf.Checked
        $mode = if ($whatIf) { 'DRY RUN RETRY' } else { 'RETRY FAILED via WinRM' }
        $preview = Format-FixPlanPreview -Plan $plan
        $confirm = @"
$mode - $($plan.Count) failed solution(s) from last run?

$preview
"@
        $btn = if ($whatIf) { [System.Windows.Forms.MessageBoxButtons]::OK } else { [System.Windows.Forms.MessageBoxButtons]::YesNo }
        $icon = if ($whatIf) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
        $answer = [System.Windows.Forms.MessageBox]::Show($confirm, "Remote Fix - $mode", $btn, $icon)
        if (-not $whatIf -and $answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-FixLog 'User cancelled retry.' 'Gray'
            return
        }
        $cred = Resolve-FixCredential
        $script:FixRunning = $true
        $btnFixStop.Enabled = $true
        Write-FixLog "---- Retry $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ($mode) ----" 'Yellow'
        $results = Invoke-RemoteFixPlan -Plan $plan -Credential $cred -WhatIf:$whatIf -TimeoutSec ([int]$numFixTimeout.Value) -Log { param($m) Write-FixLog $m 'DarkGray' }
        Complete-FixRun -Results $results -Mode $mode
    } catch {
        Write-FixLog "Retry error: $($_.Exception.Message)" 'Salmon'
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Retry failed', 'OK', 'Error')
    } finally {
        $script:FixRunning = $false
        $script:FixStopRequested = $false
        $btnFixStop.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        Set-ButtonsEnabled $true
    }
})

$btnFix.Add_Click({
    try {
        Set-ButtonsEnabled $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:FixStopRequested = $false

        $records = ConvertTo-ObjectArray -Value (Get-FixScopeRecords)
        if (-not $records -or $records.Count -eq 0) { return }

        $minSev = Get-FixMinSeverityChoice
        $plan = @(Get-FixPlan -Records $records -MinSeverity $minSev | ForEach-Object {
            $clone = $_ | Select-Object *
            $clone.DeviceName = Resolve-FixHostTarget -DeviceName $_.DeviceName
            $clone.Remotable = (Test-RemotableDeviceName -Name $clone.DeviceName)
            $clone
        })
        if ($plan.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No applicable solutions for the current scope/severity.', 'Remote Fix', 'OK', 'Information')
            return
        }
        $script:LastFixPlan = $plan

        $preview = Format-FixPlanPreview -Plan $plan
        $whatIf = [bool]$chkFixWhatIf.Checked
        $mode = if ($whatIf) { 'DRY RUN' } else { 'EXECUTE via WinRM' }
        $confirm = @"
$mode - run these solutions?

$preview

Only solutions matching each error rule/category are included.
Healthy rows are skipped. Manual rows are listed but not executed remotely.
Timeout per host: $([int]$numFixTimeout.Value)s
"@
        $btn = if ($whatIf) { [System.Windows.Forms.MessageBoxButtons]::OK }
               else { [System.Windows.Forms.MessageBoxButtons]::YesNo }
        $icon = if ($whatIf) { [System.Windows.Forms.MessageBoxIcon]::Information }
                else { [System.Windows.Forms.MessageBoxIcon]::Warning }
        $answer = [System.Windows.Forms.MessageBox]::Show($confirm, "Remote Fix - $mode", $btn, $icon)
        if (-not $whatIf -and $answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-FixLog 'User cancelled execution.' 'Gray'
            return
        }

        $cred = Resolve-FixCredential

        $script:FixRunning = $true
        $btnFixStop.Enabled = $true
        Write-FixLog "---- Execute $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ($mode) ----" 'Yellow'
        $results = Invoke-RemoteFixPlan `
            -Plan $plan `
            -Credential $cred `
            -WhatIf:$whatIf `
            -TimeoutSec ([int]$numFixTimeout.Value) `
            -Log { param($m) Write-FixLog $m 'DarkGray' }

        Complete-FixRun -Results $results -Mode $mode
    } catch {
        Write-FixLog "Fix error: $($_.Exception.Message)" 'Salmon'
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Remote Fix failed', 'OK', 'Error')
    } finally {
        $script:FixRunning = $false
        $script:FixStopRequested = $false
        $btnFixStop.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        Set-ButtonsEnabled $true
    }
})

$btnAbout.Add_Click({
    $msg = @"
BigFix Failed Device Dashboard v$script:AppVersion

Shows failed BigFix devices with fixlet/compliance error details
and maps each error to a root cause + remediation steps.

  Data sources:
  1. BigFix REST API  (https://<server>:52311) - lists all devices
  2. CSV / XML import fallback
  3. Built-in demo data / session cache

  New in 2.0:
  - Async Connect API (background job - UI stays responsive)
  - Charts tab (category/severity/top devices/rules)
  - Device view + failure timeline
  - HTML executive report, JSON export, Excel-friendly CSV
  - Email digest + ticket webhook (Settings)
  - Auto-refresh + tray balloon alerts
  - Session cache offline fallback
  - Settings persistence (passwords never saved)
  - 12 extra root-cause rules + Remote Fix scripts
  - Headless export: -ExportHtml/-ExportCsv/-ExportJson/-Email

  Runbook button: 10 day-to-day server/client topics.

  Remote Fix tab (last tab): scoped WinRM solutions.

  API login: operator with "Can use REST API".
  Browser check: https://server:52311/api/help

Headless examples:
  powershell -File BigFix-FailureDashboard.ps1 -ExportHtml C:\out\bf.html -UseCache
  powershell -File BigFix-FailureDashboard.ps1 -SelfTest
"@
    [void][System.Windows.Forms.MessageBox]::Show($msg, 'About', 'OK', 'Information')
})

$cmbCategory.Add_SelectedIndexChanged({ Persist-UiSettings; Apply-Filters })
$cmbSeverity.Add_SelectedIndexChanged({ Persist-UiSettings; Apply-Filters })
$txtSearch.Add_TextChanged({ Apply-Filters })

$chkDeviceView.Add_CheckedChanged({
    try {
        Update-GridFromRecords
        Persist-UiSettings
    } catch { }
})

$numAutoRefresh.Add_ValueChanged({ Update-AutoRefreshTimer; Persist-UiSettings })

$btnClearFilter.Add_Click({
    $cmbCategory.SelectedItem = 'All'
    $cmbSeverity.SelectedItem = 'All'
    $txtSearch.Text = ''
    Apply-Filters
})

$grid.Add_SelectionChanged({ Show-SelectedDetails })
# Cached column indexes for large-grid CellFormatting
$script:ColSeverity = -1
$script:ColCategory = -1
$grid.Add_ColumnHeaderMouseClick({
    # refresh cached indexes after reorder
    try {
        if ($grid.Columns['Severity']) { $script:ColSeverity = $grid.Columns['Severity'].Index }
        if ($grid.Columns['Category']) { $script:ColCategory = $grid.Columns['Category'].Index }
    } catch { }
})
$grid.Add_CellFormatting({
    param($sender, $e)
    if ($script:ColSeverity -lt 0) {
        try {
            if ($grid.Columns['Severity']) { $script:ColSeverity = $grid.Columns['Severity'].Index }
            if ($grid.Columns['Category']) { $script:ColCategory = $grid.Columns['Category'].Index }
        } catch { }
    }
    if ($e.ColumnIndex -eq $script:ColSeverity -and $script:ColSeverity -ge 0) {
        $sev = [string]$grid.Rows[$e.RowIndex].Cells['Severity'].Value
        switch -Regex ($sev) {
            '^Critical' { $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkRed }
            '^High'     { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 60, 0) }
            '^Medium'   { $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkGoldenrod }
            default     { $e.CellStyle.ForeColor = [System.Drawing.Color]::Gray }
        }
        return
    }
    if ($e.ColumnIndex -eq $script:ColCategory -and $script:ColCategory -ge 0) {
        $cat = [string]$grid.Rows[$e.RowIndex].Cells['Category'].Value
        switch ($cat) {
            'Offline'            { $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkRed }
            'ComplianceFailure'  { $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkViolet }
            'SoftwareDeployment' { $e.CellStyle.ForeColor = [System.Drawing.Color]::Teal }
            'Healthy'            { $e.CellStyle.ForeColor = [System.Drawing.Color]::ForestGreen }
            default              { $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkOrange }
        }
    }
})

# ---- Background API job poller / auto-refresh / tray ----
$script:BgJob = $null
$script:BgProgressFile = $null
$script:BgResultFile = $null
$script:BgTmpDir = $null
$script:BgPhase = ''
$script:BgBaseUrl = ''
$script:BgIsRefresh = $false
$script:BgSeed = $null
$script:PrevCriticalCount = 0
$script:CachedSummary = $null
$script:SummaryDirty = $true
$script:SmtpPassword = $env:BIGFIX_SMTP_PASSWORD

$script:BgTimer = New-Object System.Windows.Forms.Timer
$script:BgTimer.Interval = 500
$script:BgTimer.Add_Tick({
    if (-not $script:BgJob) { $script:BgTimer.Stop(); return }
    try {
        if ($script:BgProgressFile -and (Test-Path -LiteralPath $script:BgProgressFile)) {
            $p = Get-Content -LiteralPath $script:BgProgressFile -Raw -ErrorAction SilentlyContinue
            if ($p) { Set-Status "API $($script:BgPhase): $($p.Trim())" }
        }
        $state = $script:BgJob.State
        if ($state -in @('Completed', 'Failed', 'Blocked', 'Stopped')) {
            Complete-AsyncApiPhase
        }
    } catch {
        try { $script:BgTimer.Stop() } catch { }
        try { Complete-AsyncApiPhase } catch { }
    }
})

$script:AutoRefreshTimer = New-Object System.Windows.Forms.Timer
$script:AutoRefreshTimer.Add_Tick({
    try {
        if ($script:BgJob) { return }
        if ($script:AllRecords.Count -eq 0) { return }
        Invoke-CurrentSourceRefresh
    } catch { }
})

try {
    $script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $script:NotifyIcon.Icon = [System.Drawing.SystemIcons]::Information
    $script:NotifyIcon.Text = "BigFix Failure Dashboard v$script:AppVersion"
    $script:NotifyIcon.Visible = $true
    $script:NotifyIcon.Add_DoubleClick({
        try {
            $script:MainForm.WindowState = 'Normal'
            $script:MainForm.BringToFront()
            $script:MainForm.Activate()
        } catch { }
    })
} catch {
    $script:NotifyIcon = $null
}

$form.Add_Shown({
    try {
        $splitMain.Panel2MinSize = 140
        $target = [int]($splitMain.Height * 0.55)
        if ($target -lt 180) { $target = 180 }
        $max = $splitMain.Height - 150
        if ($target -gt $max -and $max -gt 180) { $target = $max }
        $splitMain.SplitterDistance = $target
        if ($script:ChartsAvailable) {
            try {
                $splitCharts.SplitterDistance = [Math]::Max(200, [int]($splitCharts.Width / 2))
                $rowTop.SplitterDistance = [Math]::Max(150, [int]($rowTop.Width / 2))
                $rowBot.SplitterDistance = [Math]::Max(150, [int]($rowBot.Width / 2))
            } catch { }
        }
    } catch { }

    try {
    Initialize-DataTable
    Apply-UiSettingsFromAppSettings
    if ($script:AllRecords.Count -eq 0) {
        if ($UseCache) {
            $cached = Get-SessionCache
            if ($cached) {
                Load-Records -Records (ConvertTo-ObjectArray -Value $cached.Failures) -SourceLabel "Cache: $($cached.SavedAt)"
            } else {
                Load-Records -Records (Get-DemoFailures) -SourceLabel 'Demo data (auto-loaded)'
            }
        } else {
            Load-Records -Records (Get-DemoFailures) -SourceLabel 'Demo data (auto-loaded)'
        }
    }
    Update-AutoRefreshTimer
    } catch {
        $err = $_
        $pos = $err.InvocationInfo.PositionMessage
        try { [System.IO.File]::AppendAllText((Join-Path $env:TEMP 'BigFixDashboard-UIErrors.log'), "$(Get-Date -Format o) Add_Shown: $($err.Exception.Message)`r`n$pos`r`n$($err.ScriptStackTrace)`r`n`r`n") } catch { }
        try { $lblStatus.Text = "Startup error: $($err.Exception.Message)" } catch { }
    }
})

$form.Add_FormClosing({
    Persist-UiSettings
    try { if ($script:NotifyIcon) { $script:NotifyIcon.Visible = $false; $script:NotifyIcon.Dispose() } } catch { }
    try { if ($script:BgJob) { Stop-Job -Job $script:BgJob -ErrorAction SilentlyContinue; Remove-Job -Job $script:BgJob -Force -ErrorAction SilentlyContinue } } catch { }
    try { if ($script:BgTimer) { $script:BgTimer.Stop() } } catch { }
    try { if ($script:AutoRefreshTimer) { $script:AutoRefreshTimer.Stop() } } catch { }
})

# ---- Launch ----
Initialize-DataTable
[void]$form.ShowDialog()
