# Probe whether this PC can host a PowerShell port of xlvbesync.
#
# Starts one hidden Excel that this script owns, exercises the VBA project
# object model on a new unsaved workbook, then quits that instance. An Excel
# you already have open is left alone, and nothing is saved.
#
# If the script hangs, Excel has a modal dialog open. Press Esc in that
# window, or press Ctrl+C here.
#
# From a normal user terminal:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-XlvbesyncCapability.ps1
#
# If a machine policy rejects -File, this still cannot escape Constrained
# Language, but you can run it inside an already-open host:
#   Get-Content .\scripts\Test-XlvbesyncCapability.ps1 -Raw | Invoke-Expression
#
# Exit codes:
#   0  live VBProject read/write works, and Add-Type plus FileSystemWatcher work
#   1  live VBProject read/write works, but a secondary capability does not
#   3  only the cscript fallback can drive VBProject
#   2  VBProject automation is blocked

$ErrorActionPreference = 'Continue'

$script:Results = @()
$script:Lines = @()
$script:TempDir = $null
$script:OwnedPids = @()
$script:PsExcel = $false
$script:PsVbomRead = $false
$script:PsVbomWrite = $false
$script:PsImport = $false
$script:PsDocument = $false
$script:AddTypeOk = $false
$script:WatcherOk = $false
$script:CscriptWrite = $false
$script:TimingMedian = $null

function Add-ProbeLine {
    param([string]$Text)
    $script:Lines += $Text
    Write-Host $Text
}

function Add-ProbeResult {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][ValidateSet('ok', 'warn', 'fail')]$Status,
        [string]$Detail = '',
        [string]$Remedy = ''
    )

    $script:Results += [pscustomobject]@{
        Name   = $Name
        Status = $Status
        Detail = $Detail
        Remedy = $Remedy
    }

    $tag = switch ($Status) {
        'ok'   { '[ ok ]' }
        'warn' { '[warn]' }
        'fail' { '[fail]' }
    }
    $color = switch ($Status) {
        'ok'   { 'Green' }
        'warn' { 'Yellow' }
        'fail' { 'Red' }
    }
    $line = '{0} {1}' -f $tag, $Name
    if ($Detail) { $line = '{0} - {1}' -f $line, $Detail }
    $script:Lines += $line
    Write-Host $line -ForegroundColor $color
    if ($Remedy) {
        $hint = '       {0}' -f $Remedy
        $script:Lines += $hint
        Write-Host $hint
    }
}

function Add-ProbeSection {
    param([string]$Title)
    $script:Lines += ''
    $script:Lines += ('== {0} ==' -f $Title)
    Write-Host ''
    Write-Host ('== {0} ==' -f $Title)
}

function Get-SignedHResult {
    param($ErrorRecord)
    # Prefer the inner COM ErrorCode. A wrapper's own HResult is often the
    # generic .NET 0x80131500, which would hide MK_E_UNAVAILABLE and friends.
    try {
        $ex = $null
        if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
            $ex = $ErrorRecord.Exception
        }
        elseif ($ErrorRecord -is [System.Exception]) {
            $ex = $ErrorRecord
        }
        $fallback = $null
        while ($null -ne $ex) {
            $codeProp = $ex.PSObject.Properties['ErrorCode']
            if ($null -ne $codeProp -and $codeProp.Value -is [int] -and [int]$codeProp.Value -ne 0) {
                return [int]$codeProp.Value
            }
            $hrProp = $ex.PSObject.Properties['HResult']
            if ($null -ne $hrProp -and $hrProp.Value -is [int]) {
                $code = [int]$hrProp.Value
                # 0x80131500 is a .NET wrapper, not the COM failure underneath.
                if ($code -ne 0 -and $code -ne -2146233088) { $fallback = $code }
            }
            $ex = $ex.InnerException
        }
        return $fallback
    }
    catch {
        return $null
    }
}

function Get-ErrorDetail {
    param($ErrorRecord)
    try {
        if ($null -eq $ErrorRecord) { return 'unknown error' }
        $msg = [string]$ErrorRecord.Exception.Message
        if ([string]::IsNullOrWhiteSpace($msg) -and $null -ne $ErrorRecord.Exception.InnerException) {
            $msg = [string]$ErrorRecord.Exception.InnerException.Message
        }
        $hr = Get-SignedHResult $ErrorRecord
        if ($null -ne $hr) {
            $name = switch ($hr) {
                -2147418111 { 'RPC_E_CALL_REJECTED' }
                -2147417846 { 'RPC_E_SERVERCALL_RETRYLATER' }
                -2146822220 { 'VBA_E_ACCESS_DENIED' }
                -2146777998 { 'E_FAIL_AUTOMATION' }
                -2147024891 { 'E_ACCESSDENIED' }
                -2147352567 { 'DISP_E_EXCEPTION' }
                -2147221021 { 'MK_E_UNAVAILABLE' }
                default { $null }
            }
            $hex = '0x{0:X8}' -f ($hr -band 0xFFFFFFFF)
            if ($name) { $msg = '{0} ({1} {2})' -f $msg, $name, $hex }
            else { $msg = '{0} (HRESULT {1})' -f $msg, $hex }
        }
        return (($msg -replace '\s+', ' ').Trim())
    }
    catch {
        return 'could not format the error'
    }
}

function Invoke-ComRetry {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)
    # Same transient HRESULTs as xlvbesync.core.retry: Excel busy, or a dialog
    # is open. Anything else is reported immediately.
    $delay = 200
    for ($i = 1; $i -le 6; $i++) {
        try {
            return (& $Action)
        }
        catch {
            $hr = Get-SignedHResult $_
            $transient = ($hr -eq -2147418111 -or $hr -eq -2147417846)
            if (-not $transient -or $i -eq 6) { throw }
            Start-Sleep -Milliseconds $delay
            if ($delay -lt 2000) { $delay = $delay * 2 }
        }
    }
}

function Get-ExcelPids {
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Id)
}

function Get-NewExcelPids {
    param([int[]]$Before)
    $after = Get-ExcelPids
    @($after | Where-Object { $Before -notcontains $_ })
}

function Release-ComObject {
    param($Object)
    if ($null -eq $Object) { return }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) }
    catch { }
}

function Stop-OwnedExcel {
    # Quit only the process this script started. Never Stop-Process an Excel
    # that was already running: CreateObject usually starts a new process, but
    # if it did not, $OwnedPids stays empty and Quit is skipped.
    if ($script:OwnedPids.Count -eq 0) { return }
    $deadline = (Get-Date).AddSeconds(3)
    $still = @($script:OwnedPids)
    do {
        $alive = @()
        foreach ($procId in $still) {
            if (Get-Process -Id $procId -ErrorAction SilentlyContinue) { $alive += $procId }
        }
        $still = $alive
        if ($still.Count -eq 0) { return }
        if ((Get-Date) -ge $deadline) { break }
        Start-Sleep -Milliseconds 300
    } while ($true)

    foreach ($procId in $still) {
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    }
    Add-ProbeResult 'Excel cleanup' 'warn' (
        'Quit left PID {0} running; stopped only the instance this script started.' -f ($still -join ', ')
    )
}

function Get-ReportPath {
    $name = 'xlvbesync-capability-report.txt'
    $candidates = @()
    if ($PSScriptRoot) {
        $candidates += (Join-Path $PSScriptRoot $name)
    }
    $candidates += (Join-Path (Join-Path (Get-Location).Path 'scripts') $name)
    $candidates += (Join-Path $env:TEMP $name)

    foreach ($candidate in $candidates) {
        $dir = Split-Path -Parent $candidate
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        $probe = Join-Path $dir '.xlvbesync-probe-write'
        try {
            Set-Content -LiteralPath $probe -Value 'ok' -Encoding ASCII -ErrorAction Stop
            Remove-Item -LiteralPath $probe -Force -ErrorAction Stop
            return $candidate
        }
        catch {
            continue
        }
    }
    return (Join-Path $env:TEMP $name)
}

function Get-ExeMachine {
    param([string]$Path)
    $stream = $null
    $reader = $null
    try {
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )
        $reader = New-Object System.IO.BinaryReader $stream
        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset + 4
        $machine = $reader.ReadUInt16()
        switch ($machine) {
            0x14C { return 'x86' }
            0x8664 { return 'x64' }
            default { return ('machine 0x{0:X}' -f $machine) }
        }
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        elseif ($null -ne $stream) { $stream.Dispose() }
    }
}

function Test-HostEnvironment {
    Add-ProbeSection 'Host'

    $edition = [string]$PSVersionTable.PSEdition
    if ([string]::IsNullOrWhiteSpace($edition)) { $edition = 'Desktop' }
    $hostName = if ($edition -eq 'Core') { 'pwsh' } else { 'Windows PowerShell' }
    $clr = [string]$PSVersionTable.CLRVersion
    $bitness = if ([Environment]::Is64BitProcess) { '64-bit process' } else { '32-bit process' }
    Add-ProbeResult 'PowerShell host' 'ok' (
        '{0} {1}, CLR {2}, {3}' -f $hostName, $PSVersionTable.PSVersion, $clr, $bitness
    )

    $mode = [string]$ExecutionContext.SessionState.LanguageMode
    if ($mode -eq 'FullLanguage') {
        Add-ProbeResult 'Language mode' 'ok' $mode
    }
    else {
        Add-ProbeResult 'Language mode' 'fail' $mode (
            'Constrained Language blocks New-Object -ComObject, Add-Type, and most .NET calls. ExecutionPolicy Bypass does not turn this off.'
        )
    }

    $lock = [string]$env:__PSLockdownPolicy
    $lockLabel = switch ($lock) {
        '' { 'not set' }
        '4' { '4 (WDAC enforced)' }
        '8' { '8 (WDAC audit)' }
        default { $lock }
    }
    $lockStatus = if ($lock -eq '4') { 'fail' } elseif ($lock -eq '8') { 'warn' } else { 'ok' }
    Add-ProbeResult 'WDAC lockdown policy' $lockStatus ('__PSLockdownPolicy={0}' -f $lockLabel)

    try {
        $policies = @(Get-ExecutionPolicy -List -ErrorAction Stop)
        $text = ($policies | ForEach-Object { '{0}={1}' -f $_.Scope, $_.ExecutionPolicy }) -join '; '
        $machine = $policies | Where-Object { $_.Scope -eq 'MachinePolicy' } | Select-Object -First 1
        $userPol = $policies | Where-Object { $_.Scope -eq 'UserPolicy' } | Select-Object -First 1
        $blocked = @('Restricted', 'AllSigned')
        $policyHit = ($null -ne $machine -and $blocked -contains [string]$machine.ExecutionPolicy) -or
            ($null -ne $userPol -and $blocked -contains [string]$userPol.ExecutionPolicy)
        if ($policyHit) {
            Add-ProbeResult 'Execution policy' 'warn' $text (
                'MachinePolicy or UserPolicy is Restricted or AllSigned. -ExecutionPolicy Bypass will not override it. Paste the script into an open host with Get-Content -Raw | Invoke-Expression if -File is rejected.'
            )
        }
        else {
            Add-ProbeResult 'Execution policy' 'ok' $text
        }
    }
    catch {
        Add-ProbeResult 'Execution policy' 'warn' (Get-ErrorDetail $_)
    }
}

function Test-PolicyPresence {
    Add-ProbeSection 'AppLocker and WDAC'

    $srp = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2'
    try {
        if (-not (Test-Path -LiteralPath $srp)) {
            Add-ProbeResult 'AppLocker policy' 'ok' 'no SrpV2 policy key'
        }
        else {
            $kids = @(Get-ChildItem -LiteralPath $srp -ErrorAction Stop)
            $names = @($kids | ForEach-Object { $_.PSChildName })
            $scriptRules = 0
            $scriptKey = Join-Path $srp 'Script'
            if (Test-Path -LiteralPath $scriptKey) {
                $scriptRules = @(Get-ChildItem -LiteralPath $scriptKey -ErrorAction SilentlyContinue).Count
            }
            $detail = 'collections: {0}' -f ($(if ($names.Count) { $names -join ', ' } else { '(empty)' }))
            if ($names -contains 'Script') {
                $detail = '{0}; Script rules={1}' -f $detail, $scriptRules
                Add-ProbeResult 'AppLocker policy' 'warn' $detail (
                    'Script rules are present. If Language mode above is FullLanguage, they are not forcing Constrained Language for this host.'
                )
            }
            else {
                Add-ProbeResult 'AppLocker policy' 'ok' $detail
            }
        }
    }
    catch {
        Add-ProbeResult 'AppLocker policy' 'warn' (Get-ErrorDetail $_)
    }

    $safer = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers'
    try {
        if (Test-Path -LiteralPath $safer) {
            $props = Get-ItemProperty -LiteralPath $safer -ErrorAction Stop
            $level = $null
            if (@($props.PSObject.Properties.Name) -contains 'DefaultLevel') {
                $level = $props.DefaultLevel
            }
            if ($null -ne $level -and "$level" -ne '') {
                Add-ProbeResult 'SAFER CodeIdentifiers' 'warn' ('DefaultLevel={0}' -f $level)
            }
            else {
                Add-ProbeResult 'SAFER CodeIdentifiers' 'ok' 'policy key present, DefaultLevel not set'
            }
        }
        else {
            Add-ProbeResult 'SAFER CodeIdentifiers' 'ok' 'policy key not present'
        }
    }
    catch {
        Add-ProbeResult 'SAFER CodeIdentifiers' 'warn' (Get-ErrorDetail $_)
    }

    $dg = $null
    $dgError = $null
    $job = $null
    try {
        $job = Start-Job -ScriptBlock {
            Get-CimInstance -Namespace root\Microsoft\Windows\DeviceGuard -ClassName Win32_DeviceGuard
        }
        if (Wait-Job -Job $job -Timeout 15) {
            if ($job.State -eq 'Completed') {
                $dg = Receive-Job -Job $job -ErrorAction SilentlyContinue
            }
            else {
                $reason = $job.ChildJobs |
                    ForEach-Object { $_.JobStateInfo.Reason } |
                    Select-Object -First 1
                $dgError = if ($reason) { [string]$reason.Message } else { 'Device Guard query failed' }
            }
        }
        else {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            $dgError = 'timed out after 15s'
        }
    }
    catch {
        $dgError = Get-ErrorDetail $_
    }
    finally {
        if ($null -ne $job) {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
    }

    if ($null -eq $dg) {
        Add-ProbeResult 'Device Guard / WDAC' 'warn' ($(if ($dgError) { $dgError } else { 'no instance returned' }))
    }
    else {
        $row = $dg | Select-Object -First 1
        $enforce = {
            switch ([string]$args[0]) {
                '0' { 'Off' }
                '1' { 'Audit' }
                '2' { 'Enforced' }
                default { [string]$args[0] }
            }
        }
        $ci = & $enforce $row.CodeIntegrityPolicyEnforcementStatus
        $um = & $enforce $row.UsermodeCodeIntegrityPolicyEnforcementStatus
        $detail = 'code integrity={0}; user-mode={1}' -f $ci, $um
        $status = if ($um -eq 'Enforced' -or $um -eq 'Audit') { 'warn' } else { 'ok' }
        $remedy = ''
        if ($um -eq 'Enforced' -or $um -eq 'Audit') {
            $remedy = 'User-mode code integrity is what puts PowerShell into Constrained Language. Language mode is the result that matters.'
        }
        Add-ProbeResult 'Device Guard / WDAC' $status $detail $remedy
    }
}

function Get-OfficeVersions {
    $versions = @('14.0', '15.0', '16.0')
    foreach ($root in @(
            'HKCU:\Software\Microsoft\Office',
            'HKLM:\SOFTWARE\Microsoft\Office',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office'
        )) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        try {
            foreach ($child in @(Get-ChildItem -LiteralPath $root -ErrorAction Stop)) {
                if ($child.PSChildName -match '^\d+\.\d+$') {
                    $versions += $child.PSChildName
                }
            }
        }
        catch { }
    }
    @($versions | Select-Object -Unique)
}

function Test-OfficeRegistration {
    Add-ProbeSection 'Office and Trust Center'

    $excelPath = $null
    foreach ($key in @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'
        )) {
        if (-not (Test-Path -LiteralPath $key)) { continue }
        try {
            $item = Get-ItemProperty -LiteralPath $key -ErrorAction Stop
            $candidate = [string]$item.'(default)'
            if ([string]::IsNullOrWhiteSpace($candidate)) { $candidate = [string]$item.Path }
            if ($candidate -and (Test-Path -LiteralPath $candidate)) {
                $excelPath = $candidate
                break
            }
            if ($item.Path -and (Test-Path -LiteralPath (Join-Path $item.Path 'EXCEL.EXE'))) {
                $excelPath = Join-Path $item.Path 'EXCEL.EXE'
                break
            }
        }
        catch { }
    }

    $platform = $null
    $c2rVersion = $null
    $c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path -LiteralPath $c2r) {
        try {
            $c2rProps = Get-ItemProperty -LiteralPath $c2r -ErrorAction Stop
            $platform = [string]$c2rProps.Platform
            $c2rVersion = [string]$c2rProps.VersionToReport
        }
        catch { }
    }

    if ($excelPath) {
        $machine = $null
        try { $machine = Get-ExeMachine -Path $excelPath } catch { $machine = $null }
        $bits = @()
        if ($c2rVersion) { $bits += ('ClickToRun {0}' -f $c2rVersion) }
        if ($platform) { $bits += ('platform {0}' -f $platform) }
        if ($machine) { $bits += $machine }
        $bits += $excelPath
        Add-ProbeResult 'Excel installation' 'ok' ($bits -join '; ')
    }
    else {
        Add-ProbeResult 'Excel installation' 'fail' 'excel.exe was not found via App Paths' (
            'Desktop Excel has to be installed. Office on the web has no VBProject API.'
        )
    }

    $progId = 'Registry::HKEY_CLASSES_ROOT\Excel.Application'
    if (Test-Path -LiteralPath $progId) {
        Add-ProbeResult 'Excel.Application ProgID' 'ok' 'registered'
    }
    else {
        Add-ProbeResult 'Excel.Application ProgID' 'fail' 'HKEY_CLASSES_ROOT\Excel.Application is missing' (
            'COM activation of Excel will fail until the desktop application is registered.'
        )
    }

    $hives = @(
        @{ Label = 'HKCU user';   Root = 'HKCU:\Software\Microsoft\Office' },
        @{ Label = 'HKCU policy'; Root = 'HKCU:\Software\Policies\Microsoft\Office' },
        @{ Label = 'HKLM policy'; Root = 'HKLM:\SOFTWARE\Policies\Microsoft\Office' }
    )
    $versions = Get-OfficeVersions
    $accessLines = @()
    $warningLines = @()
    $policyOff = $false
    $userOn = $false
    $userOff = $false

    foreach ($ver in $versions) {
        foreach ($hive in $hives) {
            $path = '{0}\{1}\Excel\Security' -f $hive.Root, $ver
            if (-not (Test-Path -LiteralPath $path)) { continue }
            try {
                $props = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
            }
            catch {
                continue
            }
            $names = @($props.PSObject.Properties.Name)
            if ($names -contains 'AccessVBOM') {
                $value = [int]$props.AccessVBOM
                $accessLines += ('{0} Office {1} AccessVBOM={2}' -f $hive.Label, $ver, $value)
                if ($hive.Label -like '*policy*' -and $value -eq 0) { $policyOff = $true }
                if ($hive.Label -eq 'HKCU user' -and $value -eq 1) { $userOn = $true }
                if ($hive.Label -eq 'HKCU user' -and $value -eq 0) { $userOff = $true }
            }
            if ($names -contains 'VBAWarnings') {
                $warnValue = [int]$props.VBAWarnings
                $warnLabel = switch ($warnValue) {
                    1 { 'enable all' }
                    2 { 'disable with notification' }
                    3 { 'disable except signed' }
                    4 { 'disable all' }
                    default { 'unrecognised' }
                }
                $warningLines += ('{0} Office {1} VBAWarnings={2} ({3})' -f $hive.Label, $ver, $warnValue, $warnLabel)
            }
        }
    }

    if ($accessLines.Count -eq 0) {
        Add-ProbeResult 'AccessVBOM registry' 'warn' 'setting not found for Office 14.0/15.0/16.0 or any other installed version' (
            'Enable Excel, File, Options, Trust Center, Trust Center Settings, Macro Settings, Trust access to the VBA project object model. The live VBProject check below is definitive.'
        )
    }
    elseif ($policyOff) {
        Add-ProbeResult 'AccessVBOM registry' 'fail' ($accessLines -join '; ') (
            'Policy AccessVBOM=0 wins over the user value and greys out the Trust Center checkbox. A script cannot turn that on.'
        )
    }
    elseif ($userOn) {
        Add-ProbeResult 'AccessVBOM registry' 'ok' ($accessLines -join '; ')
    }
    elseif ($userOff) {
        Add-ProbeResult 'AccessVBOM registry' 'fail' ($accessLines -join '; ') (
            'Turn on Trust access to the VBA project object model, then fully restart Excel.'
        )
    }
    else {
        Add-ProbeResult 'AccessVBOM registry' 'warn' ($accessLines -join '; ')
    }

    if ($warningLines.Count -eq 0) {
        Add-ProbeResult 'VBAWarnings registry' 'ok' 'not set (Excel default)'
    }
    else {
        $disableAll = $false
        foreach ($line in $warningLines) {
            if ($line -match 'VBAWarnings=4') { $disableAll = $true }
        }
        $status = if ($disableAll) { 'warn' } else { 'ok' }
        $remedy = ''
        if ($disableAll) {
            $remedy = 'VBAWarnings=4 disables running macros. It is separate from AccessVBOM, which controls whether a port can edit them.'
        }
        Add-ProbeResult 'VBAWarnings registry' $status ($warningLines -join '; ') $remedy
    }

    $scratch = 'HKCU:\Software\XlvbesyncProbe'
    try {
        if (Test-Path -LiteralPath $scratch) {
            Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction Stop
        }
        New-Item -Path $scratch -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $scratch -Name 'Probe' -Value 1 -PropertyType DWord -Force -ErrorAction Stop | Out-Null
        Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction Stop
        Add-ProbeResult 'HKCU registry write' 'ok' 'created and deleted HKCU:\Software\XlvbesyncProbe (AccessVBOM was not changed)'
    }
    catch {
        Add-ProbeResult 'HKCU registry write' 'warn' (Get-ErrorDetail $_) (
            'This only matters for turning AccessVBOM on from a script. An already-enabled setting does not need a write.'
        )
    }
}

function Test-Filesystem {
    Add-ProbeSection 'Filesystem and encoding'

    $root = Join-Path $env:TEMP ('xlvbesync-probe-' + [guid]::NewGuid().ToString('n'))
    try {
        New-Item -ItemType Directory -Path $root -Force -ErrorAction Stop | Out-Null
        $file = Join-Path $root 'probe.txt'
        Set-Content -LiteralPath $file -Value 'probe' -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $file -Force -ErrorAction Stop
        $script:TempDir = $root
        Add-ProbeResult 'Temp create/delete' 'ok' $root
    }
    catch {
        Add-ProbeResult 'Temp create/delete' 'fail' (Get-ErrorDetail $_) (
            'The port has to write UTF-8 source and a hash file under a user-writable directory.'
        )
        return
    }

    try {
        $utfPath = Join-Path $script:TempDir 'utf8.txt'
        $eacute = [string][char]0x00E9
        $payload = ('{0}{1}' -f $eacute, "`r`n")
        $utf8 = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($utfPath, $payload, $utf8)
        $bytes = [System.IO.File]::ReadAllBytes($utfPath)
        $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
        $round = [System.IO.File]::ReadAllText($utfPath, $utf8)
        if ($hasBom) { throw 'UTF-8 BOM was written' }
        if ($round -ne $payload) { throw 'UTF-8 readback did not match' }
        Add-ProbeResult 'UTF-8 no BOM' 'ok' ('U+00E9 round-trip, {0} bytes, no BOM' -f $bytes.Length)
    }
    catch {
        Add-ProbeResult 'UTF-8 no BOM' 'fail' (Get-ErrorDetail $_) (
            'Source files have to be UTF-8 without a BOM. The VBE itself is ANSI; the port transcodes.'
        )
    }

    try {
        $eacute = [string][char]0x00E9
        $grin = [char]::ConvertFromUtf32(0x1F600)
        $strict = [System.Text.Encoding]::GetEncoding(
            1252,
            (New-Object System.Text.EncoderExceptionFallback),
            (New-Object System.Text.DecoderExceptionFallback)
        )
        $encoded = $strict.GetBytes($eacute)
        $decoded = $strict.GetString($encoded)
        if ($decoded -ne $eacute) { throw 'U+00E9 did not round-trip through cp1252' }
        $rejected = $false
        try { [void]$strict.GetBytes($grin) } catch { $rejected = $true }
        if (-not $rejected) { throw 'U+1F600 was accepted by cp1252' }
        Add-ProbeResult 'cp1252 transcode' 'ok' 'U+00E9 round-trips; U+1F600 is rejected'
    }
    catch {
        Add-ProbeResult 'cp1252 transcode' 'fail' (Get-ErrorDetail $_) (
            'VBA source is stored in the ANSI code page. The port has to detect characters that would become question marks.'
        )
    }

    try {
        $source = Join-Path $script:TempDir 'replace-src.txt'
        $dest = Join-Path $script:TempDir 'replace-dst.txt'
        $backup = Join-Path $script:TempDir 'replace-dst.bak'
        $utf8 = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($source, 'new', $utf8)
        [System.IO.File]::WriteAllText($dest, 'old', $utf8)
        [System.IO.File]::Replace($source, $dest, $backup)
        $destText = [System.IO.File]::ReadAllText($dest, $utf8)
        $bakText = [System.IO.File]::ReadAllText($backup, $utf8)
        if ($destText -ne 'new' -or $bakText -ne 'old') { throw 'Replace did not swap the bytes' }
        if (Test-Path -LiteralPath $source) { throw 'Replace left the source file in place' }
        Add-ProbeResult 'Atomic file replace' 'ok' '[IO.File]::Replace swapped contents and left a backup'
    }
    catch {
        Add-ProbeResult 'Atomic file replace' 'fail' (Get-ErrorDetail $_) (
            'Add-in swap mode uses an atomic replace. Live mode does not need it.'
        )
    }
}

function Test-DotNetCapabilities {
    Add-ProbeSection '.NET capabilities'

    try {
        $className = 'XlvbesyncProbe{0}' -f (Get-Random -Minimum 100000 -Maximum 999999)
        $source = 'public class {0} {{ }}' -f $className
        Add-Type -TypeDefinition $source -ErrorAction Stop -WarningAction SilentlyContinue
        $loaded = $className -as [type]
        if ($null -eq $loaded) { throw 'Add-Type returned but the type was not visible' }
        $script:AddTypeOk = $true
        Add-ProbeResult 'Add-Type' 'ok' $className
    }
    catch {
        Add-ProbeResult 'Add-Type' 'fail' (Get-ErrorDetail $_) (
            'Multi-instance Excel discovery needs a small P/Invoke. Live inject into the active instance does not.'
        )
    }

    try {
        if ([string]::IsNullOrWhiteSpace($script:TempDir)) { throw 'temp directory is not available' }
        $watcher = New-Object System.IO.FileSystemWatcher
        try {
            $watcher.Path = $script:TempDir
            $watcher.Filter = '*.bas'
            $watcher.EnableRaisingEvents = $false
        }
        finally {
            $watcher.Dispose()
        }
        $script:WatcherOk = $true
        Add-ProbeResult 'FileSystemWatcher' 'ok' 'constructed against the temp directory'
    }
    catch {
        Add-ProbeResult 'FileSystemWatcher' 'fail' (Get-ErrorDetail $_) (
            'The watcher can poll LastWriteTime once a second instead. Saves then land about a second later.'
        )
    }

    $archive = $null
    $writer = $null
    try {
        if ([string]::IsNullOrWhiteSpace($script:TempDir)) { throw 'temp directory is not available' }
        try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }
        try { Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop } catch { }
        $zipPath = Join-Path $script:TempDir 'probe.zip'
        $archive = [System.IO.Compression.ZipFile]::Open(
            $zipPath,
            [System.IO.Compression.ZipArchiveMode]::Create
        )
        $entry = $archive.CreateEntry('customUI/customUI14.xml')
        $writer = New-Object System.IO.StreamWriter ($entry.Open())
        $writer.Write('<customUI/>')
        $writer.Dispose()
        $writer = $null
        $archive.Dispose()
        $archive = $null
        $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        $found = $false
        foreach ($item in $archive.Entries) {
            if ($item.FullName -eq 'customUI/customUI14.xml') { $found = $true }
        }
        if (-not $found) { throw 'zip entry was missing on readback' }
        Add-ProbeResult 'ZipArchive' 'ok' 'wrote and read customUI/customUI14.xml'
    }
    catch {
        Add-ProbeResult 'ZipArchive' 'fail' (Get-ErrorDetail $_) (
            'Ribbon customUI sync needs zip access. Module extract and inject do not.'
        )
    }
    finally {
        if ($null -ne $writer) { $writer.Dispose() }
        if ($null -ne $archive) { $archive.Dispose() }
    }

    try {
        $found = $false
        $marshal = [type]::GetType('System.Runtime.InteropServices.Marshal', $true)
        foreach ($method in $marshal.GetMethods()) {
            if ($method.Name -eq 'GetActiveObject') { $found = $true; break }
        }
        if ($found) {
            Add-ProbeResult 'Marshal.GetActiveObject' 'ok' 'method is present (not called yet)'
        }
        else {
            Add-ProbeResult 'Marshal.GetActiveObject' 'warn' 'method is missing on this runtime' (
                'Windows PowerShell 5.1 has it. PowerShell 7 on .NET 6 does not. Microsoft.VisualBasic.Interaction.GetObject is the fallback, tested next.'
            )
        }
    }
    catch {
        Add-ProbeResult 'Marshal.GetActiveObject' 'fail' (Get-ErrorDetail $_)
    }

    $sha = $null
    try {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes('xlvbesync'))
        if ($bytes.Length -ne 32) { throw 'SHA256 returned an unexpected length' }
        Add-ProbeResult 'SHA256' 'ok' 'in-process 32-byte digest'
    }
    catch {
        Add-ProbeResult 'SHA256' 'fail' (Get-ErrorDetail $_) (
            'Unchanged-file skip uses a content hash. Get-FileHash is the cmdlet fallback.'
        )
    }
    finally {
        if ($null -ne $sha) { $sha.Dispose() }
    }
}

function Test-ExcelAbsent {
    param($ErrorRecord)
    if ((Get-SignedHResult $ErrorRecord) -eq -2147221021) { return $true }
    return ([string]$ErrorRecord.Exception.Message -match '800401E3')
}

function Connect-RunningExcel {
    # Attach only. Do not Quit a session that was already open. GetObject can
    # start Excel when nothing is running; a new PID in that case is ours and
    # is closed so the report does not claim an existing session.
    $before = Get-ExcelPids
    $app = $null
    $how = $null
    $primary = $null
    try {
        $marshal = [type]::GetType('System.Runtime.InteropServices.Marshal', $true)
        $method = $null
        foreach ($candidate in $marshal.GetMethods()) {
            if ($candidate.Name -ne 'GetActiveObject') { continue }
            $params = @($candidate.GetParameters())
            if ($params.Count -eq 1 -and $params[0].ParameterType.FullName -eq 'System.String') {
                $method = $candidate
                break
            }
        }
        if ($null -eq $method) { throw 'GetActiveObject(string) is not available' }
        $app = $method.Invoke($null, @('Excel.Application'))
        $how = 'Marshal.GetActiveObject'
    }
    catch {
        $primary = $_
    }

    if ($null -eq $app) {
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
            $app = [Microsoft.VisualBasic.Interaction]::GetObject($null, 'Excel.Application')
            $how = 'Interaction.GetObject'
        }
        catch {
            if ((Test-ExcelAbsent $primary) -or (Test-ExcelAbsent $_)) { return $null }
            if ($null -ne $primary) { throw $primary }
            throw
        }
    }

    $started = @(Get-NewExcelPids -Before $before)
    if ($how -eq 'Interaction.GetObject' -and $started.Count -gt 0) {
        try { $app.Quit() } catch { }
        Release-ComObject $app
        Start-Sleep -Milliseconds 400
        foreach ($procId in $started) {
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        }
        return $null
    }

    try {
        $version = [string](Invoke-ComRetry { $app.Version })
        $count = [int](Invoke-ComRetry { $app.Workbooks.Count })
        Add-ProbeResult 'Attach to running Excel' 'ok' (
            '{0}: Excel {1}, {2} workbook(s) open' -f $how, $version, $count
        )
    }
    finally {
        Release-ComObject $app
    }
    return $how
}

function Remove-VbaComponentByName {
    param($Project, [string]$Name)
    # Walk by index. A PowerShell foreach over a VBComponents collection breaks
    # as soon as a component is removed.
    for ($i = [int]$Project.VBComponents.Count; $i -ge 1; $i--) {
        $item = $Project.VBComponents.Item($i)
        if ([string]$item.Name -eq $Name) {
            $Project.VBComponents.Remove($item)
            return
        }
    }
}

function Test-ExcelAutomation {
    Add-ProbeSection 'Excel COM and VBProject'
    Add-ProbeLine '.. If this section hangs, press Esc in Excel or Ctrl+C here.'

    try {
        $attached = Connect-RunningExcel
        if (-not $attached) {
            Add-ProbeResult 'Attach to running Excel' 'warn' 'not running' (
                'Live inject attaches to an Excel that is already open. Launching a separate instance is the next check.'
            )
        }
    }
    catch {
        Add-ProbeResult 'Attach to running Excel' 'fail' (Get-ErrorDetail $_) (
            'The port prefers the running instance. Launching one can still work when nothing is open.'
        )
    }

    $excel = $null
    $workbook = $null
    $owned = $false
    $saved = $null

    try {
        Add-ProbeLine '.. Starting a hidden Excel instance.'
        $before = Get-ExcelPids
        try {
            $excel = New-Object -ComObject Excel.Application
        }
        catch {
            Add-ProbeResult 'Excel.Application create' 'fail' (Get-ErrorDetail $_) (
                'New-Object -ComObject Excel.Application was refused. Constrained Language, a missing ProgID, or a DCOM / attack-surface rule will do this.'
            )
            $script:OwnedPids = @(Get-NewExcelPids -Before $before)
            return
        }

        $script:OwnedPids = @(Get-NewExcelPids -Before $before)
        $owned = ($script:OwnedPids.Count -gt 0)
        if ($owned) {
            try { $excel.Visible = $false } catch { }
            try { $excel.DisplayAlerts = $false } catch { }
            try { $excel.ScreenUpdating = $false } catch { }
            try { $excel.EnableEvents = $false } catch { }
            try { $excel.AskToUpdateLinks = $false } catch { }
        }
        else {
            # CreateObject reused a running session. Hide nothing, and restore
            # every flag this script changes. Quit is not called on the way out.
            $saved = [pscustomobject]@{
                DisplayAlerts   = $excel.DisplayAlerts
                ScreenUpdating  = $excel.ScreenUpdating
                EnableEvents    = $excel.EnableEvents
            }
            try { $excel.DisplayAlerts = $false } catch { }
            try { $excel.ScreenUpdating = $false } catch { }
        }

        $version = [string](Invoke-ComRetry { $excel.Version })
        $script:PsExcel = $true
        $pidText = if ($owned) { 'new process PID {0}' -f ($script:OwnedPids -join ', ') } else { 'no new process; tests run on an already-open Excel and it will not be quit' }
        Add-ProbeResult 'Excel.Application create' 'ok' ('Excel {0}; {1}' -f $version, $pidText)

        $workbook = Invoke-ComRetry { $excel.Workbooks.Add() }
        $project = $null
        try {
            $project = Invoke-ComRetry { $workbook.VBProject }
            $count = [int](Invoke-ComRetry { $project.VBComponents.Count })
            $script:PsVbomRead = $true
            Add-ProbeResult 'VBProject read' 'ok' ('VBComponents.Count={0} on a new unsaved workbook' -f $count)
        }
        catch {
            $detail = Get-ErrorDetail $_
            $remedy = 'Enable Trust access to the VBA project object model, then fully restart Excel. Policy AccessVBOM=0 cannot be overridden.'
            if ($detail -match 'password|locked|unviewable') {
                $remedy = 'The VBA project is locked for viewing. A brand-new workbook should not be; this Excel may be applying a project password by policy.'
            }
            Add-ProbeResult 'VBProject read' 'fail' $detail $remedy
            return
        }

        $moduleName = 'XlvbesyncProbeMod'
        $marker = 'XlvbesyncProbe'
        $code = "Public Function $marker() As Long`r`n    $marker = 1`r`nEnd Function`r`n"
        try {
            $component = Invoke-ComRetry { $project.VBComponents.Add(1) }
            Invoke-ComRetry { $component.Name = $moduleName } | Out-Null
            Invoke-ComRetry { $component.CodeModule.AddFromString($code) } | Out-Null
            $lineCount = [int](Invoke-ComRetry { $component.CodeModule.CountOfLines })
            $readBack = [string](Invoke-ComRetry { $component.CodeModule.Lines(1, $lineCount) })
            if ($readBack -notlike "*$marker*") { throw 'AddFromString readback did not contain the probe function' }
            $script:PsVbomWrite = $true
            Add-ProbeResult 'VBProject write' 'ok' ('added {0}, read {1} line(s) back' -f $moduleName, $lineCount)
        }
        catch {
            Add-ProbeResult 'VBProject write' 'fail' (Get-ErrorDetail $_) (
                'Standard-module inject needs VBComponents.Add and CodeModule.AddFromString, or Import.'
            )
            try { Invoke-ComRetry { Remove-VbaComponentByName -Project $project -Name $moduleName } | Out-Null } catch { }
        }

        if (-not $script:PsVbomWrite) {
            Add-ProbeResult 'VBComponent export' 'warn' 'skipped because the write failed'
            Add-ProbeResult 'VBComponent import' 'warn' 'skipped because the write failed'
        }
        else {
            $exportPath = $null
            $exportedOk = $false
            try {
                if ([string]::IsNullOrWhiteSpace($script:TempDir)) { throw 'temp directory is not available' }
                $exportPath = Join-Path $script:TempDir 'XlvbesyncProbeMod.bas'
                if (Test-Path -LiteralPath $exportPath) { Remove-Item -LiteralPath $exportPath -Force }
                Invoke-ComRetry { $component.Export($exportPath) } | Out-Null
                if (-not (Test-Path -LiteralPath $exportPath)) { throw 'Export returned without creating the file' }
                $exported = [System.IO.File]::ReadAllText($exportPath, [System.Text.Encoding]::GetEncoding(1252))
                if ($exported -notlike "*$marker*") { throw 'Exported .bas did not contain the probe function' }
                $exportedOk = $true
                Add-ProbeResult 'VBComponent export' 'ok' (
                    '{0} characters, cp1252, contains the probe function' -f $exported.Length
                )
            }
            catch {
                Add-ProbeResult 'VBComponent export' 'fail' (Get-ErrorDetail $_) (
                    'Extract uses VBComponent.Export, which writes the ANSI code page.'
                )
            }

            if (-not $exportedOk) {
                Add-ProbeResult 'VBComponent import' 'warn' 'skipped because export failed'
                try { Invoke-ComRetry { Remove-VbaComponentByName -Project $project -Name $moduleName } | Out-Null } catch { }
            }
            else {
                try {
                    Invoke-ComRetry { $project.VBComponents.Remove($component) } | Out-Null
                    $component = $null
                    $imported = Invoke-ComRetry { $project.VBComponents.Import($exportPath) }
                    $importedCount = [int](Invoke-ComRetry { $imported.CodeModule.CountOfLines })
                    $importedText = [string](Invoke-ComRetry { $imported.CodeModule.Lines(1, $importedCount) })
                    if ($importedText -notlike "*$marker*") { throw 'Import readback did not contain the probe function' }
                    Invoke-ComRetry { Remove-VbaComponentByName -Project $project -Name $moduleName } | Out-Null
                    $script:PsImport = $true
                    Add-ProbeResult 'VBComponent import' 'ok' 'removed, imported, read back, removed again'
                }
                catch {
                    Add-ProbeResult 'VBComponent import' 'fail' (Get-ErrorDetail $_) (
                        'Standard modules, classes, and userforms are replaced by remove-then-import.'
                    )
                    try { Invoke-ComRetry { Remove-VbaComponentByName -Project $project -Name $moduleName } | Out-Null } catch { }
                }
            }
        }

        try {
            $document = $null
            $total = [int]$project.VBComponents.Count
            for ($i = 1; $i -le $total; $i++) {
                $item = $project.VBComponents.Item($i)
                if ([int]$item.Type -eq 100 -and [string]$item.Name -eq 'ThisWorkbook') {
                    $document = $item
                    break
                }
            }
            if ($null -eq $document) { throw 'ThisWorkbook (type 100) was not in the new workbook' }
            $docModule = $document.CodeModule
            $original = [int](Invoke-ComRetry { $docModule.CountOfLines })
            Invoke-ComRetry { $docModule.AddFromString("' xlvbesync probe") } | Out-Null
            $after = [int](Invoke-ComRetry { $docModule.CountOfLines })
            if ($after -le $original) { throw 'AddFromString did not add a line to ThisWorkbook' }
            $added = [string](Invoke-ComRetry { $docModule.Lines($original + 1, $after - $original) })
            if ($added -notlike '*xlvbesync probe*') { throw 'document-module readback did not contain the marker' }
            Invoke-ComRetry { $docModule.DeleteLines($original + 1, $after - $original) } | Out-Null
            $restored = [int](Invoke-ComRetry { $docModule.CountOfLines })
            if ($restored -ne $original) {
                throw ('ThisWorkbook line count changed from {0} to {1}' -f $original, $restored)
            }
            $script:PsDocument = $true
            Add-ProbeResult 'Document module edit' 'ok' 'ThisWorkbook AddFromString + DeleteLines restored the original line count'
        }
        catch {
            Add-ProbeResult 'Document module edit' 'fail' (Get-ErrorDetail $_) (
                'Sheet and ThisWorkbook modules cannot be imported. They are edited in place.'
            )
        }

        if ($script:PsVbomWrite) {
            try {
                $samples = @()
                for ($n = 1; $n -le 5; $n++) {
                    $watch = [System.Diagnostics.Stopwatch]::StartNew()
                    $tempName = 'XlvProbe{0}' -f $n
                    $tempCode = "Public Function $tempName() As Long`r`n    $tempName = $n`r`nEnd Function`r`n"
                    $tempComponent = $project.VBComponents.Add(1)
                    $tempComponent.Name = $tempName
                    $tempComponent.CodeModule.AddFromString($tempCode)
                    $project.VBComponents.Remove($tempComponent)
                    $watch.Stop()
                    $samples += $watch.Elapsed.TotalMilliseconds
                }
                $sorted = @($samples | Sort-Object)
                $script:TimingMedian = [math]::Round($sorted[2], 1)
                $shown = (@($samples | ForEach-Object { '{0:0.0}' -f $_ })) -join ', '
                Add-ProbeResult 'Inject cycle timing' 'ok' (
                    'median {0:0.0} ms over 5 add/write/remove cycles ({1})' -f $script:TimingMedian, $shown
                )
            }
            catch {
                Add-ProbeResult 'Inject cycle timing' 'warn' (Get-ErrorDetail $_)
            }
        }
    }
    catch {
        Add-ProbeResult 'Excel automation' 'fail' (Get-ErrorDetail $_) (
            'The hidden Excel instance was started, but a later COM call failed before the probe finished.'
        )
    }
    finally {
        # Drop every RCW except the Application and Workbook before Quit.
        # PowerShell keeps loop variables alive, and Excel will not exit while
        # those references exist.
        $project = $null
        $component = $null
        $imported = $null
        $document = $null
        $docModule = $null
        $item = $null
        $tempComponent = $null
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
        }
        if ($owned -and $null -ne $excel) {
            try { $excel.Quit() } catch { }
        }
        elseif ($null -ne $excel -and $null -ne $saved) {
            try { $excel.DisplayAlerts = $saved.DisplayAlerts } catch { }
            try { $excel.ScreenUpdating = $saved.ScreenUpdating } catch { }
            try { $excel.EnableEvents = $saved.EnableEvents } catch { }
        }
        Release-ComObject $workbook
        Release-ComObject $excel
        $workbook = $null
        $excel = $null
        try {
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
        catch { }
        # CreateObject sometimes returns the session that is already open and
        # does not start a process. That list must not be treated as ours.
        if ($script:PsExcel -and -not $owned) { $script:OwnedPids = @() }
        Stop-OwnedExcel
    }
}

function Test-CscriptFallback {
    Add-ProbeSection 'cscript fallback'

    $cscript = Join-Path $env:SystemRoot 'System32\cscript.exe'
    if (-not (Test-Path -LiteralPath $cscript)) {
        Add-ProbeResult 'cscript.exe' 'fail' 'System32\cscript.exe is missing'
        return
    }

    $psCore = $script:PsExcel -and $script:PsVbomRead -and $script:PsVbomWrite
    if ($psCore) {
        Add-ProbeResult 'cscript.exe' 'ok' 'present; VBProject probe skipped because PowerShell COM write succeeded'
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:TempDir)) {
        Add-ProbeResult 'cscript VBProject' 'fail' 'temp directory is not available, so the VBScript probe was not written'
        return
    }

    $vbsPath = Join-Path $script:TempDir 'probe.vbs'
    $stdout = Join-Path $script:TempDir 'cscript.out'
    $stderr = Join-Path $script:TempDir 'cscript.err'
    $body = @'
On Error Resume Next
Dim xl, wb, proj, comp, got, ver
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "COM_FAIL " & Err.Number & " " & Err.Description
    WScript.Quit 2
End If
xl.Visible = False
xl.DisplayAlerts = False
Err.Clear
Set wb = xl.Workbooks.Add()
If Err.Number <> 0 Then
    WScript.Echo "COM_FAIL Workbooks.Add " & Err.Number & " " & Err.Description
    xl.Quit
    WScript.Quit 2
End If
Err.Clear
Set proj = wb.VBProject
If Err.Number <> 0 Then
    WScript.Echo "VBOM_FAIL " & Err.Number & " " & Err.Description
    wb.Close False
    xl.Quit
    WScript.Quit 3
End If
Err.Clear
Set comp = proj.VBComponents.Add(1)
If Err.Number <> 0 Then
    WScript.Echo "WRITE_FAIL Add " & Err.Number & " " & Err.Description
    wb.Close False
    xl.Quit
    WScript.Quit 3
End If
comp.Name = "XlvbesyncProbeMod"
comp.CodeModule.AddFromString "Public Function XlvbesyncProbe() As Long" & vbCrLf & "    XlvbesyncProbe = 1" & vbCrLf & "End Function" & vbCrLf
If Err.Number <> 0 Then
    WScript.Echo "WRITE_FAIL AddFromString " & Err.Number & " " & Err.Description
    wb.Close False
    xl.Quit
    WScript.Quit 3
End If
got = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
ver = xl.Version
proj.VBComponents.Remove comp
wb.Close False
xl.Quit
If InStr(1, got, "XlvbesyncProbe", 1) = 0 Then
    WScript.Echo "WRITE_FAIL readback"
    WScript.Quit 3
End If
WScript.Echo "VBOM_OK Excel " & ver
WScript.Quit 0
'@
    Set-Content -LiteralPath $vbsPath -Value $body -Encoding ASCII

    Add-ProbeLine '.. Running cscript VBProject probe (killed if it runs past 20s).'
    $before = Get-ExcelPids
    $proc = $null
    try {
        $proc = Start-Process -FilePath $cscript -ArgumentList @(
            '//NoLogo', '//B', $vbsPath
        ) -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        if (-not $proc.WaitForExit(20000)) {
            try { $proc.Kill() } catch { }
            Add-ProbeResult 'cscript VBProject' 'warn' 'still running after 20s; killed' (
                'A cold Excel start can exceed 20 seconds. Open Excel once so it is warm, then run this probe again before treating the fallback as blocked.'
            )
        }
        else {
            $outText = ''
            if (Test-Path -LiteralPath $stdout) {
                $outText = [string](Get-Content -LiteralPath $stdout -Raw -ErrorAction SilentlyContinue)
            }
            $errText = ''
            if (Test-Path -LiteralPath $stderr) {
                $errText = [string](Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue)
            }
            $outText = (($outText -replace '\s+', ' ').Trim())
            $errText = (($errText -replace '\s+', ' ').Trim())
            if ($outText -match 'VBOM_OK') {
                $script:CscriptWrite = $true
                Add-ProbeResult 'cscript VBProject' 'ok' $outText
            }
            else {
                $detail = $outText
                if ($errText) { $detail = ('{0} {1}' -f $detail, $errText).Trim() }
                if ([string]::IsNullOrWhiteSpace($detail)) {
                    $detail = 'exit {0}, no output' -f $proc.ExitCode
                }
                Add-ProbeResult 'cscript VBProject' 'fail' $detail (
                    'If both PowerShell and cscript are refused, nothing outside Excel can edit the VBA project.'
                )
            }
        }
    }
    catch {
        Add-ProbeResult 'cscript VBProject' 'fail' (Get-ErrorDetail $_) (
            'cscript.exe is present but could not be started. AppLocker often blocks it.'
        )
    }
    finally {
        $leaked = @(Get-NewExcelPids -Before $before)
        foreach ($procId in $leaked) {
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        }
        if ($leaked.Count -gt 0) {
            Add-ProbeResult 'cscript Excel cleanup' 'warn' (
                'Stopped Excel PID {0}, started by the cscript probe.' -f ($leaked -join ', ')
            )
        }
    }
}

function Get-ExitCode {
    $core = $script:PsExcel -and $script:PsVbomRead -and $script:PsVbomWrite -and
        $script:PsImport -and $script:PsDocument
    if ($core -and $script:AddTypeOk -and $script:WatcherOk) { return 0 }
    if ($core) { return 1 }
    if ($script:PsExcel -and $script:PsVbomRead -and $script:PsVbomWrite) { return 1 }
    if ($script:CscriptWrite) { return 3 }
    return 2
}

function Write-Verdict {
    param([int]$Code)
    Add-ProbeSection 'Verdict'

    $timing = ''
    if ($null -ne $script:TimingMedian) {
        $timing = (' Measured median add/write/remove cycle: {0:0.0} ms.' -f $script:TimingMedian)
    }

    switch ($Code) {
        0 {
            Add-ProbeLine ('STATUS ready exit=0')
            Add-ProbeLine ('A PowerShell port can do live extract and inject on this PC. Add-Type and FileSystemWatcher are available, so multi-instance discovery and a save watcher are available too.{0}' -f $timing)
        }
        1 {
            $gaps = @()
            if (-not $script:AddTypeOk) { $gaps += 'Add-Type (no Running Object Table scan; the active Excel still works)' }
            if (-not $script:WatcherOk) { $gaps += 'FileSystemWatcher (poll about once a second instead)' }
            if (-not $script:PsImport) { $gaps += 'VBComponent.Import (standard-module replace is remove-then-import)' }
            if (-not $script:PsDocument) { $gaps += 'document-module in-place edit' }
            Add-ProbeLine 'STATUS degraded exit=1'
            Add-ProbeLine ('Live VBProject write works.{0} Missing: {1}.' -f $timing, ($gaps -join '; '))
        }
        3 {
            Add-ProbeLine 'STATUS cscript exit=3'
            Add-ProbeLine 'PowerShell cannot drive Excel COM here. cscript.exe can create Excel and write a standard module, so a port would be VBScript rather than PowerShell.'
        }
        default {
            Add-ProbeLine 'STATUS blocked exit=2'
            Add-ProbeLine 'VBProject automation is blocked from both PowerShell and cscript. Office.js and Office Scripts cannot edit VBA either. The usual causes are Constrained Language plus a blocked cscript, or Trust Center / policy AccessVBOM left off.'
        }
    }

    Add-ProbeLine ''
    Add-ProbeLine 'Paste this whole report back. It does not include workbook contents.'
}

function Write-ProbeReport {
    param([string]$Path)
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $text = ($script:Lines -join "`r`n") + "`r`n"
    try {
        [System.IO.File]::WriteAllText($Path, $text, $utf8)
    }
    catch {
        Set-Content -LiteralPath $Path -Value $script:Lines -Encoding UTF8
    }
}

Add-ProbeLine 'xlvbesync work-PC capability probe'
Add-ProbeLine 'Starts a hidden Excel of its own. Does not save, and does not quit an Excel that was already open.'
Add-ProbeLine 'If this hangs, Excel has a modal dialog open. Press Esc there, or Ctrl+C here.'

Test-HostEnvironment
Test-PolicyPresence
Test-OfficeRegistration
Test-Filesystem
Test-DotNetCapabilities
Test-ExcelAutomation
Test-CscriptFallback

$code = Get-ExitCode
Write-Verdict -Code $code

$report = Get-ReportPath
Add-ProbeLine ''
Add-ProbeLine ('Report written to {0}' -f $report)
try {
    Write-ProbeReport -Path $report
}
catch {
    Write-Host ('Could not write the report: {0}' -f (Get-ErrorDetail $_))
}

if ($script:TempDir -and (Test-Path -LiteralPath $script:TempDir)) {
    Remove-Item -LiteralPath $script:TempDir -Recurse -Force -ErrorAction SilentlyContinue
}

if ($PSCommandPath) {
    exit $code
}

Write-Host ('Exit code: {0} (host left open because this was not run with -File)' -f $code)
$global:LASTEXITCODE = $code
