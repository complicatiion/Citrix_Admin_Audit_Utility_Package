param(
    [string]$Action = 'QuickAudit',
    [string]$AdminAddress = '',
    [string]$ReportPath = ''
)

$ErrorActionPreference = 'SilentlyContinue'
$script:Lines = New-Object System.Collections.Generic.List[string]

function Add-Line {
    param([string]$Text = '')
    $script:Lines.Add($Text)
}

function Add-Section {
    param([string]$Title)
    Add-Line ('=' * 78)
    Add-Line $Title
    Add-Line ('=' * 78)
}

function Add-TextBlock {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) {
        Add-Line
        return
    }
    $normalized = $Text -replace "`r`n", "`n"
    foreach ($line in $normalized.Split("`n")) {
        Add-Line $line.TrimEnd()
    }
}

function Add-Table {
    param(
        $InputObject,
        [string[]]$Property = @()
    )
    if ($null -eq $InputObject) {
        Add-Line 'No data returned.'
        return
    }

    $items = @($InputObject)
    if ($items.Count -eq 0) {
        Add-Line 'No data returned.'
        return
    }

    try {
        if ($Property.Count -gt 0) {
            $text = $items | Select-Object $Property | Format-Table -AutoSize | Out-String -Width 500
        }
        else {
            $text = $items | Format-Table -AutoSize | Out-String -Width 500
        }
        Add-TextBlock $text.TrimEnd()
    }
    catch {
        Add-Line ('Formatting error: ' + $_.Exception.Message)
    }
}

function Add-List {
    param(
        $InputObject,
        [string[]]$Property = @()
    )
    if ($null -eq $InputObject) {
        Add-Line 'No data returned.'
        return
    }

    $items = @($InputObject)
    if ($items.Count -eq 0) {
        Add-Line 'No data returned.'
        return
    }

    try {
        if ($Property.Count -gt 0) {
            $text = $items | Select-Object $Property | Format-List | Out-String -Width 500
        }
        else {
            $text = $items | Format-List | Out-String -Width 500
        }
        Add-TextBlock $text.TrimEnd()
    }
    catch {
        Add-Line ('Formatting error: ' + $_.Exception.Message)
    }
}

function Add-Note {
    param([string]$Text)
    Add-Line ('- ' + $Text)
}

function Invoke-Block {
    param(
        [string]$Title,
        [scriptblock]$ScriptBlock
    )

    Add-Section $Title
    try {
        & $ScriptBlock
    }
    catch {
        Add-Line ('ERROR: ' + $_.Exception.Message)
    }
    Add-Line
}

function Get-AdminSplat {
    if ([string]::IsNullOrWhiteSpace($AdminAddress)) {
        return @{}
    }
    return @{ AdminAddress = $AdminAddress }
}

function Test-CitrixCommand {
    param([string]$Name)
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Import-CitrixSdk {
    $loaded = New-Object System.Collections.Generic.List[string]

    try {
        $registeredSnapins = Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'Citrix*' }
        foreach ($snap in $registeredSnapins) {
            try {
                if (-not (Get-PSSnapin -Name $snap.Name -ErrorAction SilentlyContinue)) {
                    Add-PSSnapin -Name $snap.Name -ErrorAction Stop
                }
                if (-not $loaded.Contains($snap.Name)) { [void]$loaded.Add($snap.Name) }
            }
            catch { }
        }
    }
    catch { }

    try {
        $mods = Get-Module -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'Citrix*' }
        foreach ($mod in $mods) {
            try { Import-Module $mod.Name -ErrorAction SilentlyContinue | Out-Null } catch { }
        }
    }
    catch { }

    return @($loaded)
}

$script:LoadedSdk = Import-CitrixSdk
$script:AA = Get-AdminSplat

function Add-Header {
    Add-Section 'Citrix Admin Audit Utility'
    Add-Line ('Timestamp        : ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    Add-Line ('Local computer   : ' + $env:COMPUTERNAME)
    if ([string]::IsNullOrWhiteSpace($AdminAddress)) {
        Add-Line 'AdminAddress     : Localhost default'
    }
    else {
        Add-Line ('AdminAddress     : ' + $AdminAddress)
    }
    Add-Line ('User context     : ' + [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    Add-Line
}

function Get-LocalRoleInfo {
    $svc = Get-Service -ErrorAction SilentlyContinue
    $isController = [bool]($svc | Where-Object { $_.DisplayName -match 'Citrix Broker Service|Citrix Configuration Service|Citrix Machine Creation Service' })
    $isVDA = [bool]($svc | Where-Object { $_.DisplayName -match 'Citrix Desktop Service|Citrix ICA Service' -or $_.Name -match 'BrokerAgent|PortICA' })

    $studioInstalled = $false
    try {
        $apps = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' `
                                -ErrorAction SilentlyContinue
        $studioInstalled = [bool]($apps | Where-Object { $_.DisplayName -match 'Citrix Studio|Citrix Virtual Apps and Desktops' })
    }
    catch { }

    return [pscustomobject]@{
        ComputerName        = $env:COMPUTERNAME
        DeliveryController  = $isController
        VDA                 = $isVDA
        StudioOrSdkLikely   = $studioInstalled -or ($script:LoadedSdk.Count -gt 0)
        LoadedCitrixSnapins = if ($script:LoadedSdk.Count -gt 0) { $script:LoadedSdk -join '; ' } else { '' }
    }
}

function Section-Environment {
    Invoke-Block 'Environment and role detection' {
        Add-List (Get-LocalRoleInfo)
        Add-Line 'PowerShell version:'
        Add-Table ([pscustomobject]@{
            PSVersion      = $PSVersionTable.PSVersion.ToString()
            Edition        = $PSVersionTable.PSEdition
            CLRVersion     = if ($PSVersionTable.CLRVersion) { $PSVersionTable.CLRVersion.ToString() } else { '' }
            LoadedSdkCount = $script:LoadedSdk.Count
        })
    }
}

function Section-SiteOverview {
    Invoke-Block 'Site overview' {
        if (Test-CitrixCommand 'Get-BrokerSite') {
            $site = Get-BrokerSite @script:AA
            Add-List $site -Property @(
                'Name','Description','UUID','DefaultMinimumFunctionalLevel',
                'LicensedSessionsActive','LicensingGracePeriodActive',
                'LicensingOutOfBoxGracePeriodActive','LicensingGraceHoursLeft',
                'LicenseGraceSessionsRemaining'
            )
        }
        else {
            Add-Line 'Get-BrokerSite is not available on this machine.'
        }

        if (Test-CitrixCommand 'Get-ConfigSite') {
            Add-Line 'Configuration site:'
            Add-List (Get-ConfigSite @script:AA)
        }

        if (Test-CitrixCommand 'Get-LogSite') {
            Add-Line 'Configuration logging site:'
            Add-List (Get-LogSite @script:AA)
        }
    }
}

function Section-Controllers {
    Invoke-Block 'Controllers and SDK service status' {
        if (Test-CitrixCommand 'Get-BrokerController') {
            $controllers = Get-BrokerController @script:AA
            Add-Table $controllers -Property @(
                'DNSName','State','Version','DesktopsRegistered','LastActivityTime','ActiveSiteServices','ZoneName'
            )
        }
        else {
            Add-Line 'Get-BrokerController is not available on this machine.'
        }

        $statusCmds = @(
            'Get-BrokerServiceStatus',
            'Get-ConfigServiceStatus',
            'Get-AcctServiceStatus',
            'Get-ProvServiceStatus',
            'Get-OrchServiceStatus',
            'Get-LogServiceStatus',
            'Get-MonitorServiceStatus',
            'Get-AdminServiceStatus'
        )

        foreach ($cmd in $statusCmds) {
            Add-Line
            Add-Line ('[' + $cmd + ']')
            if (Test-CitrixCommand $cmd) {
                try {
                    $result = & $cmd @script:AA
                    Add-List $result
                }
                catch {
                    Add-Line ('ERROR: ' + $_.Exception.Message)
                }
            }
            else {
                Add-Line 'Command not available.'
            }
        }
    }
}

function Section-CatalogsAndGroups {
    Invoke-Block 'Machine catalogs' {
        if (Test-CitrixCommand 'Get-BrokerCatalog') {
            $catalogs = Get-BrokerCatalog @script:AA
            Add-Table $catalogs -Property @(
                'Name','ProvisioningType','AllocationType','SessionSupport',
                'PersistUserChanges','MachinesArePhysical','IsRemotePC',
                'MinimumFunctionalLevel','AvailableCount','UsedCount','UnassignedCount','ZoneName'
            )
        }
        else {
            Add-Line 'Get-BrokerCatalog is not available on this machine.'
        }
    }

    Invoke-Block 'Delivery groups' {
        if (Test-CitrixCommand 'Get-BrokerDesktopGroup') {
            $groups = Get-BrokerDesktopGroup @script:AA
            Add-Table $groups -Property @(
                'Name','PublishedName','DeliveryType','SessionSupport','Enabled',
                'InMaintenanceMode','IsRemotePC','TotalDesktops',
                'DesktopsAvailable','DesktopsInUse','DesktopsDisconnected',
                'DesktopsUnregistered','ZoneName'
            )
        }
        else {
            Add-Line 'Get-BrokerDesktopGroup is not available on this machine.'
        }
    }
}

function Section-Machines {
    Invoke-Block 'Machine health summary' {
        if (Test-CitrixCommand 'Group-BrokerMachine') {
            Add-Line 'Grouped by registration state:'
            Add-Table (Group-BrokerMachine @script:AA -Property RegistrationState | Select-Object Count,Name)

            Add-Line
            Add-Line 'Grouped by power state:'
            Add-Table (Group-BrokerMachine @script:AA -Property PowerState | Select-Object Count,Name)

            Add-Line
            Add-Line 'Grouped by summary state:'
            Add-Table (Group-BrokerMachine @script:AA -Property SummaryState | Select-Object Count,Name)
        }
        elseif (Test-CitrixCommand 'Get-BrokerMachine') {
            $machines = Get-BrokerMachine @script:AA -MaxRecordCount 2000
            Add-Line 'Grouped by registration state (local grouping, up to first 2000 machines):'
            Add-Table ($machines | Group-Object RegistrationState | Select-Object Count,Name)
        }
        else {
            Add-Line 'Broker machine cmdlets are not available on this machine.'
        }
    }

    Invoke-Block 'Unregistered machines (top 100)' {
        if (Test-CitrixCommand 'Get-BrokerMachine') {
            $items = Get-BrokerMachine @script:AA -RegistrationState Unregistered -MaxRecordCount 100
            Add-Table $items -Property @(
                'MachineName','CatalogName','DesktopGroupName','PowerState',
                'FaultState','InMaintenanceMode','LastDeregistrationReason',
                'LastDeregistrationTime','AgentVersion','ControllerDNSName'
            )
        }
        else {
            Add-Line 'Get-BrokerMachine is not available on this machine.'
        }
    }

    Invoke-Block 'Machines in maintenance mode (top 100)' {
        if (Test-CitrixCommand 'Get-BrokerMachine') {
            $items = Get-BrokerMachine @script:AA -InMaintenanceMode $true -MaxRecordCount 100
            Add-Table $items -Property @(
                'MachineName','CatalogName','DesktopGroupName','PowerState',
                'RegistrationState','FaultState','ImageOutOfDate','AgentVersion'
            )
        }
        else {
            Add-Line 'Get-BrokerMachine is not available on this machine.'
        }
    }
}

function Section-Sessions {
    Invoke-Block 'Session summary' {
        if (Test-CitrixCommand 'Group-BrokerSession') {
            Add-Line 'Grouped by session state:'
            Add-Table (Group-BrokerSession @script:AA -Property SessionState | Select-Object Count,Name)

            Add-Line
            Add-Line 'Grouped by protocol:'
            Add-Table (Group-BrokerSession @script:AA -Property Protocol | Select-Object Count,Name)
        }
        else {
            Add-Line 'Group-BrokerSession is not available on this machine.'
        }
    }

    Invoke-Block 'Current sessions (top 100)' {
        if (Test-CitrixCommand 'Get-BrokerSession') {
            $sessions = Get-BrokerSession @script:AA -MaxRecordCount 100
            Add-Table $sessions -Property @(
                'UserName','MachineName','DesktopGroupName','SessionState',
                'SessionType','Protocol','StartTime','ClientName','ClientAddress'
            )
        }
        else {
            Add-Line 'Get-BrokerSession is not available on this machine.'
        }
    }

    Invoke-Block 'Disconnected sessions (top 100)' {
        if (Test-CitrixCommand 'Get-BrokerSession') {
            $sessions = Get-BrokerSession @script:AA -SessionState Disconnected -MaxRecordCount 100
            Add-Table $sessions -Property @(
                'UserName','MachineName','DesktopGroupName','SessionState',
                'SessionStateChangeTime','Protocol','ClientName','ClientAddress'
            )
        }
        else {
            Add-Line 'Get-BrokerSession is not available on this machine.'
        }
    }
}

function Section-Policies {
    Invoke-Block 'Access policy rules' {
        if (Test-CitrixCommand 'Get-BrokerAccessPolicyRule') {
            $rules = Get-BrokerAccessPolicyRule @script:AA
            Add-Table $rules -Property @(
                'Name','Enabled','DesktopGroupName','AllowedConnections','AllowedProtocols',
                'IncludedSmartAccessFilterEnabled','IncludedUserFilterEnabled','ExcludedUserFilterEnabled'
            )
        }
        else {
            Add-Line 'Get-BrokerAccessPolicyRule is not available on this machine.'
        }
    }

    Invoke-Block 'Desktop entitlement rules' {
        if (Test-CitrixCommand 'Get-BrokerEntitlementPolicyRule') {
            $rules = Get-BrokerEntitlementPolicyRule @script:AA
            Add-Table $rules -Property @(
                'Name','Enabled','DesktopGroupName','PublishedName','BrowserName',
                'IncludedUserFilterEnabled','ExcludedUserFilterEnabled'
            )
        }
        else {
            Add-Line 'Get-BrokerEntitlementPolicyRule is not available on this machine.'
        }
    }

    Invoke-Block 'Desktop assignment rules' {
        if (Test-CitrixCommand 'Get-BrokerAssignmentPolicyRule') {
            $rules = Get-BrokerAssignmentPolicyRule @script:AA
            Add-Table $rules -Property @(
                'Name','Enabled','DesktopGroupName','PublishedName','MaxDesktops',
                'IncludedUserFilterEnabled','ExcludedUserFilterEnabled'
            )
        }
        else {
            Add-Line 'Get-BrokerAssignmentPolicyRule is not available on this machine.'
        }
    }

    Invoke-Block 'Citrix Group Policy SDK availability' {
        $gpSnapinLoaded = [bool](Get-PSSnapin -Name 'citrix.common.grouppolicy' -ErrorAction SilentlyContinue)
        $gpRegistered = [bool](Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'citrix.common.grouppolicy' })
        Add-Table ([pscustomobject]@{
            GroupPolicySnapinLoaded = $gpSnapinLoaded
            GroupPolicySdkDetected  = $gpRegistered
        })
        Add-Note 'This utility reports rule-oriented policy data and SDK availability only.'
        Add-Note 'It intentionally does not export the opaque site-wide desktop policy blob.'
    }
}

function Section-MCS {
    Invoke-Block 'Provisioning schemes (MCS / provisioning)' {
        if (Test-CitrixCommand 'Get-ProvScheme') {
            $schemes = Get-ProvScheme @script:AA
            Add-Table $schemes -Property @(
                'ProvisioningSchemeName','HostingUnitName','IdentityPoolName','CleanOnBoot',
                'CpuCount','MemoryMB','MasterImageVM','MasterImageVMDate',
                'UseWriteBackCache','WriteBackCacheDiskSize','WriteBackCacheMemorySize',
                'UseFullDiskCloneProvisioning','CurrentMasterImageUid','TaskId'
            )
        }
        else {
            Add-Line 'Get-ProvScheme is not available on this machine.'
        }
    }

    Invoke-Block 'Provisioning tasks (top 25)' {
        if (Test-CitrixCommand 'Get-ProvTask') {
            $tasks = Get-ProvTask @script:AA -MaxRecordCount 25
            Add-Table $tasks
            Add-Note 'Tasks associated with machine catalog creation may be absent for catalogs created through Web Studio.'
        }
        else {
            Add-Line 'Get-ProvTask is not available on this machine.'
        }
    }

    Invoke-Block 'Hypervisor connections' {
        if (Test-CitrixCommand 'Get-BrokerHypervisorConnection') {
            $conns = Get-BrokerHypervisorConnection @script:AA
            Add-Table $conns -Property @(
                'Name','State','PluginId','ZoneName','Scopes','Capabilities'
            )
        }
        else {
            Add-Line 'Get-BrokerHypervisorConnection is not available on this machine.'
        }
    }

    Invoke-Block 'Hosting units' {
        if (Test-CitrixCommand 'Get-HypHostingUnit') {
            $units = Get-HypHostingUnit @script:AA
            Add-Table $units -Property @(
                'Name','HypervisorConnectionName','RootPath','NetworkPath','StoragePath','StoragePaths','PluginId'
            )
        }
        else {
            Add-Line 'Get-HypHostingUnit is not available on this machine.'
        }
    }
}

function Section-LocalChecks {
    Invoke-Block 'Local Citrix software inventory' {
        try {
            $apps = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' `
                                    -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.DisplayName -match 'Citrix|Unidesk|App Layering|Virtual Delivery Agent|Studio|Workspace|Director'
                    } |
                    Select-Object DisplayName,DisplayVersion,Publisher,InstallDate |
                    Sort-Object DisplayName
            Add-Table $apps
        }
        catch {
            Add-Line ('Inventory error: ' + $_.Exception.Message)
        }
    }

    Invoke-Block 'Local Citrix services' {
        try {
            $services = Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.DisplayName -match 'Citrix|Unidesk|App Layering' -or
                    $_.Name -match 'Broker|Citrix|Ctx|PortICA|Unidesk'
                } |
                Select-Object Name,DisplayName,State,StartMode,StartName |
                Sort-Object DisplayName
            Add-Table $services
        }
        catch {
            Add-Line ('Service query error: ' + $_.Exception.Message)
        }
    }

    Invoke-Block 'Local VDA registry and controller list' {
        $paths = @(
            'HKLM:\SOFTWARE\Citrix\VirtualDesktopAgent',
            'HKLM:\SOFTWARE\WOW6432Node\Citrix\VirtualDesktopAgent'
        )

        $found = $false
        foreach ($path in $paths) {
            if (Test-Path $path) {
                $found = $true
                Add-Line ('Registry path: ' + $path)
                try {
                    $item = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
                    Add-List ([pscustomobject]@{
                        ListOfDDCs              = $item.ListOfDDCs
                        ListOfSIDs              = $item.ListOfSIDs
                        ControllerRegistrarPort = $item.ControllerRegistrarPort
                        EnableRemoteManagement  = $item.EnableRemoteManagement
                        InstallDir              = $item.InstallDir
                        StartMenuShortcuts      = $item.StartMenuShortcuts
                    })
                }
                catch {
                    Add-Line ('Registry read error: ' + $_.Exception.Message)
                }
            }
        }

        if (-not $found) {
            Add-Line 'No local VirtualDesktopAgent registry path found.'
        }
    }

    Invoke-Block 'Local App Layering indicators' {
        $indicators = [ordered]@{
            UnideskProgramFiles    = [bool](Test-Path 'C:\Program Files\Citrix\Unidesk')
            AppLayeringProgramFiles= [bool](Test-Path 'C:\Program Files\Citrix\App Layering')
            UnideskServicePresent  = [bool](Get-Service -Name '*unidesk*' -ErrorAction SilentlyContinue)
            LayeringUninstallEntry = $false
        }

        try {
            $apps = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' `
                                    -ErrorAction SilentlyContinue
            $indicators.LayeringUninstallEntry = [bool]($apps | Where-Object { $_.DisplayName -match 'App Layering|Unidesk' })
        }
        catch { }

        Add-List ([pscustomobject]$indicators)
        Add-Note 'This is a presence check only. It does not validate ELM connectivity or OS Machine Tools health.'
    }
}

function Get-RelevantEvents {
    param(
        [string]$LogName,
        [int]$DaysBack = 7,
        [int]$First = 60
    )

    $events = Get-WinEvent -FilterHashtable @{ LogName = $LogName; StartTime = (Get-Date).AddDays(-$DaysBack) } -ErrorAction SilentlyContinue
    $rows = foreach ($ev in $events) {
        $msg = ''
        try { $msg = [string]$ev.Message } catch { $msg = '' }

        if ($ev.ProviderName -match 'Citrix|Broker|ICA|TerminalServices|App Layering|Unidesk' -or
            $msg -match 'Citrix|Broker Service|Desktop Service|BrokerAgent|PortICA|Machine Creation|Virtual Delivery Agent|HDX|ICA|App Layering|Unidesk') {
            [pscustomobject]@{
                TimeCreated = $ev.TimeCreated
                Id          = $ev.Id
                Provider    = $ev.ProviderName
                Level       = $ev.LevelDisplayName
                Message     = (($msg -replace '\r?\n', ' ') -replace '\s{2,}', ' ').Trim()
            }
        }
    }

    return @($rows | Select-Object -First $First)
}

function Section-LocalEvents {
    Invoke-Block 'Local Citrix-related services (runtime view)' {
        try {
            $services = Get-Service -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.DisplayName -match 'Citrix|Unidesk|App Layering' -or
                    $_.Name -match 'Broker|Citrix|Ctx|PortICA|Unidesk'
                } |
                Select-Object Status,Name,DisplayName |
                Sort-Object DisplayName
            Add-Table $services
        }
        catch {
            Add-Line ('Service query error: ' + $_.Exception.Message)
        }
    }

    Invoke-Block 'Application log (Citrix-related, last 7 days)' {
        Add-Table (Get-RelevantEvents -LogName 'Application' -DaysBack 7 -First 80)
    }

    Invoke-Block 'System log (Citrix-related, last 7 days)' {
        Add-Table (Get-RelevantEvents -LogName 'System' -DaysBack 7 -First 80)
    }
}

function Section-Notes {
    Invoke-Block 'Design notes' {
        Add-Note 'This utility is read-only. It does not modify catalogs, delivery groups, policies, or sessions.'
        Add-Note 'It intentionally uses Get-BrokerMachine instead of the deprecated Get-BrokerDesktop cmdlet.'
        Add-Note 'For separate Studio servers, set AdminAddress to a Delivery Controller FQDN if local SDK calls do not resolve the site automatically.'
        Add-Note 'MCS and App Layering are reported separately. Provisioning schemes are queried through MCS cmdlets, while local App Layering checks are presence-based only.'
    }
}

function Build-QuickAudit {
    Add-Header
    Section-Environment
    Section-SiteOverview
    Section-Controllers
    Section-CatalogsAndGroups
    Section-Machines
    Section-Sessions
    Section-MCS
    Section-Notes
}

function Build-SiteOverview {
    Add-Header
    Section-Environment
    Section-SiteOverview
    Section-Notes
}

function Build-Controllers {
    Add-Header
    Section-Environment
    Section-Controllers
    Section-Notes
}

function Build-Catalogs {
    Add-Header
    Section-Environment
    Section-CatalogsAndGroups
    Section-Notes
}

function Build-Machines {
    Add-Header
    Section-Environment
    Section-Machines
    Section-Notes
}

function Build-Sessions {
    Add-Header
    Section-Environment
    Section-Sessions
    Section-Notes
}

function Build-Policies {
    Add-Header
    Section-Environment
    Section-Policies
    Section-Notes
}

function Build-MCS {
    Add-Header
    Section-Environment
    Section-MCS
    Section-Notes
}

function Build-LocalChecks {
    Add-Header
    Section-Environment
    Section-LocalChecks
    Section-Notes
}

function Build-LocalEvents {
    Add-Header
    Section-Environment
    Section-LocalEvents
    Section-Notes
}

function Build-FullReport {
    Add-Header
    Section-Environment
    Section-SiteOverview
    Section-Controllers
    Section-CatalogsAndGroups
    Section-Machines
    Section-Sessions
    Section-Policies
    Section-MCS
    Section-LocalChecks
    Section-LocalEvents
    Section-Notes
}

switch ($Action) {
    'QuickAudit'  { Build-QuickAudit }
    'SiteOverview'{ Build-SiteOverview }
    'Controllers' { Build-Controllers }
    'Catalogs'    { Build-Catalogs }
    'Machines'    { Build-Machines }
    'Sessions'    { Build-Sessions }
    'Policies'    { Build-Policies }
    'MCS'         { Build-MCS }
    'LocalChecks' { Build-LocalChecks }
    'LocalEvents' { Build-LocalEvents }
    'FullReport'  { Build-FullReport }
    default {
        Add-Header
        Add-Line ('Unknown action: ' + $Action)
        Add-Line
        Add-Line 'Valid actions: QuickAudit, SiteOverview, Controllers, Catalogs, Machines, Sessions, Policies, MCS, LocalChecks, LocalEvents, FullReport'
    }
}

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    foreach ($line in $script:Lines) {
        Write-Host $line
    }
}
else {
    try {
        $dir = Split-Path -Path $ReportPath -Parent
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
        $script:Lines | Set-Content -Path $ReportPath -Encoding UTF8
        Write-Host ('Report saved to: ' + $ReportPath)
    }
    catch {
        foreach ($line in $script:Lines) {
            Write-Host $line
        }
        Write-Host
        Write-Host ('Report save failed: ' + $_.Exception.Message)
    }
}
