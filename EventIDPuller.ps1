[CmdletBinding()]
param(
    [int]$DaysBack = 7,
    [int]$HoursBack = 0,
    [string[]]$ComputerName = $env:COMPUTERNAME,

    [switch]$IncludeHardware,
    [switch]$IncludeSoftware,
    [switch]$IncludeUserProfile,
    [switch]$IncludeGroupPolicy,
    [switch]$IncludeActiveDirectory,
    [switch]$IncludeDns,
    [switch]$IncludeDhcp,
    [switch]$IncludeDfsr,
    [switch]$IncludeSecurity,
    [switch]$IncludeNetwork,
    [switch]$IncludeAll,
    [switch]$IncludeNoise,

    [switch]$Detailed,
    [switch]$ShowConsoleEvents,

    # CSV is enabled by default. ExportCsv is kept for backward compatibility.
    [switch]$ExportCsv,
    [switch]$NoCsv,
    [switch]$ExportXlsx,
    [string]$CsvPath = "C:\Temp",
    [string]$XlsxPath = "",

    [int]$MaxEventsPerQuery = 0,
    [switch]$Help
)

function Show-QuickHelp {
    Write-Host ""
    Write-Host "Get-IssueEvents.ps1 usage examples" -ForegroundColor Cyan
    Write-Host "  .\Get-IssueEvents.ps1 -HoursBack 24"
    Write-Host "  .\Get-IssueEvents.ps1 -HoursBack 72 -IncludeAll"
    Write-Host "  .\Get-IssueEvents.ps1 -HoursBack 72 -IncludeAll -CsvPath C:\Temp"
    Write-Host "  .\Get-IssueEvents.ps1 -HoursBack 24 -IncludeAll -ExportXlsx"
    Write-Host "  .\Get-IssueEvents.ps1 -HoursBack 24 -IncludeAll -ExportXlsx -XlsxPath C:\Temp\EventReport.xlsx"
    Write-Host "  .\Get-IssueEvents.ps1 -HoursBack 24 -NoCsv -ShowConsoleEvents"
    Write-Host "  .\Get-IssueEvents.ps1 -ComputerName DC01 -HoursBack 72 -IncludeActiveDirectory -IncludeDns -IncludeDfsr -IncludeGroupPolicy -IncludeSecurity -CsvPath C:\Temp\DC_Report.csv -ExportXlsx"
    Write-Host ""
    Write-Host "CSV is on by default and saves to C:\Temp unless -CsvPath is used." -ForegroundColor Yellow
    Write-Host "Use -ExportXlsx to also create an Excel file. Excel must be installed for XLSX export."
    Write-Host "Matching event rows go to CSV/XLSX. Console shows summaries unless -ShowConsoleEvents or -Detailed is used."
    Write-Host ""
}

if ($Help) {
    Show-QuickHelp
    return
}

$CollectionTime = Get-Date

if ($HoursBack -gt 0) {
    $StartTime = $CollectionTime.AddHours(-$HoursBack)
    $RangeText = "Last $HoursBack hour(s)"
}
else {
    $StartTime = $CollectionTime.AddDays(-$DaysBack)
    $RangeText = "Last $DaysBack day(s)"
}

$UseCsv = -not $NoCsv
$CsvOutputFile = $null
$XlsxOutputFile = $null

if ($NoCsv -and -not $ShowConsoleEvents -and -not $Detailed -and -not $ExportXlsx) {
    $ShowConsoleEvents = $true
}

function Get-OutputFilePath {
    param(
        [string]$InputPath,
        [string]$DefaultFolder,
        [string]$Extension,
        [string]$Prefix,
        [datetime]$CollectionTime
    )

    $Timestamp = $CollectionTime.ToString("yyyyMMdd_HHmmss")

    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        $InputPath = $DefaultFolder
    }

    $InputPath = $InputPath.Trim('"')

    if ([System.IO.Path]::GetExtension($InputPath) -ieq $Extension) {
        $ParentFolder = Split-Path -Path $InputPath -Parent

        if (-not [string]::IsNullOrWhiteSpace($ParentFolder)) {
            if (-not (Test-Path -Path $ParentFolder -PathType Container)) {
                New-Item -Path $ParentFolder -ItemType Directory -Force | Out-Null
            }
        }

        return $InputPath
    }

    if (-not (Test-Path -Path $InputPath -PathType Container)) {
        New-Item -Path $InputPath -ItemType Directory -Force | Out-Null
    }

    return (Join-Path -Path $InputPath -ChildPath "$Prefix`_$Timestamp$Extension")
}

if ($UseCsv -or $ExportXlsx) {
    $CsvOutputFile = Get-OutputFilePath -InputPath $CsvPath -DefaultFolder "C:\Temp" -Extension ".csv" -Prefix "IssueEvents" -CollectionTime $CollectionTime
}

if ($ExportXlsx) {
    if ([string]::IsNullOrWhiteSpace($XlsxPath)) {
        $XlsxOutputFile = [System.IO.Path]::ChangeExtension($CsvOutputFile, ".xlsx")
    }
    else {
        $XlsxOutputFile = Get-OutputFilePath -InputPath $XlsxPath -DefaultFolder "C:\Temp" -Extension ".xlsx" -Prefix "IssueEvents" -CollectionTime $CollectionTime
    }
}

function Export-CsvToXlsx {
    param(
        [Parameter(Mandatory = $true)][string]$CsvFile,
        [Parameter(Mandatory = $true)][string]$XlsxFile
    )

    if (-not (Test-Path -Path $CsvFile -PathType Leaf)) {
        throw "CSV file was not found: $CsvFile"
    }

    $Excel = $null
    $Workbook = $null
    $Worksheet = $null

    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false

        $Workbook = $Excel.Workbooks.Open($CsvFile)
        $Worksheet = $Workbook.Worksheets.Item(1)

        $Worksheet.Columns.AutoFit() | Out-Null
        $Worksheet.Rows.AutoFit() | Out-Null

        $Workbook.SaveAs($XlsxFile, 51)
    }
    finally {
        if ($Workbook) {
            $Workbook.Close($false) | Out-Null
        }

        if ($Excel) {
            $Excel.Quit() | Out-Null
        }

        if ($Worksheet) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
        }

        if ($Workbook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
        }

        if ($Excel) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$NoCategorySwitches = -not (
    $IncludeHardware -or
    $IncludeSoftware -or
    $IncludeUserProfile -or
    $IncludeGroupPolicy -or
    $IncludeActiveDirectory -or
    $IncludeDns -or
    $IncludeDhcp -or
    $IncludeDfsr -or
    $IncludeSecurity -or
    $IncludeNetwork -or
    $IncludeAll
)

$UseHardware        = $IncludeAll -or $IncludeHardware -or $NoCategorySwitches
$UseSoftware        = $IncludeAll -or $IncludeSoftware -or $NoCategorySwitches
$UseUserProfile     = $IncludeAll -or $IncludeUserProfile -or $NoCategorySwitches
$UseGroupPolicy     = $IncludeAll -or $IncludeGroupPolicy -or $NoCategorySwitches
$UseActiveDirectory = $IncludeAll -or $IncludeActiveDirectory
$UseDns             = $IncludeAll -or $IncludeDns
$UseDhcp            = $IncludeAll -or $IncludeDhcp
$UseDfsr            = $IncludeAll -or $IncludeDfsr
$UseSecurity        = $IncludeAll -or $IncludeSecurity
$UseNetwork         = $IncludeAll -or $IncludeNetwork

$EventCatalog = @()
$LogExistsCache = @{}

function Add-EventPack {
    param(
        [Parameter(Mandatory = $true)][string]$IssueType,
        [Parameter(Mandatory = $true)][string]$LogName,
        [Parameter(Mandatory = $true)][string]$ProviderName,
        [Parameter(Mandatory = $true)][int[]]$Ids,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $script:EventCatalog += [pscustomobject]@{
        IssueType    = $IssueType
        LogName      = $LogName
        ProviderName = $ProviderName
        Ids          = $Ids
        Description  = $Description
    }
}

function Get-DiagnosticSeverity {
    param(
        [Parameter(Mandatory = $true)][int]$EventId,
        [string]$IssueType,
        [string]$Level
    )

    $CriticalIds = @(18,41,55,1988,2042,2095,1102,2213,4012,13568)

    $HighIds = @(
        7,11,15,51,129,153,154,
        1000,1002,11708,
        7000,7001,7009,7011,7022,7023,7024,7031,7034,7040,
        20,34,
        1500,1502,1505,1508,1511,1542,
        1030,1058,1084,1096,1129,
        1311,1566,1865,2087,2088,2108,
        5719,5722,5723,5805,
        4004,4013,4015,4016,
        4625,4740,4771,4776,
        5002,5014
    )

    if ($CriticalIds -contains $EventId) {
        return "Critical"
    }

    if ($HighIds -contains $EventId) {
        return "High"
    }

    if ($Level -eq "Critical") {
        return "Critical"
    }

    if ($Level -eq "Error") {
        return "High"
    }

    if ($Level -eq "Warning") {
        return "Medium"
    }

    return "Low"
}

function Get-SeverityRank {
    param([string]$Severity)

    switch ($Severity) {
        "Critical" { return 1 }
        "High"     { return 2 }
        "Medium"   { return 3 }
        "Low"      { return 4 }
        default     { return 5 }
    }
}

function Test-EventLogExists {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][string]$LogName
    )

    $CacheKey = "$Computer|$LogName"

    if ($script:LogExistsCache.ContainsKey($CacheKey)) {
        return $script:LogExistsCache[$CacheKey]
    }

    try {
        $null = Get-WinEvent -ComputerName $Computer -ListLog $LogName -ErrorAction Stop
        $script:LogExistsCache[$CacheKey] = $true
        return $true
    }
    catch {
        Write-Verbose ("Log not found or not accessible on {0}: {1}" -f $Computer, $LogName)
        $script:LogExistsCache[$CacheKey] = $false
        return $false
    }
}

function Get-IssueEventSafe {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][pscustomobject]$Entry,
        [Parameter(Mandatory = $true)][datetime]$StartTime,
        [Parameter(Mandatory = $true)][int]$MaxEventsPerQuery
    )

    if (-not (Test-EventLogExists -Computer $Computer -LogName $Entry.LogName)) {
        return
    }

    $Filter = @{
        LogName      = $Entry.LogName
        ProviderName = $Entry.ProviderName
        Id           = $Entry.Ids
        StartTime    = $StartTime
    }

    try {
        $QueryParams = @{
            ComputerName    = $Computer
            FilterHashtable = $Filter
            ErrorAction     = "Stop"
        }

        if ($MaxEventsPerQuery -gt 0) {
            $QueryParams.MaxEvents = $MaxEventsPerQuery
        }

        Get-WinEvent @QueryParams | ForEach-Object {
            $Severity = Get-DiagnosticSeverity -EventId $_.Id -IssueType $Entry.IssueType -Level $_.LevelDisplayName
            $SeverityRank = Get-SeverityRank -Severity $Severity

            [pscustomobject]@{
                TimeCreated  = $_.TimeCreated
                ComputerName = $Computer
                Severity     = $Severity
                SeverityRank = $SeverityRank
                Level        = $_.LevelDisplayName
                IssueType    = $Entry.IssueType
                EventId      = $_.Id
                LogName      = $_.LogName
                ProviderName = $_.ProviderName
                Description  = $Entry.Description
                Message      = if ($_.Message) { ($_.Message -replace "`r|`n", " ").Trim() } else { "" }
            }
        }
    }
    catch {
        Write-Verbose ("Skipped {0} / {1} / {2} / IDs {3}. {4}" -f $Computer, $Entry.LogName, $Entry.ProviderName, ($Entry.Ids -join ","), $_.Exception.Message)
    }
}

# Hardware, storage, power, and driver events
if ($UseHardware) {
    Add-EventPack -IssueType "Hardware" -LogName "System" -ProviderName "Microsoft-Windows-WHEA-Logger" -Ids @(1,17,18,19,47) -Description "CPU, memory, PCIe, motherboard, or other hardware error"
    Add-EventPack -IssueType "Hardware" -LogName "System" -ProviderName "Disk" -Ids @(7,11,15,51,153,157) -Description "Disk, controller, bad block, I/O timeout, or storage removal issue"
    Add-EventPack -IssueType "Hardware" -LogName "System" -ProviderName "Ntfs" -Ids @(55,98,130,140) -Description "NTFS volume corruption, repair, or file system issue"
    Add-EventPack -IssueType "Hardware" -LogName "System" -ProviderName "Microsoft-Windows-Kernel-Power" -Ids @(41) -Description "Unexpected shutdown, power loss, hard reset, or crash"
    Add-EventPack -IssueType "Hardware" -LogName "System" -ProviderName "EventLog" -Ids @(6008) -Description "Previous system shutdown was unexpected"
    Add-EventPack -IssueType "Hardware" -LogName "System" -ProviderName "Microsoft-Windows-Kernel-PnP" -Ids @(219,400,410,411) -Description "Plug and Play device or driver issue"
    Add-EventPack -IssueType "Storage" -LogName "System" -ProviderName "storahci" -Ids @(129,153,154) -Description "SATA/AHCI storage timeout, retry, or reset"
    Add-EventPack -IssueType "Storage" -LogName "System" -ProviderName "stornvme" -Ids @(129,153,154) -Description "NVMe storage timeout, retry, or reset"
    Add-EventPack -IssueType "Storage" -LogName "System" -ProviderName "storport" -Ids @(129,153,154) -Description "Storage port timeout, retry, or reset"
    Add-EventPack -IssueType "Graphics" -LogName "System" -ProviderName "Display" -Ids @(4101) -Description "Display driver stopped responding and recovered"
    Add-EventPack -IssueType "Driver" -LogName "System" -ProviderName "Microsoft-Windows-DriverFrameworks-UserMode" -Ids @(10110,10111,10116) -Description "User-mode driver crash, hang, or device issue"
}

# Software, services, and Windows Update events
if ($UseSoftware) {
    Add-EventPack -IssueType "Software" -LogName "Application" -ProviderName "Application Error" -Ids @(1000) -Description "Application crash"
    Add-EventPack -IssueType "Software" -LogName "Application" -ProviderName "Application Hang" -Ids @(1002) -Description "Application stopped responding"
    Add-EventPack -IssueType "Software" -LogName "Application" -ProviderName "Windows Error Reporting" -Ids @(1001) -Description "Windows Error Reporting crash or fault bucket"
    Add-EventPack -IssueType "Software" -LogName "Application" -ProviderName "MsiInstaller" -Ids @(1023,11708) -Description "Software install, uninstall, or MSI failure"
    Add-EventPack -IssueType "Software" -LogName "System" -ProviderName "Service Control Manager" -Ids @(7000,7001,7009,7011,7022,7023,7024,7031,7034,7040,7045) -Description "Windows service failed, crashed, timed out, changed startup type, or was installed"
    Add-EventPack -IssueType "Windows Update" -LogName "System" -ProviderName "Microsoft-Windows-WindowsUpdateClient" -Ids @(20,34) -Description "Windows Update failure or service issue"
}

# User profile events
if ($UseUserProfile) {
    Add-EventPack -IssueType "User/Profile" -LogName "Application" -ProviderName "Microsoft-Windows-User Profiles Service" -Ids @(1500,1501,1502,1504,1505,1508,1509,1511,1515,1530,1533,1534,1542) -Description "User profile load, unload, temp profile, registry hive, or cleanup issue"
    Add-EventPack -IssueType "User/Profile" -LogName "Microsoft-Windows-User Profile Service/Operational" -ProviderName "Microsoft-Windows-User Profiles Service" -Ids @(1500,1501,1502,1504,1505,1508,1509,1511,1515,1530,1533,1534,1542) -Description "User profile operational load, unload, temporary profile, registry hive, or cleanup issue"
}

# Group Policy events
if ($UseGroupPolicy) {
    Add-EventPack -IssueType "Group Policy" -LogName "System" -ProviderName "Microsoft-Windows-GroupPolicy" -Ids @(1030,1055,1058,1085,1096,1129) -Description "Group Policy processing, SYSVOL, network, LDAP, or DC connectivity issue"
    Add-EventPack -IssueType "Group Policy" -LogName "Microsoft-Windows-GroupPolicy/Operational" -ProviderName "Microsoft-Windows-GroupPolicy" -Ids @(4003,4016,5016,7016,7017,7320,8006,8007) -Description "Group Policy warning, failure, or processing issue"
}

# Active Directory and domain controller events
if ($UseActiveDirectory) {
    Add-EventPack -IssueType "Active Directory" -LogName "Directory Service" -ProviderName "Microsoft-Windows-ActiveDirectory_DomainService" -Ids @(1311,1566,1865,1084,2108,1988,2042,2095,2087,2088,1644,2886,2887,2888,2889) -Description "AD DS replication, KCC topology, LDAP, lingering object, or USN rollback issue"
    Add-EventPack -IssueType "Active Directory" -LogName "System" -ProviderName "NETLOGON" -Ids @(5719,5722,5723,5727,5774,5781,5805) -Description "DC discovery, secure channel, trust, or DNS registration issue"
    Add-EventPack -IssueType "Active Directory" -LogName "System" -ProviderName "Microsoft-Windows-Time-Service" -Ids @(29,36,47,50,129,134) -Description "Time sync issue affecting Kerberos or domain authentication"
}

# DNS events
if ($UseDns) {
    Add-EventPack -IssueType "DNS" -LogName "DNS Server" -ProviderName "Microsoft-Windows-DNS-Server-Service" -Ids @(4000,4004,4007,4010,4013,4015,4016,4515,4521) -Description "DNS service, AD-integrated DNS, zone loading, LDAP timeout, or DNS replication issue"
    Add-EventPack -IssueType "DNS Client" -LogName "System" -ProviderName "Microsoft-Windows-DNS-Client" -Ids @(1014) -Description "DNS client name resolution timeout"
}

# DHCP events
if ($UseDhcp) {
    Add-EventPack -IssueType "DHCP" -LogName "Microsoft-Windows-DHCP Server Events/Admin" -ProviderName "Microsoft-Windows-DHCP-Server" -Ids @(1001,1002,1003,1004,1045,1046,1051,1052,1053,1054,1055,1056,1058,1060,1061,1062,1063,1064,1144,1341,1342,1376,1377,20090) -Description "DHCP authorization, database, DNS credentials, scope exhaustion, audit, backup, or service issue"
    Add-EventPack -IssueType "DHCP" -LogName "Microsoft-Windows-DHCP Server Events/Operational" -ProviderName "Microsoft-Windows-DHCP-Server" -Ids @(20058,20059,20060,20061,20062,20063,20064,20065,20066,20067,20286,20287,20288,20289,20290,20291,20292,20293,20294,20295,20296,20297,20298,20299,20300,20301,20302,20303,20304,20305,20306,20307,20308,20309,20310) -Description "DHCP failover, partner communication, failover state, or lease sync issue"
}

# DFSR, SYSVOL, and legacy FRS events
if ($UseDfsr) {
    Add-EventPack -IssueType "DFSR / SYSVOL" -LogName "DFS Replication" -ProviderName "DFSR" -Ids @(2212,2213,4012,4202,4204,4206,4208,4602,4604,4612,5002,5004,5012,5014) -Description "DFSR, SYSVOL, staging, replication partner, or recovery issue"
    Add-EventPack -IssueType "FRS / Legacy SYSVOL" -LogName "File Replication Service" -ProviderName "NtFrs" -Ids @(13508,13509,13516,13522,13568) -Description "Legacy FRS SYSVOL replication issue"
}

# Network events
if ($UseNetwork) {
    Add-EventPack -IssueType "Network" -LogName "System" -ProviderName "Tcpip" -Ids @(4199,4201,4231) -Description "TCP/IP conflict, adapter binding, or port exhaustion issue"
    Add-EventPack -IssueType "Network" -LogName "System" -ProviderName "NlaSvc" -Ids @(4002,4004) -Description "Network Location Awareness or domain network detection issue"
}

# Security, authentication, and account events
if ($UseSecurity) {
    Add-EventPack -IssueType "Authentication" -LogName "Security" -ProviderName "Microsoft-Windows-Security-Auditing" -Ids @(4625,4648,4771,4776,4740) -Description "Failed logon, explicit credentials, Kerberos failure, NTLM failure, or account lockout"
    Add-EventPack -IssueType "Account Management" -LogName "Security" -ProviderName "Microsoft-Windows-Security-Auditing" -Ids @(4720,4722,4723,4724,4725,4726,4728,4729,4732,4733,4738,4741,4742,4743,4756,4757) -Description "User, computer, or group account change"
    Add-EventPack -IssueType "Security Monitoring" -LogName "Security" -ProviderName "Microsoft-Windows-Security-Auditing" -Ids @(4719,1102) -Description "Audit policy changed or Security log cleared"
}

# Optional noisy/general events
if ($IncludeNoise) {
    if ($UseHardware) {
        Add-EventPack -IssueType "Hardware / General" -LogName "System" -ProviderName "EventLog" -Ids @(6005,6006) -Description "Event log service started or stopped"
        Add-EventPack -IssueType "Power / General" -LogName "System" -ProviderName "Microsoft-Windows-Kernel-Power" -Ids @(109,172) -Description "Power transition, sleep, or hibernation activity"
        Add-EventPack -IssueType "Driver / General" -LogName "System" -ProviderName "Microsoft-Windows-DriverFrameworks-UserMode" -Ids @(10114) -Description "User-mode driver reflector startup or transient device issue"
    }

    if ($UseSoftware) {
        Add-EventPack -IssueType "Software / General" -LogName "Application" -ProviderName "MsiInstaller" -Ids @(1033,11707) -Description "Successful software installation events"
        Add-EventPack -IssueType "Windows Update / General" -LogName "System" -ProviderName "Microsoft-Windows-WindowsUpdateClient" -Ids @(25,31,41,43) -Description "General Windows Update activity"
    }

    if ($UseGroupPolicy) {
        Add-EventPack -IssueType "Group Policy / General" -LogName "Microsoft-Windows-GroupPolicy/Operational" -ProviderName "Microsoft-Windows-GroupPolicy" -Ids @(4000,4001,8000,8001,8004,8005) -Description "General or successful Group Policy processing activity"
    }

    if ($UseNetwork) {
        Add-EventPack -IssueType "Network / General" -LogName "System" -ProviderName "Microsoft-Windows-NetworkProfile" -Ids @(10000,10001) -Description "Network connected or disconnected"
    }

    if ($UseSecurity) {
        Add-EventPack -IssueType "Authentication / General" -LogName "Security" -ProviderName "Microsoft-Windows-Security-Auditing" -Ids @(4624,4634,4647,4672,4768,4769,4770) -Description "Successful logon, logoff, privileged logon, or normal Kerberos ticket activity"
    }
}

if (-not $EventCatalog -or $EventCatalog.Count -eq 0) {
    Write-Warning "No event categories selected."
    return
}

Write-Host ""
Write-Host "Event Diagnostic Scan" -ForegroundColor Cyan
Write-Host "Range      : $RangeText"
Write-Host "Start Time : $StartTime"
Write-Host "End Time   : $CollectionTime"
Write-Host "Computers  : $($ComputerName -join ', ')"
Write-Host "CSV Output : $UseCsv"

if ($UseCsv -or $ExportXlsx) {
    Write-Host "CSV Path   : $CsvOutputFile"
}

Write-Host "XLSX Output: $ExportXlsx"

if ($ExportXlsx) {
    Write-Host "XLSX Path  : $XlsxOutputFile"
}

Write-Host ""

$Results = foreach ($Computer in $ComputerName) {
    Write-Host "Scanning $Computer..." -ForegroundColor Cyan

    foreach ($Entry in $EventCatalog) {
        Get-IssueEventSafe -Computer $Computer -Entry $Entry -StartTime $StartTime -MaxEventsPerQuery $MaxEventsPerQuery
    }
}

if (-not $Results) {
    Write-Host ""
    Write-Host "No matching events found." -ForegroundColor Yellow
    return
}

$SortedResults = $Results | Sort-Object `
    @{Expression = "ComputerName"; Ascending = $true},
    @{Expression = "SeverityRank"; Ascending = $true},
    @{Expression = "TimeCreated"; Descending = $true}

if ($UseCsv -or $ExportXlsx) {
    try {
        if ([string]::IsNullOrWhiteSpace($CsvOutputFile)) {
            $CsvOutputFile = Get-OutputFilePath -InputPath $CsvPath -DefaultFolder "C:\Temp" -Extension ".csv" -Prefix "IssueEvents" -CollectionTime $CollectionTime
        }

        $CsvLines = New-Object System.Collections.Generic.List[string]

        foreach ($Computer in ($SortedResults.ComputerName | Sort-Object -Unique)) {
            $ComputerEvents = $SortedResults |
                Where-Object { $_.ComputerName -eq $Computer } |
                Sort-Object `
                    @{Expression = "SeverityRank"; Ascending = $true},
                    @{Expression = "TimeCreated"; Descending = $true}

            if ($CsvLines.Count -gt 0) {
                $CsvLines.Add("")
                $CsvLines.Add("")
            }

            $CsvLines.Add("ReportGenerated: $($CollectionTime.ToString('yyyy-MM-dd HH:mm:ss'))")
            $CsvLines.Add("SearchStart: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))")
            $CsvLines.Add("SearchEnd: $($CollectionTime.ToString('yyyy-MM-dd HH:mm:ss'))")
            $CsvLines.Add("Range: $RangeText")
            $CsvLines.Add("ComputerName: $Computer")
            $CsvLines.Add("")

            $ComputerEvents |
                Select-Object `
                    TimeCreated,
                    Severity,
                    Level,
                    IssueType,
                    EventId,
                    LogName,
                    ProviderName,
                    Description,
                    Message |
                ConvertTo-Csv -NoTypeInformation |
                ForEach-Object {
                    $CsvLines.Add($_)
                }
        }

        Set-Content -Path $CsvOutputFile -Value $CsvLines -Encoding UTF8 -ErrorAction Stop

        if ($UseCsv) {
            Write-Host ""
            Write-Host "CSV exported to:" -ForegroundColor Green
            Write-Host $CsvOutputFile -ForegroundColor Green
        }

        if ($ExportXlsx) {
            if ([string]::IsNullOrWhiteSpace($XlsxOutputFile)) {
                $XlsxOutputFile = [System.IO.Path]::ChangeExtension($CsvOutputFile, ".xlsx")
            }

            Export-CsvToXlsx -CsvFile $CsvOutputFile -XlsxFile $XlsxOutputFile

            Write-Host ""
            Write-Host "XLSX exported to:" -ForegroundColor Green
            Write-Host $XlsxOutputFile -ForegroundColor Green
        }
    }
    catch {
        Write-Host ""
        Write-Warning ("Failed to export report. {0}" -f $_.Exception.Message)
        Write-Host ("CsvPath value was: {0}" -f $CsvPath) -ForegroundColor Yellow
        Write-Host ("CsvOutputFile value was: {0}" -f $CsvOutputFile) -ForegroundColor Yellow
        Write-Host ("XlsxPath value was: {0}" -f $XlsxPath) -ForegroundColor Yellow
        Write-Host ("XlsxOutputFile value was: {0}" -f $XlsxOutputFile) -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Summary by Severity" -ForegroundColor Green
$SortedResults |
    Group-Object Severity |
    Select-Object Name, Count |
    Sort-Object @{Expression = { Get-SeverityRank -Severity $_.Name }; Ascending = $true} |
    Format-Table -AutoSize

Write-Host ""
Write-Host "Summary by Issue Type" -ForegroundColor Green
$SortedResults |
    Group-Object IssueType |
    Select-Object Name, Count |
    Sort-Object Count -Descending |
    Format-Table -AutoSize

Write-Host ""
Write-Host "Summary by Event ID" -ForegroundColor Green
$SortedResults |
    Group-Object EventId, IssueType, Severity |
    Select-Object `
        @{Name = "EventId"; Expression = { $_.Group[0].EventId }},
        @{Name = "Severity"; Expression = { $_.Group[0].Severity }},
        @{Name = "IssueType"; Expression = { $_.Group[0].IssueType }},
        Count |
    Sort-Object Count -Descending |
    Format-Table -AutoSize

if ($Detailed) {
    Write-Host ""
    Write-Host "Matching Events" -ForegroundColor Green
    $SortedResults |
        Select-Object TimeCreated, ComputerName, Severity, Level, IssueType, EventId, LogName, ProviderName, Description, Message |
        Format-List
}
elseif ($ShowConsoleEvents) {
    Write-Host ""
    Write-Host "Matching Events" -ForegroundColor Green
    $SortedResults |
        Select-Object `
            TimeCreated,
            ComputerName,
            Severity,
            Level,
            IssueType,
            EventId,
            LogName,
            ProviderName,
            @{Name = "Message"; Expression = {
                if ($_.Message.Length -gt 120) {
                    $_.Message.Substring(0,120) + "..."
                }
                else {
                    $_.Message
                }
            }} |
        Format-Table -AutoSize -Wrap
}
else {
    Write-Host ""
    Write-Host "Matching event rows were written to the report. Use -ShowConsoleEvents or -Detailed if you also want them printed in the console." -ForegroundColor Yellow
}
