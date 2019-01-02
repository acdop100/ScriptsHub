#region Functions
#### HTML Output Formatting #######
################################################################################################
#### Functions ####
################################################################################################

param([Int]$days, [string]$to)

write-host "Going back $days days; sending report to $to" -ForegroundColor Green

Function Set-CellColor {
    
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory, Position = 0)]
        [string]$Property,
        [Parameter(Mandatory, Position = 1)]
        [string]$Color,
        [Parameter(Mandatory, ValueFromPipeline)]
        [Object[]]$InputObject,
        [Parameter(Mandatory)]
        [string]$Filter,
        [switch]$Row
    )
    
    Begin {
        Write-Verbose "$(Get-Date): Function Set-CellColor begins"
        If ($Filter) {
            If ($Filter.ToUpper().IndexOf($Property.ToUpper()) -ge 0) {
                $Filter = $Filter.ToUpper().Replace($Property.ToUpper(), "`$Value")
                Try {
                    [scriptblock]$Filter = [scriptblock]::Create($Filter)
                }
                Catch {
                    Write-Warning "$(Get-Date): ""$Filter"" caused an error, stopping script!"
                    Write-Warning $Error[0]
                    Exit
                }
            }
            Else {
                Write-Warning "Could not locate $Property in the Filter, which is required.  Filter: $Filter"
                Exit
            }
        }
    }
    
    Process {
        ForEach ($Line in $InputObject) {
            If ($Line.IndexOf("<tr><th") -ge 0) {
                Write-Verbose "$(Get-Date): Processing headers..."
                $Search = $Line | Select-String -Pattern '<th ?[a-z\-:;"=]*>(.*?)<\/th>' -AllMatches
                $Index = 0
                ForEach ($Match in $Search.Matches) {
                    If ($Match.Groups[1].Value -eq $Property) {
                        Break
                    }
                    $Index ++
                }
                If ($Index -eq $Search.Matches.Count) {
                    Write-Warning "$(Get-Date): Unable to locate property: $Property in table header"
                    Exit
                }
                Write-Verbose "$(Get-Date): $Property column found at index: $Index"
            }
            If ($Line -match "<tr( style=""background-color:.+?"")?><td") {
                $Search = $Line | Select-String -Pattern '<td ?[a-z\-:;"=]*>(.*?)<\/td>' -AllMatches
                $Value = $Search.Matches[$Index].Groups[1].Value -as [double]
                If (-not $Value) {
                    $Value = $Search.Matches[$Index].Groups[1].Value
                }
                If (Invoke-Command $Filter) {
                    If ($Row) {
                        Write-Verbose "$(Get-Date): Criteria met!  Changing row to $Color..."
                        If ($Line -match "<tr style=""background-color:(.+?)"">") {
                            $Line = $Line -replace "<tr style=""background-color:$($Matches[1])", "<tr style=""background-color:$Color"
                        }
                        Else {
                            $Line = $Line.Replace("<tr>", "<tr style=""background-color:$Color"">")
                        }
                    }
                    Else {
                        Write-Verbose "$(Get-Date): Criteria met!  Changing cell to $Color..."
                        $Line = $Line.Replace($Search.Matches[$Index].Value, "<td style=""background-color:$Color"">$Value</td>")
                    }
                }
            }
            Write-Output $Line
        }
    }
    
    End {
        Write-Verbose "$(Get-Date): Function Set-CellColor completed"
    }
}

function Get-LoggedOnUser {
    param([String[]]$ComputerName = $env:COMPUTERNAME)

    $ComputerName | ForEach-Object {
        (quser /SERVER:$_) -replace '\s{2,}', ',' | 
            ConvertFrom-CSV |
            Add-Member -MemberType NoteProperty -Name ComputerName -Value $_ -PassThru
    }
} 

#endregion

#region Styles and Variables
$a = @" 
<style>
BODY{background-color:white; font-family:Calibri;font-size:12pt;}
.main{margin-right: 140px}
TABLE{display: none;border-width: 1px;border-style: solid;font-family:Calibri;font-size:10pt;border-color: black;border-collapse: collapse; border-radius: 5px;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:rgb(102, 145, 180)}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:rgb(255, 220, 75)}
nth-child(even) {background-color:#ffffff;}
.collapsible {background-color: #ddd;color: #000000;cursor: pointer;padding: 15px;width: 100%;border: none;text-align: left;outline: none;font-size: 15px; border-radius: 5px; margin: 5px;}
.active, .collapsible:hover {background-color: #ccc;}
.content {padding: 0 18px;display: none;overflow: hidden;background-color: #f1f1f1;}
.sidenav {width: 130px;position: fixed;z-index: 1;top: 20px;right: 10px;background: #eee;overflow-x: hidden;padding: 8px 0;}
.sidenav a {padding: 6px 8px 6px 16px; text-decoration: none;font-size: 25px;color: #2196F3;display: block;}
.sidenav a:hover {color: #064579;}
</style>
"@
#.even { background-color:#dddddd; }

################################################################################################
#### Global variables ####
################################################################################################

$options = Select-Xml -Path ".\config.xml" -XPath ".\GUI_Config\MegaScript_Config\*" | foreach {$_.node.InnerXML}
$date = (Get-Date).AddDays( - $days)
$startDTM = (Get-Date)
$ErrorActionPreference = 'SilentlyContinue'
$filepath = "C:\Users\$env:UserName\Desktop"		## this is user profile  using environment variable

$vComputerName = Get-Content -Path .\PCList.txt


################################################################################################
#### Get Imputs ####
################################################################################################

For ($i = 0; $i -le ($vComputerName.Length - 2); $i++) {

    $currentPC = $vComputerName[$i] #Using just $vComputerName[$i] will break the report file

    If (-not (Test-Connection -Computername $currentPC -BufferSize 16 -Count 1 -Quiet)) {  
        (Write-Host "$currentPC is Not online!"  -ForegroundColor Red)
        Break Script
    }
    Else {
        (Write-Host "$currentPC is responding."  -ForegroundColor Green) 
    }

    ################################################
    #  Get Logged on User
    #################################################

    write-host "Step 1: Gathering data" -ForegroundColor Green

    IF ([string]::IsNullOrWhitespace($currentPC)) {$currentPC = (Get-Item env:\Computername).Value} else {}

    ConvertTo-Html  -Precontent "
    <img style='display: block;margin-left: auto;margin-right: auto;margin-bottom:18px;'src='https://www.nelsonmullins.com/img/nm_logo.png'> 
    <button type='button' id='collapseBtn' style='float: right; margin: 10px 10px;'>Collapse All</button> 
    <button type='button' id='expandBtn' style='float: right; margin: 10px 10px;'>Expand All</button>" `
    -Title "Startup-Shutdown Information for $currentPC" -Body "<h1 style='text-align:center;'> Computer Name : $currentPC </h1>" | Out-File  "$filepath\$currentPC.html" 

    $ComputerOwner = Get-LoggedOnUser -Computername $currentPC | Select-Object -Property UserName
    Get-LoggedOnUser -Computername $currentPC | Select-Object -Property UserName, SessionName, 'Logon Time', 'Idle Time' `
        | ConvertTo-html  -Body "<button class='collapsible'>Machine User </button>" >> "$filepath\$currentPC.html" -EA SilentlyContinue
    
    $computeruser = Get-LoggedOnUser -Computername $currentPC | Select-Object -Property UserName, SessionName, 'Logon Time', 'Idle Time' -EA SilentlyContinue
    

    ################################################
    #  Get Image Version and Bios
    #################################################

    if ($options[0] -eq 1) {
        
        ConvertTo-Html -Body "<H1>Image & BIOS Information</H1>" >> "$filepath\$currentPC.html"

        Invoke-Command -cn $currentPC -ScriptBlock {Get-ItemProperty HKLM:\Software\* | Select-Object CaptureDate, version, BuildDate} | Select-Object -ExcludeProperty PSComputerName, RunspaceId | Where-Object {!(($_.Version -eq $null))} `
            | ConvertTo-html -Head $a -Body "<button class='collapsible'>Image Version and Free Memory</button>" >> "$filepath\$currentPC.html" -EA SilentlyContinue

        Get-WmiObject win32_operatingsystem -ComputerName $currentPC | Select-Object Caption, Organization, @{LABEL = 'InstallDate'; EXPRESSION = {$_.ConverttoDateTime($_.InstallDate)}}, OSArchitecture, Version, SerialNumber, BootDevice, WindowsDirectory, CountryCode `
            | ConvertTo-html -Body "<button class='collapsible'>System Info</button>" >> "$filepath\$currentPC.html" -EA SilentlyContinue

        Get-WmiObject win32_bios -ComputerName $currentPC | Select-Object Status, Version, PrimaryBIOS, Manufacturer, @{LABEL = 'ReleaseDate'; EXPRESSION = {$_.ConverttoDateTime($_.ReleaseDate)}}, SerialNumber `
            | ConvertTo-html -Body "<button class='collapsible'>BIOS Info</button>" >> "$filepath\$currentPC.html" -EA SilentlyContinue
    }
    ################################################
    #  Get last reboot, Free Memory, Processor, and disk info
    #################################################

    if ($options[1] -eq 1) {

        ConvertTo-Html -Body "<H1>Last Reboot Information</H1>" >> "$filepath\$currentPC.html"

        Get-WmiObject -Class win32_operatingsystem -ComputerName $currentPC | Select-Object @{LABEL = 'LastBootUpTime'; EXPRESSION = {$_.ConverttoDateTime($_.lastbootuptime)}}, FreePhysicalMemory `
            | ConvertTo-html -Body "<button class='collapsible'>Last Reboot Time and Free RAM</button>" >> "$filepath\$currentPC.html" -EA SilentlyContinue

        Invoke-WmiMethod -ComputerName $currentPC -Namespace "ROOT\ccm\ClientSDK" -Class "CCM_ClientUtilities" -Name DetermineIfRebootPending | Select-Object -Property RebootPending `
            | ConvertTo-html  -Head $a -Body "<button class='collapsible'> Pending Updates Reboot?</button>" >>  "$filepath\$currentPC.html"
    }

    if ($options[2] -eq 1) {

        ConvertTo-Html -Body "<H1>Hardware Information</H1>" >> "$filepath\$currentPC.html"

        Get-WmiObject -Class Win32_ComputerSystem -ComputerName $currentPC | Select-Object Status, Manufacturer, SystemFamily, Model, Name, TotalPhysicalMemory `
            | ConvertTo-html -Body "<button class='collapsible'>PC Model Info</button>" | Set-CellColor TotalPhysicalMemory green -Filter "TotalPhysicalMemory -gt 7000000000" | Set-CellColor TotalPhysicalMemory red -Filter "TotalPhysicalMemory -lt 7000000000" >> "$filepath\$currentPC.html" -EA SilentlyContinue #| Set-CellColor TotalPhysicalMemory green -Filter "TotalPhysicalMemory -gt 7000000000" | Set-CellColor TotalPhysicalMemory red -Filter "TotalPhysicalMemory -lt 7000000000" 

        Get-WmiObject win32_processor -computername $currentPC  | Measure-Object -property LoadPercentage -Average | Select-Object @{Name = "CPU Load"; Expression = {$_.Average}}  `
            | ConvertTo-html -Body "<button class='collapsible'> CPU Load </button>" >> "$filepath\$currentPC.html" -EA SilentlyContinue # -Property 'CPU Load' -Fragment | Set-CellColor "CPU Load" red -Filter "CPU Load -gt 50" | Set-CellColor "CPU Load" green -Filter "CPU Load -lt 50" 

        Get-Counter -ComputerName $currentPC '\Process(*)\% Processor Time' `
            | Select-Object -ExpandProperty countersamples `
            | Select-Object -Property instancename, cookedvalue `
            | Sort-Object -Property cookedvalue -Descending | Select-Object -First 10 InstanceName, @{L = 'CPU'; E = {($_.Cookedvalue / 100).toString('P')}} `
            | ConvertTo-html -Body "<button class='collapsible'> Current Process Information </button>" `
            >> "$filepath\$currentPC.html" -EA SilentlyContinue #| Set-CellColor "CPU" red -Filter "CPU -gt 80" 

        Invoke-Command -cn $currentPC -ScriptBlock {Get-ChildItem "HKLM:\SOFTWARE\Intel\Setup and Configuration Software\SystemDiscovery" | ForEach-Object {Get-ItemProperty $_.pspath |Select-Object IsATEnabledInBios, SystemDataVersion, LastTimeUpdated, SCSVersion, BIOSVersion, AMTSKU, AMTVersion, FWVersion, MEIVersion, IsMEIEnabled, LMSVersion, SBAState}} `
            | ConvertTo-html  -Head $a -Body "<button class='collapsible'> INTEL AMT (If Applicable)</button>" >>  "$filepath\$currentPC.html" -EA SilentlyContinue

        Invoke-Command -cn $currentPC -ScriptBlock {Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Intel\Setup and Configuration Software\SystemDiscovery" | ForEach-Object {Get-ItemProperty $_.pspath | Select-Object IsATEnabledInBios, SystemDataVersion, LastTimeUpdated, SCSVersion, BIOSVersion, AMTSKU, AMTVersion, FWVersion, MEIVersion, IsMEIEnabled, LMSVersion, SBAState}} `
            | ConvertTo-html -Body "<button class='collapsible'> More Intel AMT (If Applicable) </button>" >>  "$filepath\$currentPC.html" -EA SilentlyContinue

        Get-WmiObject win32_process -ComputerName $currentPC | Select-Object Caption, ProcessId, @{Expression = {$_.Vm / 1mb -as [Int]}; Label = "VM (MB)"}, @{Expression = {$_.Ws / 1Mb -as [Int]}; Label = "WS (MB)"} |Sort-Object "WS (MB)" -Descending | Select-Object -first 15 `
            | ConvertTo-html  -Head $a -Body "<button class='collapsible'> Process Memory Usage </button>" | Set-CellColor "WS (MB)" red -Filter "WS (MB) -gt 150" >> "$filepath\$currentPC.html" -EA SilentlyContinue #

        Invoke-WmiMethod -ComputerName $currentPC -Namespace "ROOT\ccm\ClientSDK" -Class "CCM_ClientUtilities" -Name DetermineIfRebootPending | Select-Object -Property RebootPending `
            | ConvertTo-html -Property RebootPending -Fragment Body "<button class='collapsible'>Reboot pending?</button>" | Set-CellColor "RebootPending" red -Filter "RebootPending -eq 'True'" | Set-CellColor "RebootPending" green -Filter "RebootPending -eq 'False'" >> "$filepath\$currentPC.html" -EA SilentlyContinue #

        Get-WmiObject win32_DiskDrive -ComputerName $currentPC | Select-Object Model, SerialNumber, Description, MediaType, FirmwareRevision `
            | ConvertTo-html -Body "<button class='collapsible'> Disk Information </button>" >>  "$filepath\$currentPC.html" 										  

        Get-WmiObject win32_logicalDisk -ComputerName $currentPC | Select-Object DeviceID, VolumeName, Freespace, @{Expression = {$_.Size / 1Gb -as [int]}; Label = "Total Size - GB"}, @{Expression = {$_.Freespace / 1Gb -as [int]}; Label = "Free Size - GB"} `
            | ConvertTo-html -Body "<button class='collapsible'> Physcial Drives </button>" | Set-CellColor Freespace red -Filter "Freespace -lt 20000000000" | Set-CellColor Freespace green -Filter "Freespace -gt 30000000000">> "$filepath\$currentPC.html"

    }

    ################################################
    #  Display HTTP Response Time
    ################################################

    $URL1 = "http://one.nmrs.com/Pages/Firm%20View.aspx"
    $URL2 = "https://app.chromeriver.com/login"
    $URL3 = "https://www.yahoo.com/"
    $URL4 = "http://www.msnbc.com/"

    if ($options[3] -eq 1) {

        ConvertTo-Html -Body "<H1>HTTP Response Information</H1>" >> "$filepath\$currentPC.html"

        $timeTaken = Measure-Command -Expression {
        }
        $milliseconds = $timeTaken.TotalMilliseconds
        $milliseconds = [Math]::Round($milliseconds, 0)
        $Context.SetValue($milliseconds);
        write-host Website download time to $URL1 took -ForegroundColor Green 
        write-host $milliseconds Milliseconds -ForegroundColor Yellow | ConvertTo-html -Head $a -Body "<H3> HTTP Test 1 - $URL1 --- $milliseconds ms</H3>" >>  "$filepath\$currentPC.html" 	

        $timeTaken2 = Measure-Command -Expression {
        }
        $milliseconds2 = $timeTaken2.TotalMilliseconds
        $milliseconds2 = [Math]::Round($milliseconds2, 0)
        $Context.SetValue($milliseconds2);
        write-host Website download time to $url2 took -ForegroundColor Green 
        write-host $milliseconds2 Milliseconds -ForegroundColor Yellow | ConvertTo-html -Body "<H3> HTTP Test 2 - $URL2 --- $milliseconds2 ms</H3>" >>  "$filepath\$currentPC.html"
        
        $timeTaken3 = Measure-Command -Expression {
        }
        $milliseconds3 = $timeTaken3.TotalMilliseconds
        $milliseconds3 = [Math]::Round($milliseconds3, 0)
        $Context.SetValue($milliseconds3);
        write-host Website download time to $url3 took -ForegroundColor Green 
        write-host $milliseconds3 Milliseconds -ForegroundColor Yellow | ConvertTo-html -Body "<H3> HTTP Test 3 - $URL3 --- $milliseconds3 ms</H3>" >>  "$filepath\$currentPC.html"

        $timeTaken4 = Measure-Command -Expression {
        }
        $milliseconds4 = $timeTaken4.TotalMilliseconds
        $milliseconds4 = [Math]::Round($milliseconds4, 0)
        $Context.SetValue($milliseconds4);
        write-host Website download time to $url4 took -ForegroundColor Green 
        write-host $milliseconds4 Milliseconds -ForegroundColor Yellow | ConvertTo-html -Body "<H3> HTTP Test 4 - $URL4 --- $milliseconds4 ms</H3>" >>  "$filepath\$currentPC.html"

    }
    ################################################
    #  Power Plan
    ################################################
    
    if ($options[4] -eq 1) {

        ConvertTo-Html -Body "<H1>Power Plan Information </H1>" >> "$filepath\$currentPC.html"

        Get-WmiObject -computername $currentPC -Class win32_powerplan -Namespace root\cimv2\power | Select-Object Description, ElementName, IsActive  `
            | ConvertTo-html -Body "<button class='collapsible'>Power Plan Settings </button>" `
            | Set-CellColor IsActive green -Filter "IsActive -eq 'True'" >> "$filepath\$currentPC.html" -ErrorAction SilentlyContinue
        $sw.Stop()

        Invoke-Command -cn $currentPC -scriptblock {Get-NetAdapter} | Select-Object MacAddress, LinkSpeed, MediaConnectionState, ifOperStatus, ifAlias, ActiveMaximumTransmissionUnit, DriverVersion, DriverInformation |
            ConvertTo-html -Head $a -Body "<button class='collapsible'>Power Plan: Network Adapters </button>" | Set-CellColor "ifOperStatus" green -Filter "ifOperStatus -like 'UP'" >> "$filepath\$currentPC.html" -EA SilentlyContinue 

    }
    ################################################
    #  Updates and Hotfixes
    ################################################

    if ($options[5] -eq 1) {
        ConvertTo-Html -Body "<H1>Updates & Hotfixes</H1>" >> "$filepath\$currentPC.html"

        Get-WmiObject -ComputerName $currentPC -Class CCM_SoftwareUpdate -Filter ComplianceState=0 -Namespace root\CCM\ClientSDK | Select-Object Name, ArticleID, @{LABEL = 'StartTime'; EXPRESSION = {$_.ConverttoDateTime($_.StartTime)}}, @{LABEL = 'Deadline'; EXPRESSION = {$_.ConverttoDateTime($_.Deadline)}} `
            | ConvertTo-html  -Head $a -Body "<button class='collapsible'> Pending Updates From SCCM (If Applicable)</button>" >>  "$filepath\$currentPC.html"
    
        Invoke-WmiMethod -ComputerName $currentPC -Namespace "ROOT\ccm\ClientSDK" -Class "CCM_ClientUtilities" -Name DetermineIfRebootPending | Select-Object -Property RebootPending `
            | ConvertTo-html  -Head $a -Body "<button class='collapsible'> Pending Updates For Reboot</button>" >>  "$filepath\$currentPC.html"

    }
    ################################################
    #  Network Information
    ################################################
    
    if ($options[6] -eq 1) {

        ConvertTo-Html -Head $a -Body "<H1>Network Information</H1>" >> "$filepath\$currentPC.html"

        Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $currentPC |
            Select-Object Description, DHCPServer, 
        @{Name = 'IpAddress'; Expression = {$_.IpAddress -join '; '}}, 
        @{Name = 'IpSubnet'; Expression = {$_.IpSubnet -join '; '}}, 
        @{Name = 'DefaultIPgateway'; Expression = {$_.DefaultIPgateway -join '; '}}, 
        @{Name = 'DNSServerSearchOrder'; Expression = {$_.DNSServerSearchOrder -join '; '}}, 
        WinsPrimaryServer, WINSSecondaryServer| Where-Object {$_.IpAddress -ge 1 } `
            | ConvertTo-html -Body "<button class='collapsible'>IP Address </button>" >>  "$filepath\$currentPC.html" 

        Invoke-Command -cn $currentPC -scriptblock {get-smbconnection} | Select-Object Credential, Dialect, Encrypted, NumOpens, Redirected, ServerName, ShareName, Signed |
            ConvertTo-html -Body "<button class='collapsible'>SMB Connections </button>" | Set-CellColor Dialect red -Filter "Dialect -lt '3'" | Set-CellColor Encrypted red -Filter "Encrypted -eq 'False'" >> "$filepath\$currentPC.html" -EA SilentlyContinue # 
        
        get-WmiObject win32_networkadapter -ComputerName $currentPC | Select-Object AdapterType, Speed, MACAddress, NetConnectionID | Where-Object speed -ge 1 | Sort-Object speed -Descending `
            | ConvertTo-html -Body "<button class='collapsible'>Network Adapters</button>" >>  "$filepath\$currentPC.html" 

        Invoke-Command -cn $currentPC -scriptblock {Get-NetAdapterAdvancedProperty} | Select-Object Name, DisplayName, DisplayValue RegistryValue |
            ConvertTo-html -Body "<button class='collapsible'>Network Advanced Properties </button>" >>  "$filepath\$currentPC.html" 

        Invoke-Command -cn $currentPC -scriptblock {Get-NetAdapterPowerManagement} | Select-Object InterfaceDescription, Name, AllowComputerToTurnOffDevice, Select-ObjectiveSuspend, DeviceSleepOnDisconnect |
            ConvertTo-html -Body "<button class='collapsible'>Power Plan: Network Power Management Properties </button>" | Set-CellColor "AllowComputerToTurnOffDevice" red -Filter "AllowComputerToTurnOffDevice -eq '2'" `
            >> "$filepath\$currentPC.html" -EA SilentlyContinue 
        
    }
    ################################################
    #  Reliability Information
    #################################################

    if ($options[7] -eq 1) {

        ConvertTo-Html -Head $a -Body "<H2>Reliability Information</H2>" >> "$filepath\$currentPC.html"

        get-wmiobject Win32_ReliabilityStabilityMetrics -computername $currentPC -property @("__SERVER", "SystemStabilityIndex") | Select-Object -first 1 __SERVER, SystemStabilityIndex `
            | ConvertTo-html -Body "<button class='collapsible'> Reliability Index </button>" `
            >> "$filepath\$currentPC.html" -EA SilentlyContinue

        get-wmiobject -computername $currentPC Win32_ReliabilityRecords -property @("SourceName", "EventIdentifier") | Where-Object {!(($_.Sourcename -eq "Microsoft-Windows-WindowsUpdateClient" -or $_.Sourcename -eq "MsiInstaller" ))} | group-object -property SourceName, EventIdentifier -noelement | sort-object -descending Count | Select-Object Count, Name `
            | ConvertTo-html -Body "<button class='collapsible'>Reliability Distribution </button>" `
            >> "$filepath\$currentPC.html" -EA SilentlyContinue

        get-wmiobject Win32_ReliabilityRecords -computername $currentPC -property Message | Select-Object -Last 30 Message | Where-Object {!(($_.Message -like "*successfully installed*" -or $_.Message -like "*error status: 0." -or $_.Message -like "*Windows Installer requires a system restart*" ))}  `
            | ConvertTo-html -Body "<button class='collapsible'>Last 30 Reliability Messages </button>" `
            >> "$filepath\$currentPC.html" -EA SilentlyContinue 

        Get-WmiObject -computername $currentPC -Class Win32_ReliabilityRecords -Filter "ProductName = 'iexplore.exe'" | Measure-Object | Select-Object count `
            | ConvertTo-html -Body "<button class='collapsible'>Total IE Hangs </button>" `
            >> "$filepath\$currentPC.html" -EA SilentlyContinue

        Get-WmiObject -computername $currentPC -Class win32_reliabilityRecords -Filter "ProductName = 'iexplore.exe'" | Foreach-Object {[regex]::matches($_.message, "Faulting module path: [a-zA-Z:0-9()\\\s]*")} | Sort-Object -Property value | Group-Object -Property value | Sort-Object -property count -Descending `
            | ConvertTo-html -Body "<button class='collapsible'>Hanging IE Modules (If Applicable)</button>" `
            >> "$filepath\$currentPC.html" -EA SilentlyContinue

        Get-WmiObject -computername $currentPC -Class Win32_ReliabilityRecords -Filter "ProductName = 'Outlook.exe'" | Measure-Object | Select-Object count `
            | ConvertTo-html -Body "<button class='collapsible'>Total Outlook Hangs </button>" `
            >> "$filepath\$currentPC.html" -EA SilentlyContinue

        Get-WmiObject -computername $currentPC -Class win32_reliabilityRecords -Filter "ProductName = 'Outlook.exe'" | Foreach-Object {[regex]::matches($_.message, "Faulting module path: [a-zA-Z:0-9()\\\s]*")} | Sort-Object -Property value | Group-Object -Property value | Sort-Object -property count -Descending `
            | ConvertTo-html -Body "<button class='collapsible'>Hanging Outlook Modules (If Applicable)</button>" `
            >> "$filepath\$currentPC.html" -ErrorAction SilentlyContinue

    }
    ################################################
    #  Log Files
    #################################################

    write-host "Step 2: Gathering all Logs" -ForegroundColor Green

    ConvertTo-Html -Head $a -Body "<H2>Logs</H2>" >> "$filepath\$currentPC.html"

    $applog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Application"; StartTime = $date } | Where-Object {$_.ProviderName -eq "Outlook" -or $_.ProviderName -eq "EAPHost" -or $_.ProviderName -eq "ERCService" -or $_.ProviderName -eq "Login.vbs"}| Select-Object LogName, TimeCreated, ID, Message, ProviderName 
    $systemlog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "System"; StartTime = $date } | Where-Object {$_.ProviderName -eq "Microsoft-Windows-Power-Troubleshooter" -or $_.ProviderName -eq "Microsoft-Windows-Kernel-Power" -or $_.ProviderName -eq "e1dexpress"} | Select-Object LogName, TimeCreated, ID, Message, ProviderName 
    $AADLog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-AAD/Operational"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName 
    $DiagnosticsLog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-Diagnostics-Performance/Operational"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName 
    $EAPhostLog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-EAPHost/Operational"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName 
    $EAPMethodsLog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-EAPMethods-RASTls/Operational"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName 
    $GroupPolicyLog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-GroupPolicy/Operational"; StartTime = $date } | Where-Object {$_.ID -eq "5312" -or $_.ID -eq "4004" -or $_.ID -eq "8004"} | Select-Object LogName, TimeCreated, ID, Message, ProviderName 
    $NCSILog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-NCSI/Operational"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName 
    $NetworkProfileLog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-NetworkProfile/Operational"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName
    $SMBCLientLog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-SMBClient/Connectivity"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName  
    $WiredAutoConfigLog = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-Wired-AutoConfig/Operational"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName  
    $WLANAutoconfig = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-WLAN-AutoConfig/Operational"; StartTime = $date } | Select-Object LogName, TimeCreated, ID, Message, ProviderName  
    $NTLM = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Microsoft-Windows-NTLM/Operational"; StartTime = $date } | Where-Object {$_.LevelDisplayName -eq "Error" -or $_.LevelDisplayName -eq "Warning"} | Select-Object LogName, TimeCreated, ID, Message, ProviderName  

    $loginarray1 = $null
    $loginarray1 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Security"; StartTime = $date; Data = $ComputerOwner.USERNAME } | Select-Object TimeCreated, ID, Message, ProviderName 
    $loginarray2 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Security"; StartTime = $date; ID = 1100 } | Select-Object TimeCreated, ID, Message, ProviderName
    $loginarray3 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Application"; StartTime = $date; ID = 801 } | Select-Object TimeCreated, ID, Message, ProviderName
    $loginarray4 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "System"; StartTime = $date; ID = 12 } | Select-Object TimeCreated, ID, Message, ProviderName
    $loginarray5 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "System"; StartTime = $date; ID = 13 } | Select-Object TimeCreated, ID, Message, ProviderName

    $loginarray6 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Security"; StartTime = $date; ID = 4800 } | Select-Object TimeCreated, ID, Message, ProviderName

    $loginarray7 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Security"; StartTime = $date; ID = 4648 } | Select-Object TimeCreated, ID, Message, ProviderName

    $loginarray8 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Security"; StartTime = $date; ID = 4624 } | Select-Object TimeCreated, ID, Message, ProviderName

    $loginarray9 = Get-WinEvent -ComputerName $currentPC -FilterHashTable @{ LogName = "Security"; StartTime = $date; ID = 4634 } | Select-Object TimeCreated, ID, Message, ProviderName

    write-host "Step 3: Sorting all Logs"  -ForegroundColor Green

    [array]$eventlogs = $null
    $eventlogs += $applog 
    $eventlogs += $systemlog
    $eventlogs += $AADLog
    $eventlogs += $DiagnosticsLog
    $eventlogs += $EAPhostLog
    $eventlogs += $EAPMethodsLog
    $eventlogs += $GroupPolicyLog
    $eventlogs += $NCSILog
    $eventlogs += $NetworkProfileLog
    $eventlogs += $SMBCLientLog
    $eventlogs += $WiredAutoConfigLog
    $eventlogs += $WLANAutoconfig
    $eventlogs += $NTLM
 
    $eventlogs | Sort-Object TimeCreated -Descending | ConvertTo-html -Body "<button class='collapsible'>TimeLine For Extra Log Events (If Applicable)</button>" `
        | Set-CellColor Id red -Filter "Id -eq 30804 -or Id -eq 30805 -or Id -eq 109 -or Id -eq 42 -or Id -eq 27 -or Id -eq 200" | Set-CellColor Message red -Filter "Message -like '*Exchange has been lost*'" `
        | Set-CellColor Message red -Filter "Message -like '*shutdown*'" `| Set-CellColor Id green -Filter "Id -eq 30807 -or Id -eq 30805 -or Id -eq 30806 -or Id -eq 801 -or Id -eq 100" | Set-CellColor Message green -Filter "Message -like '*re-established*'" `
        | Set-CellColor Message red -Filter "Message -like '*entering sleep*'" | Set-CellColor Message red -Filter "Message -like '*Power source change*'" | Set-CellColor Message red -Filter "Message -like '*link is disconnected*'" `
        | Set-CellColor Message green -Filter "Message -like '*started up*'" | Set-CellColor Message green -Filter "Message -like '*resumed from sleep*'" | Set-CellColor Message green -Filter "Message -like '*has returned*'" `
        >> "$filepath\$currentPC.html" 

    [array]$Userlogineventlogs = $null
    $Userlogineventlogs += $loginarray1 
    $Userlogineventlogs += $loginarray2
    $Userlogineventlogs += $loginarray3
    $Userlogineventlogs += $loginarray4
    $Userlogineventlogs += $loginarray5

    $Userlogineventlogs | Sort-Object TimeCreated -Descending | Where-Object {!(($_.Message -like "*$vComputerName$*" -or $_.Message -like "*UpdateTrustedSites.exe*" -or $_.Message -like "*Microsoft.Uev.SyncController.exe*" -or $_.Message -like "*A new process has been created*"))} `
        | ConvertTo-html -Body "<button class='collapsible'>TimeLine For Login Events: Basic</button>" `
        | Set-CellColor Id red -Filter "Id -eq 30804 -or Id -eq 30805 -or Id -eq 109 -or Id -eq 42 -or Id -eq 27 -or Id -eq 200 -or ID -eq 1100" | Set-CellColor Message red -Filter "Message -like '*Exchange has been lost*'" `
        | Set-CellColor Message red -Filter "Message -like '*entering sleep*'" | Set-CellColor Message red -Filter "Message -like '*Power source change*'" | Set-CellColor Message red -Filter "Message -like '*link is disconnected*'" `
        | Set-CellColor Message red -Filter "Message -like '*shutdown*'" | Set-CellColor Id green -Filter "Id -eq 30807 -or Id -eq 30805 -or Id -eq 30806 -or Id -eq 801 -or Id -eq 100 -or Id -eq 4801" | Set-CellColor Message green -Filter "Message -like '*re-established*'" `
        | Set-CellColor Message green -Filter "Message -like '*started up*'" | Set-CellColor Message green -Filter "Message -like '*resumed from sleep*'" | Set-CellColor Message green -Filter "Message -like '*has returned*'" `
        >> "$filepath\$currentPC.html" -EA SilentlyContinue

    $loginarray7 | Sort-Object TimeCreated -Descending | Where-Object {!(($_.Message -like "*$vComputerName$*" -or $_.Message -like "*UpdateTrustedSites.exe*" -or $_.Message -like "*Microsoft.Uev.SyncController.exe*" -or $_.Message -like "*A new process has been created*"))} `
        | ConvertTo-html -Body "<button class='collapsible'>TimeLine For Login Events: ID = 4648 (Failed Login Attempts)</button>" `
        >> "$filepath\$currentPC.html" -EA SilentlyContinue

    $loginarray8 | Sort-Object TimeCreated -Descending | Where-Object {!(($_.Message -like "*$vComputerName$*" -or $_.Message -like "*UpdateTrustedSites.exe*" -or $_.Message -like "*Microsoft.Uev.SyncController.exe*" -or $_.Message -like "*A new process has been created*"))} `
        | ConvertTo-html -Body "<button class='collapsible'>TimeLine For Login Events: ID = 4624 (Accounts Logged In)</button>" `
        >> "$filepath\$currentPC.html" -EA SilentlyContinue

    $loginarray9 | Sort-Object TimeCreated -Descending | Where-Object {!(($_.Message -like "*$vComputerName$*" -or $_.Message -like "*UpdateTrustedSites.exe*" -or $_.Message -like "*Microsoft.Uev.SyncController.exe*" -or $_.Message -like "*A new process has been created*"))} `
        | ConvertTo-html -Body "<button class='collapsible'>TimeLine For Login Events: ID = 4634 (Accounts Logged Off)</button>" `
        >> "$filepath\$currentPC.html" -EA SilentlyContinue

    $loginarray6 | Sort-Object TimeCreated -Descending | Where-Object {!(($_.Message -like "*$vComputerName$*" -or $_.Message -like "*UpdateTrustedSites.exe*" -or $_.Message -like "*Microsoft.Uev.SyncController.exe*" -or $_.Message -like "*A new process has been created*"))} `
        | ConvertTo-html -Body "<button class='collapsible'>TimeLine For Login Events: ID = 4800 (Workstation Was Locked)</button>" `
        >> "$filepath\$currentPC.html" 
        
    
    ConvertTo-html -PostContent '<script>
    var coll = document.getElementsByClassName("collapsible");
    var i;
    
    for (i = 0; i < coll.length; i++) {
        coll[i].addEventListener("click", function() {
            this.classList.toggle("active");
            var content = this.nextElementSibling;
            if (content.style.display === "block") {
                content.style.display = "none";
            } else {
                content.style.display = "block";
            }
        });
    }

    var elems = document.getElementsByTagName("TABLE");

    expandBtn.onclick = function(){
        for (var i=0;i<elems.length;i+=1){
            elems[i].style.display = "block";
            }
        }

    collapseBtn.onclick = function(){
        for (var i=0;i<elems.length;i+=1){
            elems[i].style.display = "none";
            }
        }

    </script>' >> "$filepath\$currentPC.html" -ErrorAction SilentlyContinue

    $Report = "The Report is generated On  $(get-date) by $((Get-Item env:\username).Value) on computer $((Get-Item env:\Computername).Value)"
    $Report >> "$filepath\$currentPC.html" 

    ######################################################################################## 
    ########################### Create & Send Email (If needed) ############################     
    $reportpath = "$filepath\$currentPC.html" 

    
    IF ($options.GUI_Config.Mail_Config.Host_Email -ne '') {
        write-host "Sending Email" -ForegroundColor Green
        Send-MailMessage -From $options.GUI_Config.Mail_Config.Host_Email -To $to -Subject "Eventlog Timeline $currentPC " -Body "Logon events for $computeruser for the past $days day(s)" -Attachments $reportpath -Priority High -dno onSuccess, onFailure -SmtpServer $options.GUI_Config.Mail_Config.SMTP_Host  
        Remove-Item –path $reportpath
    }
}
# !!!!! BRACKET ABOVE CLOSES FOR LOOP !!!!!

#################### END of SCRIPT ####################################

# Get End Time
$endDTM = (Get-Date)
# Echo Time elapsed
write-host "Elapsed Time: $(($endDTM-$startDTM)).Minutes" -ForegroundColor Green