param([int]$Cores)

$vComputerName = Get-Content -Path .\PCList.txt

#endregion
function stress {
    param([int]$Cores, [string]$currentPC)
    #$Cores = 2
    
    #   Variables

    $filepath = "C:\Support"
    $PC2 = 45.3, 25.94, 25
    $PC3 = 45.3, 36.2, 36
    $PC4 = 45.3, 45.37, 45

    $psInstances = $Cores
    if ($psInstances -gt 1) {
        $psName = "{0}" #{1}" -f $psProcess.name, $($psInstances – 1)
    }      
    else {
        $psName = $psProcess.name
    }

    # RAM in box
    $box = get-WMIobject Win32_ComputerSystem
    $Global:physMB = $box.TotalPhysicalMemory / 1024 / 1024

    #####################
    # perfmon counters
    $Global:psPerfCPU = new-object System.Diagnostics.PerformanceCounter('Processor', '% Processor Time', '_Total')
    $Global:psPerfMEM = new-object System.Diagnostics.PerformanceCounter('Memory', 'Available Mbytes')
    $psPerfCPU.NextValue() | Out-Null
    $psPerfMEM.NextValue() | Out-Null

    # timer 
    $Global:psTimer = New-Object System.Timers.Timer
    $psTimer.Interval = 1500

    # every timer interval, update the CMD shell window with RAM and CPU info.

    Write-Warning "This script will potentially saturate CPU & RAM utilization!"
    #$Prompt = Read-Host "Are you sure you want to proceed? (Y/N)"

    #if ($Prompt -eq 'Y') {

        $Log = "C:\CPUStressTest.ps1.log"
        $StartDate = Get-Date
        Write-Output "============= CPU & RAM Stress Test Started: $StartDate =============" >> $Log
        Write-Output "Started By: $env:username" >> $Log

        Write-Warning "Beginning RAM test"
        $RAMStart = Get-Date -UFormat %s
        #RAM test
        $a = "a" * 256MB
        $growArray = @()
        $growArray += $a
        # leave 512Mb for the OS to survive.
        $HEADROOM = 300
        $bigArray = @()
        $ram = $physMB - $psPerfMEM.NextValue()
        $MAXRAM = $physMB - $HEADROOM
        $k = 0
        while ($ram -lt $MAXRAM) {
            $bigArray += , @($k, $growArray)
            $k += 1
            $growArray += $a
            $ram = $physMB - $psPerfMEM.NextValue()
        }
        # and now release it all.
        Write-Warning "Ending RAM test"
        $bigArray.clear()
        remove-variable bigArray
        $growArray.clear()
        remove-variable growArray
        [System.GC]::Collect()
        $Temptime = Get-Date -UFormat %s
        $TotalTime = $Temptime- $RAMStart 
        #Write-Warning "Ram time: $TotalTime" #| Out-File -append "$filepath\StressTest.txt"
        $TotalTime | Out-File -append "$filepath\Stress.txt"

        #######################################################################################
        ### Test web traffic load
        #######################################################################################

        <# Write-Warning "Beginning Traffic test"

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $url = 'http://giphy.com/giphytv'
        while ($true) {
        try {
        [net.httpWebRequest]
        $req = [net.webRequest]::create($url)
        $req.method = "GET"
        $req.ContentType = "application/x-www-form-urlencoded"
        $req.TimeOut = 60000 

        $start = get-date
        [net.httpWebResponse] $res = $req.getResponse()
        $timetaken = ((get-date) - $start).TotalMilliseconds 

        Write-Output $res.Content
        Write-Output ("{0} {1} {2}" -f (get-date), $res.StatusCode.value__, $timetaken)
        $req = $null
        $res.Close()
        $res = $null
        }
        catch [Exception] {
            Write-Output ("{0} {1}" -f (get-date), $_.ToString())
        }
        $req = $null 

         #uncomment the line below and change the wait time to add a pause between requests
         #Start-Start-Sleep -Seconds 1
        }

        Write-Warning "Ending Traffic test" #>

        $timeTaken = Measure-Command -Expression {
        }
        $milliseconds = $timeTaken.TotalMilliseconds
        $milliseconds = [Math]::Round($milliseconds, 0)

        #######################################################################################
        ### Launch Common IE Windows
        #######################################################################################

        # Invoke Internet Explorer with some common Firm URLs
        # Find other things to do with this Get-Member -InputObject $IE
        write-host "Launching IE With Common Apps for 30ish seconds"

        $IE1 = New-Object -ComObject InternetExplorer.Application
        $IE1.visible = $true; 
        $IE2 = New-Object -ComObject InternetExplorer.Application
        $IE2.visible = $true; 
        $IE3 = New-Object -ComObject InternetExplorer.Application
        $IE3.visible = $true; 
        $IE4 = New-Object -ComObject InternetExplorer.Application
        $IE4.visible = $true; 

        #######################################################################################
        ### Maximize the Screen
        #######################################################################################

        $sw = @'
[DllImport("user32.dll")]
public static extern int ShowWindow(int hwnd, int nCmdShow);
'@
        $type = Add-Type -Name ShowWindow2 -MemberDefinition $sw -Language CSharpVersion3 -Namespace Utils -PassThru
        $type::ShowWindow($ie3.hwnd, 3) # 3 = maximize
        do {Start-Sleep 1} until (-not ($IE1.Busy)) 
        $IE1.navigate("http://one.nmrs.com/Pages/Firm%20View.aspx")
        do {Start-Sleep 1} until (-not ($IE1.Busy)) 
        $IE2.navigate("https://app.chromeriver.com/")
        do {Start-Sleep 1} until (-not ($IE1.Busy)) 
        $IE3.navigate("https://www.dayforcehcm.com/mydayforce/login.aspx")
        do {Start-Sleep 1} until (-not ($IE1.Busy)) 
        $IE4.navigate("https://vault.netvoyage.com/neWeb2/docCent.aspx")
        do {Start-Sleep 1} until (-not ($IE1.Busy)) 


        $IE = New-Object -ComObject InternetExplorer.Application
        $IE.navigate("www.yahoo.com");
        $IE.visible = $true; 

        #######################################################################################
        ### Maximize the Screen
        #######################################################################################

        $sw = @'
        [DllImport("user32.dll")]
        public static extern int ShowWindow(int hwnd, int nCmdShow);
'@
        $type = Add-Type -Name ShowWindow3 -MemberDefinition $sw -Language CSharpVersion3 -Namespace Utils -PassThru
        $type::ShowWindow($ie.hwnd, 3) # 3 = maximize
        do {Start-Sleep 1} until (-not ($IE.Busy)) 

        #######################################################################################

        <# $IE.Document.parentWindow.scrollTo(0, 10000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 20000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 30000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 0)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 40000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 50000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 60000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 70000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 80000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 90000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 0)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 100000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 110000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 120000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 130000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 140000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 0)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 150000)
        do {Start-Sleep 1} until (-not ($IE.Busy))
        $IE.Document.parentWindow.scrollTo(0, 160000) #>

        ##########################################
    
        S tart-Sleep 30
        $IE.Quit()
        $IE1.Quit()
        $IE2.Quit()
        $IE3.Quit()
        $IE4.Quit()

        

        $CPUStart = Get-Date -UFormat %s
        
        Write-Warning "Stressing CPU's"# for 45s"
        #Write-Output "Hyper Core Count: $Cores" >> $Log
        foreach ($loopnumber in 1..$Cores) {
            Start-Job -ScriptBlock {
                $result = 1
                foreach ($number in 1..11474836) {
                    $result = $result * $number
                }# end foreach
            }# end Start-Job
        }# end foreach
        #Start-Sleep 45
        if ($Cores -eq 0,2) {
            Wait-Job -id 1,3
        } elseif ($Cores -eq 3) {
            Wait-Job -id 1,3,5
        } elseif ($Cores -eq 4) {
            Wait-Job -id 1,3,5,7 
        } 
        
        Get-Job | Stop-Job
        
        $Temptime = Get-Date -UFormat %s
        $TotalTime = $Temptime - $CPUStart 
        #Write-Warning "CPU Compute time: $TotalTime" #| Out-File -append "$filepath\Stress.txt"
        $TotalTime | Out-File -append "$filepath\Stress.txt"

        $Current_Times = Get-Content -Path $filepath\Stress.txt
        $Current_Times += $Current_Times[0]
        $Current_Times += $Current_Times[1]

        [array]$Results = $null

        
        if ($Cores -eq 2) {
            For ($i = 0; $i -le ($PC2.Length - 1); $i++) {
                $Results += ($Current_Times[$i]/$PC2[$i]).tostring("P")
                $RawResults += ($Current_Times[$i]/$PC2[$i])
            }
        } elseif ($Cores -eq 3) {
            For ($i = 0; $i -le ($PC3.Length - 1); $i++) {
                $Results += ($Current_Times[$i]/$PC3[$i]).tostring("P")
                $RawResults += ($Current_Times[$i]/$PC3[$i]) 
            }
        }elseif ($Cores -eq 4) {
            For ($i = 0; $i -le ($PC4.Length - 1); $i++) {
                $Results += ($Current_Times[$i]/$PC4[$i]).tostring("P")
                $RawResults += ($Current_Times[$i]/$PC4[$i])
            }
        } 

        Write-Warning $Results[0]

        $Ratio0 = $Results[0]
        $Ratio1 = $Results[1]
        $Ratio2 = $Results[2]
        

        Write-Warning "$currentPC is $Ratio0 better/worse than the average RAM Test against 8gb RAM" >> "$filepath\Stress_Results.txt"
        Write-Warning "$currentPC is $Ratio1 better/worse than the average CPU Test against an X1 Carbon Yoga Gen 2 w/ i5 7200U @ 2.5GHz" >> "$filepath\Stress_Results.txt"
        Write-Warning "$currentPC is $Ratio2 better/worse than the average CPU Test against an M900 w/ ?" >> "$filepath\Stress_Results.txt"
        Write-Warning ""
        

        if ($RawResults[0] -lt .80) {
            Write-Warning "Compared to other PC's, it seems like either this PC has less than 8gb or RAM, or the computer is facing memory issues"
        } else {
            Write-Warning "RAM is performing as expected"
        }

        if ($RawResults[1] -lt .80) {
            Write-Warning "Compared to other laptops, it seems like the CPU is not performing optimally"
        } else {
            Write-Warning "Compared to other laptops, the CPU is performing as expected"
        }

        if ($RawResults[2] -lt .80) {
            Write-Warning "Compared to other desktops, it seems like the CPU is not performing optimally"
        } else {
            Write-Warning "Compared to other desktops, the CPU is performing as expected"
        }

        if ($milliseconds -gt 7) {
            Write-Warning "Compared to other PC's, it seems like the Network may not performing optimally with a response time of $milliseconds"
        } else {
            Write-Warning "Compared to other PC's, the Network seems to be performing as expected"
        }


    #}
    #else {
    #    Write-Output "Job Cancelled!"
    #}

    #$EndDate = Get-Date
    #Write-Warning "============= Stress Test Complete: $EndDate =============" >> $Log
}


For ($i = 0; $i -le ($vComputerName.Length - 2); $i++) {
    
    $currentPC = $vComputerName[$i] #Using vComputerName[$i] will break the report file
    Write-Warning "Currently testing: $currentPC"
    Write-Warning "Number of cores selected: $Cores"
    Invoke-Command -ComputerName $currentPC -ScriptBlock ${function:stress} -ArgumentList $Cores, $currentPC
}


