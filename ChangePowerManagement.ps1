#region Styles and Variables
$a = "<style>"
$a = $a + "BODY{background-color:LightGoldenRodYellow; font-family:Calibri;font-size:12pt;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;font-family:Calibri;font-size:10pt;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:CadetBlue}"
$a = $a + "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:tan}"
$a = $a + ".odd  { background-color:#ffffff; }"
$a = $a + ".even { background-color:#dddddd; }"
$a = $a + "</style>"


################################################################################################
#### Global variables ####
################################################################################################

$ErrorActionPreference = 'SilentlyContinue'
$vComputerName = Get-Content -Path .\PCList.txt

#endregion

For ($i = 0; $i -le ($vComputerName.Length - 2); $i++) {
    
    $currentPC = $vComputerName[$i] #Using vComputerName[$i] will break the report file

    ################################################################################################
    #### Functions ####
    ################################################################################################

    function Show-Menu {
        Clear-Host
        Write-Host "========================= Menu ========================="
     
        Write-Host "1: Press '1' to change the power management settings to 'High Performance'."
        Write-Host "2: Press '2' to change the power management settings to 'Mobile'."
        Write-Host "3: Press '3' to change the power management settings to 'Balanced'."
        Write-Host "Q: Press 'Q' to quit."
    }

    #endregion

    #region Get Inputs
  
    if ($currentPC -eq '') {
        (Write-Host No Computer Name Retrieved!  -ForegroundColor Yellow -BackgroundColor Red )
        Break Script
    }

    If (-not (Test-Connection -Computername $currentPC -BufferSize 16 -Count 1 -Quiet)) {  
        (Write-Host $currentPC is Not online!  -ForegroundColor Yellow -BackgroundColor Red )
        Break Script
    }

    Else {
        (Write-Host $currentPC is responding.  -ForegroundColor Black -BackgroundColor Green ) 
    }
    #endregion

    Show-Menu
    $input = Read-Host "Please make a selection"
    switch ($input) {
        '1' {
            Clear-Host
            'You chose option #1'
            Write-Host Changing PowerManagement Power Plan to High Performance -ForegroundColor Green -BackgroundColor Black         
            Invoke-Command -cn $currentPC -scriptblock {
                $powerPlan = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan -Filter "ElementName = 'High Performance'";
                $powerPlan.Activate()
            }

        } '2' {
            Clear-Host
            'You chose option #2'
            Write-Host Changing PowerManagement Power Plan to High Performance -ForegroundColor Green -BackgroundColor Black          
            Invoke-Command -cn $currentPC -scriptblock {
                $powerPlan = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan -Filter "ElementName = 'Mobile'";
                $powerPlan.Activate()
            }
        } '3' {
            Clear-Host
            'You chose option #3'
            Write-Host Changing PowerManagement Power Plan to High Performance -ForegroundColor Green -BackgroundColor Black        
            Invoke-Command -cn $currentPC -scriptblock {   
                $powerPlan = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan -Filter "ElementName = 'Balanced'";
                $powerPlan.Activate()}
        } 'q' {
            Clear-Host
            return
        }
    }
    pause
}
    