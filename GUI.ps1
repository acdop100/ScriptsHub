#Requires -RunAsAdministrator

$options = [xml](Get-Content ".\config.xml")

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

IF ($options.GUI_Config.Mail_Config.Host_Email -eq ""){
    $FirstRun = $true
}

$Form                                = New-Object system.Windows.Forms.Form
$Form.ClientSize                     = '458,390'
$Form.text                           = "ScriptsHub"
#$form.Topmost                        = $true
$form.StartPosition                  = "CenterScreen"
$Form.BackColor                      = "WhiteSmoke"

$Icon                                = New-Object system.drawing.icon ($PSScriptRoot + "\nm.ICO")
$Form.Icon                           = $Icon

$MainLabel                           = New-Object system.Windows.Forms.Label
$MainLabel.text                      = "ScriptsHub"
$MainLabel.AutoSize                  = $true
$MainLabel.width                     = 5
$MainLabel.height                    = 10
$MainLabel.location                  = New-Object System.Drawing.Point(175,18)
$MainLabel.Font                      = 'Tahoma,14,style=Bold'
    
$FirstRunLabel1                      = New-Object system.Windows.Forms.Label
$FirstRunLabel1.text                = "It looks like no SMTP or report emailing account has been setup"
$FirstRunLabel1.width               = 440
$FirstRunLabel1.height              = 20
$FirstRunLabel1.location            = New-Object System.Drawing.Point(34,64)
$FirstRunLabel1.Font                = 'Tahoma,10'
$FirstRunLabel1.Visible             = $FirstRun

$FirstRunLabel2                     = New-Object system.Windows.Forms.Label
$FirstRunLabel2.text                = "Enter your SMTP address and the sender email address"
$FirstRunLabel2.width               = 440
$FirstRunLabel2.height              = 20
$FirstRunLabel2.location            = New-Object System.Drawing.Point(67,84)
$FirstRunLabel2.Font                = 'Tahoma,10'
$FirstRunLabel2.Visible             = $FirstRun

$FirstRunInput1                     = New-Object system.Windows.Forms.Textbox
$FirstRunInput1.width               = 350
$FirstRunInput1.height              = 40
$FirstRunInput1.text                = $PSEmailServer
$FirstRunInput1.location            = New-Object System.Drawing.Point(55,135)
$FirstRunInput1.Font                = 'Tahoma,10'
$FirstRunInput1.Visible             = $FirstRun

$FirstRunInput2                     = New-Object system.Windows.Forms.Textbox
$FirstRunInput2.width               = 350
$FirstRunInput2.height              = 40
$FirstRunInput2.text                = ""
$FirstRunInput2.location            = New-Object System.Drawing.Point(55,170)
$FirstRunInput2.Font                = 'Tahoma,10'
$FirstRunInput2.Visible             = $FirstRun

$FirstRunBtn1                       = New-Object system.Windows.Forms.Button
$FirstRunBtn1.text                  = "Continue"
$FirstRunBtn1.width                 = 350
$FirstRunBtn1.height                = 30
$FirstRunBtn1.location              = New-Object System.Drawing.Point(55,250)
$FirstRunBtn1.Font                  = 'Tahoma,10'
$FirstRunBtn1.Visible               = $FirstRun

# Main Labels
#region begin GUI{ 

$PCNamesLabel                        = New-Object system.Windows.Forms.Label
$PCNamesLabel.text                   = "Enter the pc you want to run these on or choose a .txt file of names"
$PCNamesLabel.width                  = 440
$PCNamesLabel.height                 = 20
$PCNamesLabel.location               = New-Object System.Drawing.Point(26,64)
$PCNamesLabel.Font                   = 'Tahoma,10'
$PCNamesLabel.Visible                = !$FirstRun

$WarningLabel                        = New-Object system.Windows.Forms.Label
$WarningLabel.text                   = "Make sure you are on the same network as the PC(s)!"
$WarningLabel.width                  = 440
$WarningLabel.height                 = 20
$WarningLabel.location               = New-Object System.Drawing.Point(46,360)
$WarningLabel.Font                   = 'Tahoma,10style=Bold'
$WarningLabel.Visible                = !$FirstRun

$CommaLabel                          = New-Object system.Windows.Forms.Label
$CommaLabel.text                     = "Put a comma and space after each name"
$CommaLabel.AutoSize                 = $true
$CommaLabel.width                    = 440
$CommaLabel.height                   = 20
$CommaLabel.location                 = New-Object System.Drawing.Point(111,85)
$CommaLabel.Font                     = 'Tahoma,10,style=Underline'
$CommaLabel.Visible                  = !$FirstRun

$Seperator2                          = New-Object system.Windows.Forms.Label
$Seperator2.AutoSize                 = $false;
$Seperator2.Height                   = 1;
$Seperator2.width                    = 350
$Seperator2.location                 = New-Object System.Drawing.Point(55,260)
$Seperator2.BackColor                = "gray"
$Seperator2.Visible                  = !$FirstRun

$DayLabel                            = New-Object system.Windows.Forms.Label
$DayLabel.text                       = "Enter how many days back you want to search"
$DayLabel.AutoSize                   = $truen
$DayLabel.width                      = 100
$DayLabel.height                     = 20
$DayLabel.location                   = New-Object System.Drawing.Point(38,315)
$DayLabel.Font                       = 'Tahoma,10'
$DayLabel.Visible                    = $false

$EmailLabel                          = New-Object system.Windows.Forms.Label
$EmailLabel.text                     = "Enter the email addresses to send report to"
$EmailLabel.AutoSize                 = $true
$EmailLabel.width                    = 440
$EmailLabel.height                   = 20
$EmailLabel.location                 = New-Object System.Drawing.Point(105,195)
$EmailLabel.Font                     = 'Tahoma,10'
$EmailLabel.Visible                  = !$FirstRun

$CoresLabel                          = New-Object system.Windows.Forms.Label
$CoresLabel.text                     = "Enter number of cores/hyperthreads to stress before starting stress test!"
$CoresLabel.AutoSize                 = $true
$CoresLabel.width                    = 240
$CoresLabel.height                   = 30
$CoresLabel.location                 = New-Object System.Drawing.Point(11,205)
$CoresLabel.Font                     = 'Tahoma,10'
$CoresLabel.Visible                  = $false

$PCWarning                           = New-Object system.Windows.Forms.Label
$PCWarning.text                      = "These will use the PC name(s) you inputed in the previous screen"
$PCWarning.width                     = 440
$PCWarning.height                    = 20
$PCWarning.location                  = New-Object System.Drawing.Point(38,60)
$PCWarning.Font                      = 'Tahoma,10'
$PCWarning.Visible                   = $false

$BaseBenches                         = New-Object system.Windows.Forms.Label
$BaseBenches.text                    = "Baseline Benchmarks:"
$BaseBenches.width                   = 440
$BaseBenches.height                  = 20
$BaseBenches.location                = New-Object System.Drawing.Point(155,265)
$BaseBenches.Font                    = 'Tahoma,10style=Bold'
$BaseBenches.Visible                 = $false

$BaseBenches2                        = New-Object system.Windows.Forms.Label
$BaseBenches2.text                   = "Lenovo X1 Carbon Yoga Gen 2 (4 Cores): 45.37s"
$BaseBenches2.width                  = 440
$BaseBenches2.height                 = 20
$BaseBenches2.location               = New-Object System.Drawing.Point(85,290)
$BaseBenches2.Font                   = 'Tahoma,10'
$BaseBenches2.Visible                = $false

$BaseBenches3                        = New-Object system.Windows.Forms.Label
$BaseBenches3.text                   = "RAM Performance Average (8gb): 45.3s"
$BaseBenches3.width                  = 440
$BaseBenches3.height                 = 20
$BaseBenches3.location               = New-Object System.Drawing.Point(110,310)
$BaseBenches3.Font                   = 'Tahoma,10'
$BaseBenches3.Visible                = $false

$BaseBenches4                        = New-Object system.Windows.Forms.Label
$BaseBenches4.text                   = "Network Performance Average: 7ms"
$BaseBenches4.width                  = 440
$BaseBenches4.height                 = 20
$BaseBenches4.location               = New-Object System.Drawing.Point(120,330)
$BaseBenches4.Font                   = 'Tahoma,10'
$BaseBenches4.Visible                = $false

$NetTestLbl                          = New-Object system.Windows.Forms.Label
$NetTestLbl.text                     = "Enter the network path to test against (must be UNC format)"
$NetTestLbl.width                    = 440
$NetTestLbl.height                   = 20
$NetTestLbl.location                 = New-Object System.Drawing.Point(50,70)
$NetTestLbl.Font                     = 'Tahoma,10'
$NetTestLbl.Visible                  = $false

$NetInfoLbl                          = New-Object system.Windows.Forms.Label
$NetInfoLbl.text                     = "A good place is \\caeisi01\public\Hold\*NAME*"
$NetInfoLbl.width                    = 440
$NetInfoLbl.height                   = 20
$NetInfoLbl.location                 = New-Object System.Drawing.Point(95,170)
$NetInfoLbl.Font                     = 'Tahoma,10'
$NetInfoLbl.Visible                  = $false
#endregion GUI }

#Checkboxes
#region begin GUI{ 

$BiosCheckBox                        = New-Object system.Windows.Forms.CheckBox
$BiosCheckBox.text                   = "Get Image Version and Bios"
$BiosCheckBox.width                  = 200
$BiosCheckBox.height                 = 40
$BiosCheckBox.location               = New-Object System.Drawing.Point(38,70)
$BiosCheckBox.Font                   = 'Tahoma,10'
$BiosCheckBox.Visible                = $false
IF ($options.GUI_Config.MegaScript_Config.o1 -eq 1){$BiosCheckBox.checked = $true}

$RebootCheckBox                      = New-Object system.Windows.Forms.CheckBox
$RebootCheckBox.text                 = "Get last reboot"
$RebootCheckBox.width                = 200
$RebootCheckBox.height               = 40
$RebootCheckBox.location             = New-Object System.Drawing.Point(38,135)
$RebootCheckBox.Font                 = 'Tahoma,10'
$RebootCheckBox.Visible              = $false
IF ($options.GUI_Config.MegaScript_Config.o2 -eq 1){$RebootCheckBox.checked = $true}

$InfoCheckBox                        = New-Object system.Windows.Forms.CheckBox
$InfoCheckBox.text                   = "Get Free Memory, Processor, and Disk Info"
$InfoCheckBox.width                  = 180
$InfoCheckBox.height                 = 40
$InfoCheckBox.location               = New-Object System.Drawing.Point(38,200)
$InfoCheckBox.Font                   = 'Tahoma,10'
$InfoCheckBox.Visible                = $false
IF ($options.GUI_Config.MegaScript_Config.o3 -eq 1){$InfoCheckBox.checked = $true}

$HTTPCheckBox                        = New-Object system.Windows.Forms.CheckBox
$HTTPCheckBox.text                   = "Display HTTP Response Time"
$HTTPCheckBox.width                  = 220
$HTTPCheckBox.height                 = 40
$HTTPCheckBox.location               = New-Object System.Drawing.Point(38,265)
$HTTPCheckBox.Font                   = 'Tahoma,10'
$HTTPCheckBox.Visible                = $false
IF ($options.GUI_Config.MegaScript_Config.o4 -eq 1){$HTTPCheckBox.checked = $true}

$PowerPlanCheckBox                   = New-Object system.Windows.Forms.CheckBox
$PowerPlanCheckBox.text              = "Get Power Plan"
$PowerPlanCheckBox.width             = 200
$PowerPlanCheckBox.height            = 40
$PowerPlanCheckBox.location          = New-Object System.Drawing.Point(268,70)
$PowerPlanCheckBox.Font              = 'Tahoma,10'
$PowerPlanCheckBox.Visible           = $false
IF ($options.GUI_Config.MegaScript_Config.o5 -eq 1){$PowerPlanCheckBox.checked = $true}

$UpdatesCheckBox                     = New-Object system.Windows.Forms.CheckBox
$UpdatesCheckBox.text                = "See Updates and Hotfixes"
$UpdatesCheckBox.width               = 200
$UpdatesCheckBox.height              = 40
$UpdatesCheckBox.location            = New-Object System.Drawing.Point(268,135)
$UpdatesCheckBox.Font                = 'Tahoma,10'
$UpdatesCheckBox.Visible             = $false
IF ($options.GUI_Config.MegaScript_Config.o6 -eq 1){$UpdatesCheckBox.checked = $true}

$NetworkCheckBox                     = New-Object system.Windows.Forms.CheckBox
$NetworkCheckBox.text                = "Get Network Information"
$NetworkCheckBox.width               = 200
$NetworkCheckBox.height              = 40
$NetworkCheckBox.location            = New-Object System.Drawing.Point(268,200)
$NetworkCheckBox.Font                = 'Tahoma,10'
$NetworkCheckBox.Visible             = $false
IF ($options.GUI_Config.MegaScript_Config.o7 -eq 1){$NetworkCheckBox.checked = $true}

$ReliabilityCheckBox                 = New-Object system.Windows.Forms.CheckBox
$ReliabilityCheckBox.text            = "Reliability Information"
$ReliabilityCheckBox.width           = 200
$ReliabilityCheckBox.height          = 40
$ReliabilityCheckBox.location        = New-Object System.Drawing.Point(268,265)
$ReliabilityCheckBox.Font            = 'Tahoma,10'
$ReliabilityCheckBox.Visible         = $false
IF ($options.GUI_Config.MegaScript_Config.o8 -eq 1){$ReliabilityCheckBox.checked = $true}

#endregion GUI }

#Buttons
#region begin GUI {

$NamesListBtn                        = New-Object system.Windows.Forms.Button
$NamesListBtn.text                   = "Select list of names"
$NamesListBtn.width                  = 350
$NamesListBtn.height                 = 30
$NamesListBtn.location               = New-Object System.Drawing.Point(55,150)
$NamesListBtn.Font                   = 'Tahoma,10'
$NamesListBtn.Visible                = !$FirstRun
    
$ContinueBtn                         = New-Object system.Windows.Forms.Button
$ContinueBtn.text                    = "Continue"
$ContinueBtn.width                   = 350
$ContinueBtn.height                  = 30
$ContinueBtn.location                = New-Object System.Drawing.Point(55,275)
$ContinueBtn.Font                    = 'Tahoma,10'
$ContinueBtn.Visible                 = !$FirstRun

$OtherOptionsBtn                     = New-Object system.Windows.Forms.Button
$OtherOptionsBtn.text                = "Other Options"
$OtherOptionsBtn.width               = 350
$OtherOptionsBtn.height              = 30
$OtherOptionsBtn.location            = New-Object System.Drawing.Point(55,315)
$OtherOptionsBtn.Font                = 'Tahoma,10'
$OtherOptionsBtn.Visible             = !$FirstRun

$StartBtn                            = New-Object system.Windows.Forms.Button
$StartBtn.text                       = "Start"
$StartBtn.width                      = 90
$StartBtn.height                     = 40
$StartBtn.location                   = New-Object System.Drawing.Point(345,325)
$StartBtn.Font                       = 'Tahoma,10'
$StartBtn.Visible                    = $false

$BackBtn                             = New-Object system.Windows.Forms.Button
$BackBtn.text                        = "Back"
$BackBtn.width                       = 75
$BackBtn.height                      = 40
$BackBtn.location                    = New-Object System.Drawing.Point(10,10)
$BackBtn.Font                        = 'Tahoma,10'
$BackBtn.Visible                     = $false

$NetTestBtn                          = New-Object system.Windows.Forms.Button
$NetTestBtn.text                     = "Network Test"
$NetTestBtn.width                    = 350
$NetTestBtn.height                   = 30
$NetTestBtn.location                 = New-Object System.Drawing.Point(55,85)
$NetTestBtn.Font                     = 'Tahoma,10'
$NetTestBtn.Visible                  = $false

$PowerManagementBtn                  = New-Object system.Windows.Forms.Button
$PowerManagementBtn.text             = "Change Power Management"
$PowerManagementBtn.width            = 350
$PowerManagementBtn.height           = 30
$PowerManagementBtn.location         = New-Object System.Drawing.Point(55,125)
$PowerManagementBtn.Font             = 'Tahoma,10'
$PowerManagementBtn.Visible          = $false

$StressTestBtn                       = New-Object system.Windows.Forms.Button
$StressTestBtn.text                  = "Run Stress Test"
$StressTestBtn.width                 = 350
$StressTestBtn.height                = 30
$StressTestBtn.location              = New-Object System.Drawing.Point(55,165)
$StressTestBtn.Font                  = 'Tahoma,10'
$StressTestBtn.Visible               = $false

$NetTestStartBtn                       = New-Object system.Windows.Forms.Button
$NetTestStartBtn.text                  = "Run Network Speed Test"
$NetTestStartBtn.width                 = 350
$NetTestStartBtn.height                = 30
$NetTestStartBtn.location              = New-Object System.Drawing.Point(55,135)
$NetTestStartBtn.Font                  = 'Tahoma,10'
$NetTestStartBtn.Visible               = $false

#endregion GUI }

#Textboxes
#region begin GUI{ 

$PCNamesTextbox                      = New-Object system.Windows.Forms.Textbox
$PCNamesTextbox.width                = 350
$PCNamesTextbox.height               = 40
$PCNamesTextbox.text                 = ""
$PCNamesTextbox.location             = New-Object System.Drawing.Point(55,117)
$PCNamesTextbox.Font                 = 'Tahoma,10'
$PCNamesTextbox.Visible              = !$FirstRun

$DaysTextbox                         = New-Object system.Windows.Forms.Textbox
$DaysTextbox.width                   = 270
$DaysTextbox.height                  = 40
$DaysTextbox.text                    = ""
$DaysTextbox.location                = New-Object System.Drawing.Point(40,335)
$DaysTextbox.Font                    = 'Tahoma,10'
$DaysTextbox.Visible                 = $false

$EmailsTextbox                       = New-Object system.Windows.Forms.Textbox
$EmailsTextbox.width                 = 350
$EmailsTextbox.height                = 40
$EmailsTextbox.text                  = ""
$EmailsTextbox.location              = New-Object System.Drawing.Point(55,220)
$EmailsTextbox.Font                  = 'Tahoma,10'
$EmailsTextbox.Visible               = !$FirstRun

$CoresTextbox                        = New-Object system.Windows.Forms.Textbox
$CoresTextbox.width                  = 350
$CoresTextbox.height                 = 40
$CoresTextbox.text                   = ""
$CoresTextbox.location               = New-Object System.Drawing.Point(55,230)
$CoresTextbox.Font                   = 'Tahoma,10'
$CoresTextbox.Visible                = $false

$NetTextbox                        = New-Object system.Windows.Forms.Textbox
$NetTextbox.width                  = 350
$NetTextbox.height                 = 40
$NetTextbox.text                   = ""
$NetTextbox.location               = New-Object System.Drawing.Point(55,100)
$NetTextbox.Font                   = 'Tahoma,10'
$NetTextbox.Visible                = $false

#endregion GUI }

$Form.controls.AddRange(@($ComboBox1,$PCNamesLabel,$CommaLabel,$PCWarning,$CoresLabel,$NetTestBtn,$PowerManagementBtn,$StressTestBtn,$OtherOptionsBtn,$BackBtn,`
$StartBtn,$PCNamesTextbox,$EmailsTextbox,$CoresTextbox,$DaysTextbox,$BiosCheckBox,$RebootCheckBox,$InfoCheckBox,$HTTPCheckBox,$MainLabel,$PowerPlanCheckBox,`
$UpdatesCheckBox,$NetworkCheckBox,$ReliabilityCheckBox,$NamesListBtn,$ContinueBtn,$DayLabel,$EmailLabel,$BaseBenches2, $BaseBenches,$BaseBenches3, $BaseBenches4,  `
$Seperator2, $WarningLabel, $NetTestStartBtn, $NetTestLbl, $NetTextbox, $NetInfoLbl, $FirstRunInput1, $FirstRunInput2, $FirstRunLabel1, $FirstRunLabel2, $FirstRunBtn1))

$FirstRunBtn1.Add_Click({
    $options.GUI_Config.Mail_Config.SMTP_Host = [string]$FirstRunInput1.text
    $options.GUI_Config.Mail_Config.Host_Email = [string]$FirstRunInput2.text

    $options.save(".\config.xml")

    $wshell = New-Object -ComObject Wscript.Shell
    $response = $wshell.Popup("Make this SMTP default for all PowerShell?",0,"Saved!",32+4)

    IF ($response -eq 6) {
        $PSEmailServer = [string]$FirstRunInput1.text
    }

    $Seperator2.Visible              = $true
    $MainLabel.Visible               = $true
    $PCNamesLabel.Visible            = $true
    $CommaLabel.Visible              = $true
    $EmailLabel.Visible              = $true
    $WarningLabel.Visible            = $true
    $PCNamesTextbox.Visible          = $true
    $EmailsTextbox.Visible           = $true
    $NamesListBtn.Visible            = $true
    $ContinueBtn.Visible             = $true
    $OtherOptionsBtn.Visible         = $true
    $FirstRunInput1.Visible          = $false
    $FirstRunInput2.Visible          = $false
    $FirstRunLabel1.Visible          = $false
    $FirstRunLabel2.Visible          = $false
    $FirstRunBtn1.Visible            = $false
})

$StartBtn.Add_Click({
    IF ($BiosCheckBox.checked -eq $true){
        $options.GUI_Config.MegaScript_Config.o1 = "1"
    } else {
        $options.GUI_Config.MegaScript_Config.o1 = "0"
    }
    IF ($RebootCheckBox.checked -eq $true){
        $options.GUI_Config.MegaScript_Config.o2 = "1"
        
    } else {
        $options.GUI_Config.MegaScript_Config.o2 = "0" 
    }
    IF ($InfoCheckBox.checked -eq $true){
        $options.GUI_Config.MegaScript_Config.o3 = "1" 
    } else {
        $options.GUI_Config.MegaScript_Config.o4 = "0"
    }
    IF ($HTTPCheckBox.checked -eq $true){
        $options.GUI_Config.MegaScript_Config.o4 = "1"
    } else {
        $options.GUI_Config.MegaScript_Config.o4 = "0"
    }
    IF ($PowerPlanCheckBox.checked -eq $true){
        $options.GUI_Config.MegaScript_Config.o5 = "1"
    } else {
        $options.GUI_Config.MegaScript_Config.o5 = "0"
    }
    IF ($UpdatesCheckBox.checked -eq $true){
        $options.GUI_Config.MegaScript_Config.o6 = "1"
    } else {
        $options.GUI_Config.MegaScript_Config.o6 = "0"
    }
    IF ($NetworkCheckBox.checked -eq $true){
        $options.GUI_Config.MegaScript_Config.o7 = "1"
    } else {
        $options.GUI_Config.MegaScript_Config.o7 = "0"
    }
    IF ($ReliabilityCheckBox.checked -eq $true){
        $options.GUI_Config.MegaScript_Config.o8 = "1"
    } else {
        $options.GUI_Config.MegaScript_Config.o8 = "0"
    }

    $days = [int]$DaysTextbox.text
    $emails = [string]$EmailsTextbox.text
    
    $options.save(".\config.xml")

    .\MEGASCRIPT.ps1 -days $days -to $emails})

#}

#Opens file browser
$NamesListBtn.Add_Click({

    $openFileDialog = New-Object windows.forms.openfiledialog   
    $openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()   
    $openFileDialog.title = "Select Text File to Import"   
    $openFileDialog.filter = "All files (*.*)| *.*"   
    $openFileDialog.filter = ".txt Files|*.txt|All Files|*.*" 
    $openFileDialog.ShowHelp = $True   
        #Write-Host "Select Downloaded Settings File... (see FileOpen Dialog)" -ForegroundColor Green  
    $result = $openFileDialog.ShowDialog()    #Display the Dialog / Wait for user response 
        # in ISE you may have to alt-tab or minimize ISE to see dialog box 
    write-host $result 

    if ($result -eq "OK") { 

        $CustomNameList = Get-Content $OpenFileDialog.filename | Out-File .\PCList.txt
        Write-Host "Selected Names File:" $CustomNameList -ForegroundColor Green        
        Write-Host "Names list imported!" -ForegroundColor Green

    } else { Write-Host "Names list import cancelled!" -ForegroundColor Yellow} 

})

$Global:placeholder = 'localhost'

$ContinueBtn.Add_Click({

    if ($PCNamesTextbox.Text -ne ""){
        $PCNamesTextbox.Text -replace ", ", "`r`n" | Out-File .\PCList.txt
        $CustomNameList = Get-Content .\PCList.txt 
        #$lastcall = $CustomNameList | Select-Object -Last 1
        #if ($lastcall -ne 'localhost') {
        $placeholder | Out-File -append '.\PCList.txt'
        #}
        Write-Host "PC name(s) written to PCList.txt"

    } else {
        $CustomNameList = Get-Content .\PCList.txt 
        $CustomNameList -replace ", ", "`r`n" | Out-File .\PCList.txt
        $placeholder | Out-File -append '.\PCList.txt'
        Write-Host "PC name(s) written to PCList.txt"
    }

    $MainLabel.Visible               = $true
    $Seperator2.Visible              = $false
    $PCNamesLabel.Visible            = $false
    $CommaLabel.Visible              = $false
    $DayLabel.Visible                = $true
    $EmailLabel.Visible              = $false
    $CoresLabel.Visible              = $false
    $WarningLabel.Visible            = $false
    $PCNamesTextbox.Visible          = $false
    $DaysTextbox.Visible             = $true
    $EmailsTextbox.Visible           = $false
    $BiosCheckBox.Visible            = $true
    $RebootCheckBox.Visible          = $true
    $InfoCheckBox.Visible            = $true
    $HTTPCheckBox.Visible            = $true
    $PowerPlanCheckBox.Visible       = $true
    $UpdatesCheckBox.Visible         = $true
    $NetworkCheckBox.Visible         = $true
    $ReliabilityCheckBox.Visible     = $true
    $StartBtn.Visible                = $true
    $NamesListBtn.Visible            = $false
    $ContinueBtn.Visible             = $false
    $BackBtn.Visible                 = $true
    $OtherOptionsBtn.Visible         = $false

})

$BackBtn.Add_Click({
    $Seperator2.Visible              = $true
    $MainLabel.Visible               = $true
    $PCNamesLabel.Visible            = $true
    $CommaLabel.Visible              = $true
    $EmailLabel.Visible              = $true
    $WarningLabel.Visible            = $true
    $PCNamesTextbox.Visible          = $true
    $EmailsTextbox.Visible           = $true
    $NamesListBtn.Visible            = $true
    $ContinueBtn.Visible             = $true
    $OtherOptionsBtn.Visible         = $true
    $NetInfoLbl.Visible              = $false
    $NetTestStartBtn.Visible         = $false
    $NetTextbox.Visible              = $false
    $NetTestLbl.Visible              = $false
    $DayLabel.Visible                = $false
    $CoresLabel.Visible              = $false
    $PCWarning.Visible               = $false
    $BaseBenches.Visible             = $false
    $BaseBenches2.Visible            = $false
    $BaseBenches3.Visible            = $false
    $BaseBenches4.Visible            = $false
    $DaysTextbox.Visible             = $false
    $CoresTextbox.Visible            = $false
    $BiosCheckBox.Visible            = $false
    $RebootCheckBox.Visible          = $false
    $InfoCheckBox.Visible            = $false
    $HTTPCheckBox.Visible            = $false
    $PowerPlanCheckBox.Visible       = $false
    $UpdatesCheckBox.Visible         = $false
    $NetworkCheckBox.Visible         = $false
    $ReliabilityCheckBox.Visible     = $false
    $StartBtn.Visible                = $false
    $BackBtn.Visible                 = $false
    $PowerManagementBtn.Visible      = $false
    $StressTestBtn.Visible           = $false
    $NetTestBtn.Visible              = $false
})

$NetTestBtn.Add_Click({
    $NetInfoLbl.Visible              = $true
    $NetTextbox.Visible              = $true
    $NetTestLbl.Visible              = $true
    $NetTestStartBtn.Visible         = $true
    $Seperator2.Visible              = $false
    $MainLabel.Visible               = $true
    $PCNamesLabel.Visible            = $false
    $CommaLabel.Visible              = $false
    $DayLabel.Visible                = $false
    $EmailLabel.Visible              = $false
    $CoresLabel.Visible              = $false
    $PCWarning.Visible               = $false
    $WarningLabel.Visible            = $false
    $BaseBenches.Visible             = $false
    $BaseBenches2.Visible            = $false
    $BaseBenches3.Visible            = $false
    $BaseBenches4.Visible            = $false
    $PCNamesTextbox.Visible          = $false
    $EmailsTextbox.Visible           = $false
    $DaysTextbox.Visible             = $false
    $CoresTextbox.Visible            = $false
    $BiosCheckBox.Visible            = $false
    $RebootCheckBox.Visible          = $false
    $InfoCheckBox.Visible            = $false
    $HTTPCheckBox.Visible            = $false
    $PowerPlanCheckBox.Visible       = $false
    $UpdatesCheckBox.Visible         = $false
    $NetworkCheckBox.Visible         = $false
    $ReliabilityCheckBox.Visible     = $false
    $StartBtn.Visible                = $false
    $NamesListBtn.Visible            = $false
    $ContinueBtn.Visible             = $false
    $BackBtn.Visible                 = $true
    $OtherOptionsBtn.Visible         = $false
    $PowerManagementBtn.Visible      = $false
    $StressTestBtn.Visible           = $false
    $NetTestBtn.Visible              = $false
})

$OtherOptionsBtn.Add_Click({
    $Seperator2.Visible              = $false
    $MainLabel.Visible               = $true
    $PCNamesLabel.Visible            = $false
    $CommaLabel.Visible              = $false
    $DayLabel.Visible                = $false
    $EmailLabel.Visible              = $false
    $BaseBenches.Visible             = $true
    $BaseBenches2.Visible            = $true
    $BaseBenches3.Visible            = $true
    $BaseBenches4.Visible            = $true
    $CoresLabel.Visible              = $true
    $PCWarning.Visible               = $true
    $PCNamesTextbox.Visible          = $false
    $EmailsTextbox.Visible           = $false
    $DaysTextbox.Visible             = $false
    $CoresTextbox.Visible            = $true
    $NamesListBtn.Visible            = $false
    $ContinueBtn.Visible             = $false
    $BackBtn.Visible                 = $true
    $OtherOptionsBtn.Visible         = $false
    $PowerManagementBtn.Visible      = $true
    $StressTestBtn.Visible           = $true
    $NetTestBtn.Visible       = $true

    if ($PCNamesTextbox.Text -ne ""){
        $PCNamesTextbox.Text -replace ", ", "`r`n" | Out-File .\PCList.txt
        #$lastcall = $CustomNameList | Select-Object -Last 1
        #if ($lastcall -ne 'localhost') {
        $placeholder | Out-File -append '.\PCList.txt'
            #}
        Write-Host "PC name(s) written to PCList.txt"

    } else {
        $CustomNameList = Get-Content .\PCList.txt 
        $CustomNameList -replace ", ", "`r`n" | Out-File .\PCList.txt
        $placeholder | Out-File -append '.\PCList.txt'
        Write-Host "PC name(s) written to PCList.txt"
    }
})

#$NetTestBtn.Add_Click({.\Delete_Profile.ps1})

$PowerManagementBtn.Add_Click({.\ChangePowerManagement.ps1})

$StressTestBtn.Add_Click({
    $cores = [int]$CoresTextbox.text
   .\stress.ps1 -cores $cores
})

$NetTestStartBtn.Add_Click({
    [string]$NetTextbox.text | Out-File .\ServerPaths.txt
    #"\\caeisi01\public\Hold\Alex" 
    "\\localhost\c$" >> .\ServerPaths.txt

   #.\NetworkReportHTML.ps1 -Path (Get-Content .\ServerPaths.txt) -HTMLPath ".\" -Yellow 20 -Red 10 -Age 1
   .\Test-NetworkSpeed.ps1 -Path (Get-Content .\ServerPaths.txt) -Size 25 -Verbose >> .\ServerPaths.txt
   
})

#Write your logic code here

[void]$Form.ShowDialog()