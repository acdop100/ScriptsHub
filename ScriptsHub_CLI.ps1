#Requires -RunAsAdministrator

$options = [xml](Get-Content ".\config.xml")

IF ($options.GUI_Config.Mail_Config.Host_Email -eq ""){
    $FirstRun = $true
}

function Show-Menu {
    Clear-Host
    Write-Host "========================= Menu ========================="
     
    Write-Host "1: Press '1' to create PC diagnostic report."
    Write-Host "2: Press '2' to change PC power management settings."
    Write-Host "3: Press '3' to run benchmarks."
    Write-Host "4: Press '4' to run network test."
    Write-Host "Q: Press 'Q' to quit."
}


$email = Read-Host "Enter E-Mail address to send report to or hit enter to save report to Desktop"
switch($email){
    '' {
        Write-Host Saving report to Desktop -ForegroundColor Green
    }
    else {
        switch ($FirstRun) {
            $true {
                
                $SMPT_addr = [string]$options.GUI_Config.Mail_Config.SMTP_Host
                $Host_email = [string]$options.GUI_Config.Mail_Config.Host_Email

            } else {    
                $SMPT_addr = Read-Host "Enter SMTP Address"   
                $Host_email = Read-Host "Enter Report Sender Address"

                $options.GUI_Config.Mail_Config.SMTP_Host = [string]$SMPT_addr
                $options.GUI_Config.Mail_Config.Host_Email = [string]$Host_email

                $options.save(".\config.xml")

                $wshell = New-Object -ComObject Wscript.Shell
                $response = $wshell.Popup("Make this SMTP default for all PowerShell?",0,"Saved!",32+4)

                IF ($response -eq 6) {
                    $PSEmailServer = [string]$FirstRunInput1.text
                }
            }
        }
    }
}

Show-Menu
$input = Read-Host "Please make a selection"
switch ($input) {
    '1' {
        Clear-Host
        'You chose option #1'
        Write-Host Generating Diagnostic Report -ForegroundColor Green        
        Write-Host 'Options are based off of selected options in config.xml' -ForegroundColor Green     

        $days = Read-Host "How many days back to you want to trace?"

        .\MEGASCRIPT.ps1 -days $days -to $email

    } '2' {
        Clear-Host
        'You chose option #2'
        Write-Host Changing PowerManagement Power Plan -ForegroundColor Green
        .\ChangePowerManagement.ps1
        
    } '3' {
        Clear-Host
        'You chose option #3'
        Write-Host Benchmarks -ForegroundColor Green

        $cores = Read-Host "How many cores do you want to stress? (recommended 2-4)"

        .\stress.ps1 -cores $cores
        
    } '4' {
        Clear-Host
        'You chose option #4'
        Write-Host Network Test -ForegroundColor Green

        $cores = Read-Host "Network test not currently working"
        return

        
    } 'q' {
        Clear-Host
        return
    }
}