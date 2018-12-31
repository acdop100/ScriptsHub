# ScriptsHub
Multi-purpose PowerShell script to remotely run commands on networked computers.

*Script must be run as an administrator! Do this by right-clicking PowerShell and selecting "Run as administrator".*

Functions include:
- Collecting a diagnostic report on PC's:
    - Options for what you want included in report.
    - Emailed to as many people as you want.
    - Report generated in nice HTML page.
        - Tables.
        - Collapsable sections.
        - Color coding.
- Changing power management settings on other machines.
- Stress tests for CPU, RAM, and network conditions on machines. Displays comparisons against baseline benchmarks.
- Includes persistant options saving throughout runs 
- Can use local or external SMTP server to send report supplied by the user on the first run. 
- Can also set the SMTP server you choose to be the default SMTP server for all of Powershell

All options are located in config.xml. Should be self-explanatory for the most part.

The file PCList.txt is there for you to copy all of the PC's names which you need to test and you can select that file to use for testing. If you choose to input the names into the program's input box, they will be saved to PCList.txt. 
