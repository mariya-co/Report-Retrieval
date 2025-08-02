<# Remote report retrieval script for powercfg /battery report and netsh wlan report - copies the report to your local machine.
Useful for IT teams to gather the reports without having to contact the user to setup a remote session.
#>
# Get the hostname of the machine you wish to check and the machine you're currently on
$remotePC = Read-Host 'Enter the hostname of the machine you wish to check'

# Define selection for whether the user wants to do a battery report or a WLAN report
function Show-Menu {
    param(
        [string[]]$selections
    )
    Write-Host -ForegroundColor Green 'Select what you want to run'
    for ($i = 0; $i -lt $selections.Length; $i++) {
        Write-Host -ForegroundColor Green "$($i + 1): $($selections[$i])"
    }
    $selection = Read-Host "Please select the report to run by number"
    return $selection
}

# Define the reports to run
$selectionOptions = @("BatteryReport", "WLANReport")

$selectedReport = Show-Menu -selections $selectionOptions

# Connect to remote machine and execute script block - first test if machine is online
$ping = Test-Connection $remotePC -Count 1 -Quiet
if ($ping) {
    Invoke-Command -ComputerName $remotePC -ScriptBlock {
        # Ensure temp exists, if not, create it
        if (-not(Test-Path -Path C:\temp)) {
            New-Item -Path C:\temp -ItemType Directory
        }

        switch ($using:selectedReport) {
            1 {
                powercfg.exe /batteryreport /output "C:\temp\$($using:remotePC)-batreport.html"
                $reportPath = "C:\temp\$($using:remotePC)-batreport.html"
            }
            2 {
                netsh wlan show wlanreport
                $reportPath = "C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html"
                Copy-Item -Path $reportPath -Destination "C:\temp\$($using:remotePC)-wlanreport.html"
                $reportPath = "C:\temp\$($using:remotePC)-wlanreport.html"
            }
        }

        # Return the path of the generated report
        return $reportPath
    }
} else {
    Write-Host -ForegroundColor Red "Machine is offline - try again once the machine is online"
    exit
}

# Copy report to your local machine using try-catch, output to terminal some tips on retrieving manually
try {
    Copy-Item -Path "\\$remotePC\c$\temp\$remotePC-$(if ($selectedReport -eq 1) {'batreport.html'} else {'wlanreport.html'})" -Destination "C:\temp" -ErrorAction Stop
    Write-Host -ForegroundColor Cyan "Copied report to your machine in C:\temp directory"
} catch {
    Write-Output $_.Exception
    Write-Host -ForegroundColor Red "Failed to copy the file due to the error above, note that you may be able to retrieve the file manually by navigating to the directory in file explorer."
}
