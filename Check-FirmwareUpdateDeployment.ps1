###################################################################################################################
# Name: Check-FirmwareUpdateDeployment.ps1
# Author: Thomas Marcussen, Thomas@ThomasMarcussen.com
# Date: September, 2024
###################################################################################################################
<#
	.SYNOPSIS
	Script to check the deployment status of firmware updates via the Windows Update Deployment Service.
	
	.DESCRIPTION
	Script to check the deployment status of firmware updates on a Windows device by querying the Windows Update service.
	It searches for pending and installed firmware updates and reports the status.

	Script prerequisites:
	1. A minimum Windows PowerShell version of '5.1' is required to run this script.

	2. No additional modules are required as the script uses built-in COM objects to query Windows Update.

	3. The script must be run with administrator privileges to accurately access Windows Update deployment information.
	
	4. A stable internet connection is needed to query Windows Update for available firmware updates.

	.PARAMETER Verbose
	Enables detailed output about the process of checking firmware updates.
	
	.EXAMPLE
	.\Check-FirmwareUpdateDeployment.ps1 -Verbose
	# This example runs the script with detailed output to check for available and installed firmware updates.
	
	.EXAMPLE
	.\Check-FirmwareUpdateDeployment.ps1
	# This example runs the script silently, only outputting relevant firmware update information.
#>

# Function to check firmware update deployment via Windows Update
function Check-FirmwareUpdateDeployment {
    try {
        # Create a Windows Update session
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session

        # Create an Update Searcher object
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

        # Define search criteria: Not installed and of type 'Driver'
        $searchCriteria = "IsInstalled=0 and Type='Driver'"

        Write-Output "Searching for available firmware updates via Windows Update..."

        # Search for updates matching the criteria
        $searchResult = $UpdateSearcher.Search($searchCriteria)

        # Define keywords that typically identify firmware updates
        $firmwareKeywords = @("firmware", "BIOS", "UEFI", "System Firmware", "Embedded Controller", "Intel Management Engine")

        # Initialize an array to hold firmware updates
        $firmwareUpdates = @()

        # Iterate through the search results to filter firmware updates
        foreach ($update in $searchResult.Updates) {
            foreach ($keyword in $firmwareKeywords) {
                if ($update.Title -match $keyword) {
                    $firmwareUpdates += $update
                    break
                }
            }
        }

        # Output available firmware updates
        if ($firmwareUpdates.Count -gt 0) {
            Write-Output "`nFirmware updates available for deployment:"
            $firmwareUpdates | Select-Object Title, KBArticleIDs, Classification, IsDownloaded | Format-Table -AutoSize
        }
        else {
            Write-Output "`nNo firmware updates are currently available for deployment."
        }

        # Retrieve the total number of updates in history
        $historyCount = $UpdateSearcher.GetTotalHistoryCount()

        # Query the update history
        $updateHistory = $UpdateSearcher.QueryHistory(0, $historyCount)

        # Filter the history for installed firmware updates
        $installedFirmwareUpdates = $updateHistory | Where-Object {
            $_.Title -match "firmware" -or
            $_.Title -match "BIOS" -or
            $_.Title -match "UEFI" -or
            $_.Title -match "System Firmware" -or
            $_.Title -match "Embedded Controller" -or
            $_.Title -match "Intel Management Engine"
        }

        # Output installed firmware updates
        if ($installedFirmwareUpdates.Count -gt 0) {
            Write-Output "`nFirmware updates that have been installed:"
            $installedFirmwareUpdates | Select-Object Date, Title, Result | Format-Table -AutoSize
        }
        else {
            Write-Output "`nNo firmware updates have been installed yet."
        }

        # Optionally, check for pending firmware updates that are downloaded or not
        $pendingFirmwareUpdates = $firmwareUpdates | Where-Object { $_.IsDownloaded -eq $false }

        if ($pendingFirmwareUpdates.Count -gt 0) {
            Write-Output "`nFirmware updates pending download:"
            $pendingFirmwareUpdates | Select-Object Title, KBArticleIDs, Classification | Format-Table -AutoSize
        }
        else {
            Write-Output "`nNo firmware updates are pending download."
        }

    }
    catch {
        Write-Error "An error occurred while checking firmware updates: $_"
    }
}

# Execute the function
Check-FirmwareUpdateDeployment
