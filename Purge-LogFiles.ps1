<#
    .SYNOPSIS
    Purge Exchange 2013 and IIS logfiles on all Exchange servers 
   
   	Thomas Stensitzki
    (Based Based on the original script by Brian Reid, C7 Solutions (c)
    http://www.c7solutions.com/2013/04/removing-old-exchange-2013-log-files-html)
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.2, 2014-12-10

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    More information can be found at http://www.granikos.eu/en/scripts 
	
    .DESCRIPTION
	
    This script deletes all Exchange and IIS logs older than X days from all Exchange 2013 servers
    that are fetched using the Get-ExchangeServer cmdlet.

    The Exchange log file location is read from the environment variable and used to build an 
    adminstrative UNC path for file deletions.

    It is assumed that the Exchange setup path is IDENTICAL across all Exchange servers.

    The IIS log file location is read from the local IIS metabase of the LOCAL server
    and is used to build an administrative UNC path for IIS log file deletions.

    It is assumed that the IIS log file location is identical across all Exchange servers.

    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
    1.1     Variable fix and optional code added 
    1.2     Auto/Manual configration options added
    1.3     Check if running in elevated mode added
    1.4     Handling of IIS default location fixed
    1.5     Sorting of server names added and Write-Host output changed
	
	.PARAMETER DaysToKeep
    Number of days Exchange and IIS log files should be retained, default is 30 days

    .PARAMETER Auto
    Switch to use automatic detection of the IIS and Exchange log folder paths
   
	.EXAMPLE
    Delete Exchange and IIS log files older than 14 days 
    .\Purge-LogFiles -DaysToKeep 14
	
    #>
Param(
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Number of days for log files retention')][int]$DaysToKeep = 30,  
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Use automatic folder detection for Exchange and IIS log paths')][switch]$Auto  
)

Set-StrictMode -Version Latest

## Set fixed IIS and Exchange log paths 
## Examples: 
##   "C$\inetpub\logs\LogFiles"
##   "C$\Program Files\Microsoft\Exchange Server\V15\Logging"

[string]$IisUncLogPath = "C$\inetpub\logs\LogFiles"
[string]$ExchangeUncLogPath = "C$\Program Files\Microsoft\Exchange Server\V15\Logging"

if($Auto) {
    # detect log file locations automatically an set variables

    [string]$ExchangeInstallPath = $env:ExchangeInstallPath
    [string]$ExchangeUncLogDrive = $ExchangeInstallPath.Split(":\")[0]
    $ExchangeUncLogPath = $ExchangeUncLogDrive + "$\" + $ExchangeInstallPath.Remove(0,3) + "Logging\"

    # Fetch local IIS log location from Metabase
    # IIS default location fixed 2015-02-02
    [string]$IisLogPath = ((Get-WebConfigurationProperty "system.applicationHost/sites/siteDefaults" -Name logFile).directory).Replace("%SystemDrive%",$env:SystemDrive)

    # Extract drive letter and build log path
    [string]$IisUncLogDrive =$IisLogPath.Split(":\")[0] 
    $IisUncLogPath = $IisUncLogDrive + "$\" + $IisLogPath.Remove(0,3) 
}

# Function to clean log files from remote servers using UNC paths
Function CleanLogFiles   
{
    Param([string]$path)

    # Build full UNC path
    $TargetServerFolder = "\\" + $E15Server + "\" + $path

    # Write progress bar for current activity
    Write-Progress -Activity "Checking Server $E15Server" -Status "Checking files in $TargetServerFolder" -PercentComplete(($i/$max)*100)

    # Only try to delete files, if folder exists
    if (Test-Path $TargetServerFolder) {
        
        $Now = Get-Date
        $LastWrite = $Now.AddDays(-$DaysToKeep)

        # Select files to delete
        $Files = Get-ChildItem $TargetServerFolder -Include *.log -Recurse | Where {$_.LastWriteTime -le "$LastWrite"}

        # Lets count the files that will be deleted
        $fileCount = 0

        # Delete the files
        foreach ($File in $Files)
            {
                Remove-Item $File -ErrorAction SilentlyContinue | out-null
                $fileCount++
                }
        
        Write-Host "--> $fileCount files deleted in $TargetServerFolder" -ForegroundColor Gray
    }
    Else {
        # oops, folder does not exist or is not accessible
        Write-Host "The folder $TargetServerFolder doesn't exist or is not accessible! Check the folder path!" -ForegroundColor "red"
    }
}

# Check if we are running in elevated mode
# function (c) by Michel de Rooij, michel@eightwone.com
Function Is-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
    If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    }
    else {
        return $false
    }
}

# Main -----------------------------------------------------

If (Is-Admin) {
    # We are running in elevated mode. Let's continue.

    Write-Output "Removing IIS and Exchange logs - Keeping last $DaysToKeep days - Be patient, it might take some time"

    # Track script execution in Exchange Admin Audit Log 
    Write-AdminAuditLog -Comment "Purge-LogFiles started!"

    # Get a list of all Exchange 2013 servers
    $Ex2013 = Get-ExchangeServer | Where {$_.IsE15OrLater -eq $true} | Sort-Object Name

    # Lets count the steps for a nice progress bar
    $i = 1
    $max = $Ex2013.Count * 2 # two actions to execute per server

    # Call function for each server and each directory type
    foreach ($E15Server In $Ex2013) {
        Write-Host "Working on: $E15Server" -ForegroundColor Gray

        CleanLogFiles -path $IisUncLogPath
        $i++

        CleanLogfiles -path $ExchangeUncLogPath
        $i++
    }
}
else {
    # Ooops, the admin did it again.
    Write-Output "The script need to be executed in elevated mode. Start the Exchange Management Shell as Administrator."
}