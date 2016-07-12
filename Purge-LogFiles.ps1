<#
    .SYNOPSIS
    Purge Exchange 2013 and IIS logfiles across Exchange servers 
   
   	Thomas Stensitzki
    (Based Based on the original script by Brian Reid, C7 Solutions (c)
    http://www.c7solutions.com/2013/04/removing-old-exchange-2013-log-files-html)
	
	  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	  Version 1.94, 2016-07-07

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
    - Utilites global function library, 

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
    1.1     Variable fix and optional code added 
    1.2     Auto/Manual configration options added
    1.3     Check if running in elevated mode added
    1.4     Handling of IIS default location fixed
    1.5     Sorting of server names added and Write-Host output changed
    1.6     Count Error fixed
    1.7		  Email report functionality added
	  1.8     Support for global logging and other functions added
    1.9     Global functions updated (write to event log)
    1.91    Write DaysToKeep to log
    1.92    .Count issue fixed to run on Windows Server 2012
    1.93    Minor chances to PowerShell hygiene
    1.94    SendMail issue fixed (Thanks to denisvm, https://github.com/denisvm)
	
	  .PARAMETER DaysToKeep
    Number of days Exchange and IIS log files should be retained, default is 30 days

    .PARAMETER Auto
    Switch to use automatic detection of the IIS and Exchange log folder paths

    .PARAMETER SendMail
    Switch to send an Html report

    .PARAMETER MailFrom
    Email address of report sender

    .PARAMETER MailTo
    Email address of report recipient

    .PARAMETER MailServer
    SMTP Server for email report
   
	  .EXAMPLE
    Delete Exchange and IIS log files older than 14 days 
    .\Purge-LogFiles -DaysToKeep 14

    .EXAMPLE
    Delete Exchange and IIS log files older than 7 days with automatic discovery
    .\Purge-LogFiles -DaysToKeep 7 -Auto

    .EXAMPLE
    Delete Exchange and IIS log files older than 7 days with automatic discovery and send email report
    .\Purge-LogFiles -DaysToKeep 7 -Auto -SendMail -MailFrom postmaster@sedna-inc.com -MailTo exchangeadmin@sedna-inc.com -MailServer mail.sedna-inc.com 

    #>
Param(
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Number of days for log files retention')]
        [int]$DaysToKeep = 30,  
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Use automatic folder detection for Exchange and IIS log paths')]
        [switch]$Auto,
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Send result summary as email')]
        [switch]$SendMail,
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Sender address for result summary')]
        [string]$MailFrom = '',
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Recipient address for result summary')]
        [string]$MailTo = '',
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='SMTP Server address for sending result summary')]
        [string]$MailServer = ''
)

Set-StrictMode -Version Latest

## Set fixed IIS and Exchange log paths 
## Examples: 
##   "C$\inetpub\logs\LogFiles"
##   "C$\Program Files\Microsoft\Exchange Server\V15\Logging"

[string]$IisUncLogPath = "D$\IISLogs"
[string]$ExchangeUncLogPath = "E$\Program Files\Microsoft\Exchange Server\V15\Logging"

# 2015-06-18: Implementationof global module
Import-Module GlobalFunctions
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

if($Auto) {
    # detect log file locations automatically an set variables

    [string]$ExchangeInstallPath = $env:ExchangeInstallPath
    [string]$ExchangeUncLogDrive = $ExchangeInstallPath.Split(':\')[0]
    $ExchangeUncLogPath = $ExchangeUncLogDrive + "$\" + $ExchangeInstallPath.Remove(0,3) + 'Logging\'
    Join-Path -
    # Fetch local IIS log location from Metabase
    # IIS default location fixed 2015-02-02
    [string]$IisLogPath = ((Get-WebConfigurationProperty 'system.applicationHost/sites/siteDefaults' -Name logFile).directory).Replace('%SystemDrive%',$env:SystemDrive)

    # Extract drive letter and build log path
    [string]$IisUncLogDrive =$IisLogPath.Split(':\')[0] 
    $IisUncLogPath = $IisUncLogDrive + "$\" + $IisLogPath.Remove(0,3) 
}

# Function to clean log files from remote servers using UNC paths
Function Remove-LogFiles   
{
    Param([string]$Path)

    # Build full UNC path
    $TargetServerFolder = '\\' + $E15Server + '\' + $path

    # Write progress bar for current activity
    Write-Progress -Activity "Checking Server $E15Server" -Status "Checking files in $TargetServerFolder" -PercentComplete(($i/$max)*100)

    # Try to delete files only if folder exists
    if (Test-Path $TargetServerFolder) {
        
        $Now = Get-Date
        $LastWrite = $Now.AddDays(-$DaysToKeep)

        # Select files to delete
        $Files = Get-ChildItem $TargetServerFolder -Include *.log -Recurse | Where-Object {$_.LastWriteTime -le "$LastWrite"}

        # Lets count the files that will be deleted
        $fileCount = 0

        # Delete the files
        foreach ($File in $Files)
            {
                Remove-Item $File -ErrorAction SilentlyContinue -Force | out-null
                $fileCount++
                }

        #Write-Host "--> $fileCount files deleted in $TargetServerFolder" -ForegroundColor Gray

        $logger.Write("$($fileCount) files deleted in $($TargetServerFolder)")

        $Output = "<li>$fileCount files deleted in '$TargetServerFolder'</li>"
    }
    Else {
        # oops, folder does not exist or is not accessible
        Write-Host "The folder $TargetServerFolder doesn't exist or is not accessible! Check the folder path!" -ForegroundColor 'red'

        $Output = "The folder $TargetServerFolder doesn't exist or is not accessible! Check the folder path!"
    }

    $Output
}

# Check if we are running in elevated mode
# function (c) by Michel de Rooij, michel@eightwone.com
Function Get-IsAdmin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
    If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    }
    else {
        return $false
    }
}

Function Get-CheckSendMail {
     if (-Not ($SendMail) -or ( ($SendMail) -and ($MailFrom -ne '') -and ($MailTo -ne '') -and ($MailServer -ne '') ) ) {
        return $true
     }
     else {
        return $false
     }
}

# Main -----------------------------------------------------
If (-Not (Get-CheckSendMail)) {
    Throw 'If -SendMail specified, -MailFrom, -MailTo and -MailServer must be specified as well!'
}

If (Get-IsAdmin) {
    # We are running in elevated mode. Let's continue.

    Write-Output "Removing IIS and Exchange logs - Keeping last $DaysToKeep days - Be patient, it might take some time"

    # Track script execution in Exchange Admin Audit Log 
    Write-AdminAuditLog -Comment 'Purge-LogFiles started!'
    $logger.Write("Purge-LogFiles started, keeping last $($DaysToKeep) days of log files.")

    # Get a list of all Exchange 2013 servers
    $Ex2013 = Get-ExchangeServer | Where-Object {$_.IsE15OrLater -eq $true} | Sort-Object Name

    $logger.WriteEventLog("Script started. Script will purge log files on: $($Ex2013)")

    # Lets count the steps for a nice progress bar
    $i = 1
    $max = ($Ex2013 | Measure-Object).Count * 2 # two actions to execute per server

    # Prepare Output
    $Output = '<html>
    <body>
    <font size="1" face="Arial,sans-serif">'

    # Call function for each server and each directory type
    foreach ($E15Server In $Ex2013) {
        # Write-Host "Working on: $E15Server" -ForegroundColor Gray

        $Output += "<h5>$E15Server</h5>
        <ul>"

        $Output += Remove-LogFiles -Path $IisUncLogPath
        $i++

        $Output += Remove-LogFiles -Path $ExchangeUncLogPath
        $i++

        $Output+='</ul>'

    }

    # Finalize Output
    $Output+='</font>
    </body>
    </html>'

    if($SendMail) {
        $logger.Write("Sending email to $($MailTo)")
        Send-Mail -From $MailFrom -To $MailTo -SmtpServer $MailServer -MessageBody $Output -Subject 'Purge-Logfiles Report'         
    }

    $logger.Write('Script finished')
}
else {
    # Ooops, the admin did it again.
    Write-Output 'The script need to be executed in elevated mode. Start the Exchange Management Shell as Administrator.'
}