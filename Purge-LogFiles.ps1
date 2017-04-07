<#
  .SYNOPSIS
  Purge Exchange 2013 and IIS logfiles across Exchange servers 
   
  Thomas Stensitzki
  (Based Based on the original script by Brian Reid, C7 Solutions (c)
  http://www.c7solutions.com/2013/04/removing-old-exchange-2013-log-files-html)
	
  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
  Version 2.0, 2017-04-07

  Ideas, comments and suggestions to support@granikos.eu 
 
  .LINK  
  http://scripts-Granikos.eu
	
  .DESCRIPTION
	
  This script deletes all Exchange and IIS logs older than X days from all Exchange 2013+ servers
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
  - Utilites global function library found here: 
  - Exchange 2013+ Management Shell

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
  2.0     Script update, CopyFilesBeforeDelete implemented
	
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
  
  .PARAMETER CopyFilesBeforeDelete
  Switch to copy log files to a central repository (UNC) before final deletion
  Configure appropriate location in the script
  
  .PARAMETER ZipArchive
  Create a zipped archive after sucessfully copying log file to repository.
  CURRENTLY IN DEVELOPMENT
   
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
[CmdletBinding()]
Param(
  [int]$DaysToKeep = 30,  
  [switch]$Auto,
  [switch]$SendMail,
  [string]$MailFrom = '',
  [string]$MailTo = '',
  [string]$MailServer = '',
  [switch]$CopyFilesBeforeDelete,
  [switch]$ZipArchive
)

## Set fixed IIS and Exchange log paths 
## Examples: 
##   "C$\inetpub\logs\LogFiles"
##   "C$\Program Files\Microsoft\Exchange Server\V15\Logging"

[string]$IisUncLogPath = 'E$\IISLogs'
[string]$ExchangeUncLogPath = 'F$\Program Files\Microsoft\Exchange Server\V15\Logging'
[string]$RepositoryRootPath = '\\MYSERVER\E$\PURGEREPOSITORY'
[string[]]$IncludeFilter = @('*.log')
[string]$ArchiveFileName =  "LogArchive $(Get-Date -Format 'yyyy-MM-dd').zip"
[string]$LogSubfolderName = 'LOGS'

# 2015-06-18: Implementationof global module
Import-Module -Name GlobalFunctions
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write("Script started")

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

function Copy-LogFiles {
  [CmdletBinding()]
  param(
    [string]$SourceServer,
    [string]$SourcePath,
    $FilesToMove
  )

  if($SourceServer -ne '') { 

    # path per SERVER for zipped archives
    $ServerRepositoryPath = Join-Path -Path $RepositoryRootPath -ChildPath $SourceServer

    # subfolder used as target for copying source folders and files
    $ServerRepositoryLogsPath = Join-Path -Path $ServerRepositoryPath -ChildPath $LogSubfolderName

    $ServerRepositoryPath = Join-Path -Path $RepositoryRootPath -ChildPath $SourceServer

    if(!(Test-Path -Path $ServerRepositoryPath)) {
      # Create new target directory for server, if does not exist
      $null = New-Item -Path $ServerRepositoryPath -ItemType Directory -Force -Confirm:$false
    }

    foreach ($File in $FilesToMove) {
      # target directory
      $targetDir = $File.DirectoryName.Replace($TargetServerFolder, $ServerRepositoryLogsPath)

      # target file path
      $targetFile = $File.FullName.Replace($TargetServerFolder, $ServerRepositoryLogsPath)
      
      # create target directory, if not exists
      if(!(Test-Path -Path $targetDir)) {$null = mkdir -Path $targetDir}

      # copy file to target
      $null = Copy-Item -Path $File.FullName -Destination $targetFile -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue

    }-Force   
    
    if($ZipArchive) {
      # zip copied log files
      #
      <# NOT FULLY TESTED YET 
      $Archive = Join-Path -Path $ServerRepositoryPath -ChildPath $ArchiveFileName
      $logger.Write(('Zip copied files to {0}' -f $ArchiveFileName))

      if(Test-Path -Path $Archive) {Remove-Item $Archive -Force -Confirm:$false}

      Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
      [IO.Compression.ZipFile]::CreateFromDirectory($ServerRepositoryLogsPath,$Archive)

      #>
    } 
  }  
}

# Function to clean log files from remote servers using UNC paths
function Remove-LogFiles {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory, HelpMessage='Absolute path to log file source')]
    [string]$Path
  )

  # Build full UNC path
  $TargetServerFolder = ('\\{0}\{1}' -f ($E15Server), ($Path))

  # Write progress bar for current activity
  Write-Progress -Activity ('Checking Server {0}' -f $E15Server) -Status ('Checking files in {0}' -f $TargetServerFolder) -PercentComplete(($i/$max)*100)

  # Only try to delete files, if folder exists
  if (Test-Path -Path $TargetServerFolder) {
        
      $LastWrite = (Get-Date).AddDays(-$DaysToKeep)

      # Select files to delete
      $Files = Get-ChildItem -Path $TargetServerFolder -Include $IncludeFilter -Recurse | Where-Object {$_.LastWriteTime -le $LastWrite}
      $FilesToDelete = ($Files | Measure-Object).Count

      # Lets count the files that will be deleted
      $fileCount = 0

      if($FilesToDelete -gt 0) {

        if($CopyFilesBeforeDelete) {

          # we want to copy all files to central repository before deletion
          $logger.Write(('Copy {0} files from {1} to repository' -f $FilesToDelete.Count, $TargetServerFolder))

          Copy-LogFiles -SourceServer $E15Server -SourcePath $TargetServerFolder -FilesToMove $Files
        }

        # Delete the files
        foreach ($File in $Files) {

            if($CopyFilesBeforeDelete) {
                # 2016-11-16: TST Copy to central repository before file will be deleted
                $logger.Write('Copy to repository')
            }

            $null = Remove-Item -Path $File -ErrorAction SilentlyContinue -Force
            $fileCount++
        }

        # Write-Host "--> $fileCount files deleted in $TargetServerFolder" -ForegroundColor Gray

        $logger.Write(('{0} files deleted in {1}' -f $fileCount, $TargetServerFolder))

        #Html output
        $Output = ("<li>{0} files deleted in '{1}'</li>" -f $fileCount, $TargetServerFolder)
      }
      else {
        $logger.Write(('No files to delete in {0}' -f $TargetServerFolder))

        #Html output
        $Output = ("<li>No files to delete in '{1}'</li>" -f $TargetServerFolder)
      }
  }
  Else {
      # oops, folder does not exist or is not accessible
      Write-Host ("The folder {0} doesn't exist or is not accessible! Check the folder path!" -f $TargetServerFolder) -ForegroundColor Red

      #Html output
      $Output = ("The folder {0} doesn't exist or is not accessible! Check the folder path!" -f $TargetServerFolder)
  }

  $Output
}

# Check if we are running in elevated mode
# function (c) by Michel de Rooij, michel@eightwone.com
Function Get-IsAdmin {
    $currentPrincipal = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList ( [Security.Principal.WindowsIdentity]::GetCurrent() )

    If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    }
    else {
        return $false
    }
}

Function Check-SendMail {
     if( ($SendMail) -and ($MailFrom -ne "") -and ($MailTo -ne "") -and ($MailServer -ne "") ) {
        return $true
     }
     else {
        return $false
     }
}

# Main -----------------------------------------------------
If (-Not (Check-SendMail)) {
    Throw 'If -SendMail specified, -MailFrom, -MailTo and -MailServer must be specified as well!'
}

If (Get-IsAdmin) {
    # We are running in elevated mode. Let's continue.

    Write-Output ('Removing IIS and Exchange logs - Keeping last {0} days - Be patient, it might take some time' -f $DaysToKeep)

    # Track script execution in Exchange Admin Audit Log 
    Write-AdminAuditLog -Comment "Purge-LogFiles started!"
    $logger.Write(('Purge-LogFiles started, keeping last {0} days of log files.' -f ($DaysToKeep)))

    # Get a list of all Exchange 2013 servers
    $Ex2013 = Get-ExchangeServer | Where-Object {$_.IsE15OrLater -eq $true} | Sort-Object -Property Name

    $logger.WriteEventLog(('Script started. Script will purge log files on: {0}' -f $Ex2013))

    # Lets count the steps for a nice progress bar
    $i = 1
    $max = $Ex2013.Count * 2 # two actions to execute per server

    # Prepare Output
    $Output = '<html>
    <body>
    <font size=""1"" face=""Arial,sans-serif"">'

    # Call function for each server and each directory type
    foreach ($E15Server In $Ex2013) {
        # Write-Host "Working on: $E15Server" -ForegroundColor Gray

        $Output += ('<h5>{0}</h5>
        <ul>' -f $E15Server)

        # Remove IIS log files
        $Output += Remove-LogFiles -Path $IisUncLogPath
        $i++

        # Remove Exchange files
        $Output += Remove-LogFiles -Path $ExchangeUncLogPath
        $i++

        $Output+='</ul>'

    }

    # Finalize Output
    $Output+='</font>
    </body>
    </html>'

    if($SendMail) {
        $logger.Write(('Sending email to {0}' -f $MailTo))
        try {
            Send-Mail -From $MailFrom -To $MailTo -SmtpServer $MailServer -MessageBody $Output -Subject 'Purge-Logfiles Report'         
        }
        catch {
            $logger.Write(('Error sending email to {0}' -f $MailTo),3)
        }
    }

    $logger.Write('Script finished')

    Return 0
}
else {
    # Ooops, the admin did it again.
    Write-Output 'The script need to be executed in elevated mode. Start the Exchange Management Shell as Administrator.'

    Return 99
}