<#
    .SYNOPSIS
    Purge Exchange 2013+ and IIS log files across multiple Exchange servers 
   
    Thomas Stensitzki
    (Based Based on the original script by Brian Reid, C7 Solutions (c)
    http://www.c7solutions.com/2013/04/removing-old-exchange-2013-log-files-html)
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 2.3.2, 2020-12-04

    Please post ideas, comments, and suggestions at GitHub.
 
    .LINK  
    http://scripts.Granikos.eu
	
    .DESCRIPTION
	
    This script deletes all Exchange and IIS logs older than X days from all Exchange 2013+ servers
    that are fetched using the Get-ExchangeServer cmdlet.

    The Exchange log file location is read from a variable and used to build an 
    administrative UNC path for file deletions. It is assumed that the Exchange setup 
    path is IDENTICAL across all Exchange servers.
  
    Optionally, you can use the Active Directory
    configuration partition to determine the Exchange install path dynamically, if supported
    in your Active Directory environment.
    
    The IIS log file location is read from the local IIS metabase of the LOCAL server
    and is used to build an administrative UNC path for IIS log file deletions.

    Currently, it is assumed that the IIS log file location is identical across all Exchange servers.

    .NOTES 
    Requirements 
    - Windows Server 2012 R2 or newer
    - Utilizes the global function library found here: http://scripts.granikos.eu
    - AciveDirectory PowerShell module
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
    2.0     Script update
    2.1     Log file archiving and archive compressions added
    2.11    Issue #6 fixed 
    2.13    Issue #7 fixed
    2.14    Issue #9 fixed
    2.2     Minor changes, but no fixes
    2.21    Issue #12 fixed
    2.3     Option for HTTPERR (#13) added, Option for dynamic Exchange install paths added, Html formatting added, tested with Exchange Server 2019
    2.3.1   Issues #14, #15
    2.3.2   Issue #16 fixed
	
    .PARAMETER DaysToKeep
    Number of days Exchange and IIS log files should be retained, default is 30 days

    .PARAMETER Auto
    Switch to use automatic detection of the IIS and Exchange log folder paths

    .PARAMETER IsEdge
    Indicates the the script is executed on an Exchange Server holding the EDGE role. Without the switch servers holding the EDGE role are excluded

    .PARAMETER IncludeHttpErr
    Include the HTTPERR log files in the purge routine. Those logs are normally stored at C:\Windows\System32\LogFiles\HTTPERR

    .PARAMETER UseDynamicExchangePaths
    Determine the Exchange install path by querying the server object in AD configuration partition. This helps if your Exchange servers do not have a unified install path across all servers.

    .PARAMETER RepositoryRootPath
    Absolute path to a repository folder for storing copied log files and compressed archives. Preferably an UNC path. A new subfolder will be created fpr each Exchange server.

    .PARAMETER ArchiveMode
    Log file copy and archive mode. Possible values
    None = All log files will be purged without being copied
    CopyOnly = Simply copy log files to the RepositoryRootPath
    CopyAndZip = Copy logfiles and send copied files to compressed archive
    CopyZipAndDelete = Same as CopyAndZip, bt delete copied log files from RepositoryRootPath

    .PARAMETER SendMail
    Switch to send an Html report

    .PARAMETER MailFrom
    Email address of report sender

    .PARAMETER MailTo
    Email address of report recipient

    .PARAMETER MailServer
    SMTP Server for email report
   
    .EXAMPLE
    Delete Exchange Server and IIS log files older than 14 days 

    .\Purge-LogFiles -DaysToKeep 14

    .EXAMPLE
    Delete Exchange Server and IIS log files older than 7 days with automatic discovery

    .\Purge-LogFiles -DaysToKeep 7 -Auto

    .EXAMPLE
    Delete Exchange Server and IIS log files older than 7 days with automatic discovery and send email report

    .\Purge-LogFiles -DaysToKeep 7 -Auto -SendMail -MailFrom postmaster@sedna-inc.com -MailTo exchangeadmin@sedna-inc.com -MailServer mail.sedna-inc.com 

    .EXAMPLE
    Delete Exchange and IIS log files older than 14 days, but copy files to a central repository and compress the log files before final deletion

    .\Purge-LogFiles -DaysToKeep 14 -RepositoryRootPath \\OTHERSERVER\OtherShare\LOGS -ArchiveMode CopyZipAndDelete

    .EXAMPLE
    Delete Exchange Server, IIS, and HTTPERR log files older than 7 days, and send an HTML email. Identify Exchange file paths using AD configuration objects.

    .\Purge-LogFiles.ps1 -DaysToKeep 7 -SendMail -MailFrom postmaster@sedna-inc.com -MailTo exchangeadmin@sedna-inc.com -MailServer mail.sedna-inc.com -UseDynamicExchangePaths -IncludeHttpErr

#>
[CmdletBinding()]
Param(
  [int]$DaysToKeep = 30,  
  [switch]$Auto,
  [switch]$IsEdge,
  [switch]$IncludeHttpErr,
  [switch]$UseDynamicExchangePaths,
  [switch]$SendMail,
  [string]$MailFrom = '',
  [string]$MailTo = '',
  [string]$MailServer = '',
  [string]$RepositoryRootPath = '\\MYSERVER\SomeShare\EXCHANGELOGS',
  [ValidateSet('None','CopyOnly','CopyAndZip','CopyZipAndDelete')] #Available archive modes, default: NONE
  [string]$ArchiveMode = 'None'
)

## Set fixed IIS and Exchange log paths 
## Examples: 
##   "C$\inetpub\logs\LogFiles"
##   "C$\Program Files\Microsoft\Exchange Server\V15\Logging"

## ADJUST AS NEEDED
[string]$IisUncLogPath = 'C$\inetpub\logs\LogFiles'
[string]$ExchangeUncLogPath = 'C$\Program Files\Microsoft\Exchange Server\V15\Logging'
[string]$ExchangeUncActiveMonitoringLogPath = 'C$\Program Files\Microsoft\Exchange Server\V15\Logging\Monitoring\Monitoring'
[string]$HttpErrUncLogPath= 'C$\Windows\System32\LogFiles\HTTPERR'

# log file extension filter
[string[]]$IncludeFilter = @('*.log')

# filename template for archived log files
[string]$ArchiveFileName =  "LogArchive $(Get-Date -Format 'yyyy-MM-dd HHmm').zip"

# Folder name for a per Exchange server log files
# Folder is used as target when copying log files
# Folder will be deleted when using ArchiveMode CopyZipAndDelete
[string]$LogSubfolderName = 'LOGS'

# Some error variables
$ERR_OK = 0
$ERR_COMPRESSIONFAILED = 1080
$ERR_NONELEVATEDMODE = 1099

# Preset some archive switches
[bool]$CopyFiles = $false
[bool]$ZipArchive = $false
[bool]$DeleteZippedFiles = $false

# Import Exchange functions
Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue

# Import Active Directory Module
Import-Module -Name ActiveDirectory

# 2015-06-18: Implementation of global functions module
Import-Module -Name GlobalFunctions
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

if($Auto) {
  # detect log file locations automatically and set variables

  [string]$ExchangeInstallPath = $env:ExchangeInstallPath
  [string]$ExchangeUncLogDrive = $ExchangeInstallPath.Split(':\')[0]
  $ExchangeUncLogPath = ('{0}$\{1}Logging\' -f ($ExchangeUncLogDrive), $ExchangeInstallPath.Remove(0,3))

  # Fetch local IIS log location from Metabase
  # IIS default location fixed 2015-02-02
  [string]$IisLogPath = ((Get-WebConfigurationProperty -Filter 'system.applicationHost/sites/siteDefaults' -Name logFile).directory).Replace('%SystemDrive%',$env:SystemDrive)

  # Extract drive letter and build log path
  [string]$IisUncLogDrive =$IisLogPath.Split(':\')[0] 
  $IisUncLogPath = $IisUncLogDrive + '$\' + $IisLogPath.Remove(0,3) 
}

function Get-ExchangeServerProperty {
  [CmdletBinding()]
  param (
    [string]$ServerName = '',
    [string]$PropertyName = ''
  )

  # get forest name as array
  $ADForest = (Get-ADForest).Name.Split('.')

  $PropertyValue = ''

  if($ADForest -ne $null) {
    $ServerExists = ((Get-Item -Path ("AD:\CN={3},CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN={0},CN=Microsoft Exchange,CN=Services,CN=Configuration,DC={1},DC={2}" -f (Get-OrganizationConfig).Name, $ADForest[0], $ADForest[1], $ServerName) -ErrorAction SilentlyContinue) -ne $null)

    if($ServerExists) { 
    
      Write-Verbose -Message ('Server Object {0} found in configuration partition' -f $ServerName)
    
      $PropertyValue = (Get-ItemProperty -Path ("AD:\CN={3},CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN={0},CN=Microsoft Exchange,CN=Services,CN=Configuration,DC={1},DC={2}" -f (Get-OrganizationConfig).Name, $ADForest[0], $ADForest[1], $ServerName) -Name $PropertyName -ErrorAction SilentlyContinue).$PropertyName

      Write-Verbose -Message ('| Property {0} = {1}' -f $PropertyName, $PropertyValue)
    }
  }

  $PropertyValue
}

function Get-ExchangeServerDynamicPath {
  [CmdletBinding()]
  param(
    [string]$ServerName = ''
  )

  $UncPurgePath = ''

  [string]$InstallPath = Get-ExchangeServerProperty -ServerName $ServerName -PropertyName msExchInstallPath

  if($InstallPath -ne '') {
    $UncPurgePath = ('{0}$\{1}\Logging' -f $InstallPath.Substring(0,1), $InstallPath.Substring(3))

    Write-Verbose -Message ('Calculated Exchange Server dynamic path: {0}' -f $UncPurgePath)
  }

  $UncPurgePath

}

function Copy-LogFiles {
  [CmdletBinding()]
  param(
    [string] $SourceServer = '',
    [string]$SourcePath = '',
    $FilesToMove,
    [string]$ArchivePrefix = ''
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

    }    
    
    if($ZipArchive) {
      # zip copied log files
      
      $Archive = Join-Path -Path $ServerRepositoryPath -ChildPath ('{0}-{1}' -f $ArchivePrefix, $ArchiveFileName)
      $logger.Write(('Zip copied files to {0}' -f $ArchiveFileName))
      
      # delete archive file, if already exists
      if(Test-Path -Path $Archive) {Remove-Item -Path $Archive -Force -Confirm:$false}

      try {
        # create zipped asrchive
        # Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [IO.Compression.ZipFile]::CreateFromDirectory($ServerRepositoryLogsPath,$Archive)
      }
      catch {
        $ErrorMessage = $_.Exception.Message
        $logger.Write(('Error compressing files from {0} to {1}. Error Message: {2}' -f $ServerRepositoryLogsPath, $Archive, $ErrorMessage),3)      
      }
      finally {

        # cleanup, if compression was successful
        if($DeleteZippedFiles) {

          $logger.Write(('Deleting folder {0}' -f $ServerRepositoryLogsPath))
          $null = Remove-Item -Path $ServerRepositoryLogsPath -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue

        }
      }
    } 
  }  
}

# Function to clean log files from remote servers using UNC paths
function Remove-LogFiles {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory, HelpMessage='Absolute path to log file source')]
    [string]$Path,
    [ValidateSet('IIS','Exchange')] 
    [string]$Type = 'IIS'
  )

  # Build full UNC path
  $TargetServerFolder = ('\\{0}\{1}' -f $E15Server, $Path)

  # Write progress bar for current activity
  Write-Progress -Activity ('Checking files on Server {0}' -f $E15Server) -Status ('Checking files in {0}' -f $TargetServerFolder) -PercentComplete(($i/$PurgeActionsPerServer)*100)

  # Only try to delete files, if folder exists
  if (Test-Path -Path $TargetServerFolder) {
        
    $LastWrite = (Get-Date).AddDays(-$DaysToKeep)

    # Select files to delete
    $Files = Get-ChildItem -Path $TargetServerFolder -Include $IncludeFilter -Recurse -ErrorAction SilentlyContinue | Where-Object {$_.LastWriteTime -le $LastWrite}
    $FilesToDelete = ($Files | Measure-Object).Count

    # Lets count the files that will be deleted
    $fileCount = 0

    if($FilesToDelete -gt 0) {

      if($CopyFiles) {

        # we want to copy all files to central repository before final deletion
        $logger.Write(('Copy {0} files from {1} to repository' -f $FilesToDelete.Count, $TargetServerFolder))

        # Write progress bar for current activity
        Write-Progress -Activity ('Checking files on Server {0}' -f $E15Server) -Status 'Copying log files' -PercentComplete(($i/$PurgeActionsPerServer)*100)

        Copy-LogFiles -SourceServer $E15Server -SourcePath $TargetServerFolder -FilesToMove $Files -ArchivePrefix $Type
      }
        
      # Delete the files -> The real purge
      foreach ($File in $Files) {
          
        $null = Remove-Item -Path $File -ErrorAction SilentlyContinue -Force 
          
        $fileCount++
      }
      
      $text  = ('{0} file(s) deleted in {1}' -f $fileCount, $TargetServerFolder)

      $logger.Write($text)

      #Html output
      $Output = ('<li>{0}</li>' -f $text)
    }
    else {
      $text = ('No files deleted in {0}' -f $TargetServerFolder)

      $logger.Write($text)

      #Html output
      $Output = ('<li>{0}</li>' -f $text)
    }
  }
  Else {
    # oops, folder does not exist or is not accessible
    Write-Warning ("The folder {0} doesn't exist or is not accessible! Check the folder path!" -f $TargetServerFolder) -ForegroundColor Red

    #Html output
    $Output = ("The folder {0} doesn't exist or is not accessible! Check the folder path!" -f $TargetServerFolder)
  }

  $Output
}

# Check if we are running in elevated mode
# function (c) by Michel de Rooij, michel@eightwone.com
Function script:Test-IsAdmin {
  $currentPrincipal = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList ( [Security.Principal.WindowsIdentity]::GetCurrent() )

  If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
    return $true
  }
  else {
    return $false
  }
}

# Check validity of parameters required for sending emails 
Function script:Test-SendMail {
  if( ($MailFrom -ne '') -and ($MailTo -ne '') -and ($MailServer -ne '') ) {
    return $true
  }
  else {
    return $false
  }
}

# Main -----------------------------------------------------

If ($SendMail.IsPresent) { 
  If (-Not (Test-SendMail)) {
    Throw 'If -SendMail specified, -MailFrom, -MailTo and -MailServer must be specified as well!'
  }
}

Switch($ArchiveMode) {
  'CopyOnly' {
    $CopyFiles = $true
  }
  'CopyAndZip' {
    $CopyFiles = $true
    $ZipArchive = $true
  }
  'CopyZipAndDelete' {
    $CopyFiles = $true
    $ZipArchive = $true
    $DeleteZippedFiles = $true
  }
  default { }
}

$HtmlHeader = @"
<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Frameset//EN"" ""http://www.w3.org/TR/html4/frameset.dtd"">
<html><head><title>Purge Log Files Report</title>
<style type="text/css">
<!--
body {
  font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
h5{ 
 font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
 clear: both; 
 font-size: 11px;
 color:000000; }
p{
  font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
  font-size: 10px;
  color:#000000;
}
ul{
  font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
  font-size: 10px;
  color:#000000;
}
-->
</style>
</head><body>
"@

If (Test-IsAdmin) {
  # We are running in elevated mode. Let's continue.

  Write-Output ('Removing IIS and Exchange logs - Keeping last {0} days - Be patient, it might take some time' -f $DaysToKeep)

  # Track script execution in Exchange Admin Audit Log 
  if(-not $IsEdge) {
    Write-AdminAuditLog -Comment 'Purge-LogFiles started!'
  }
  $logger.Write(('Purge-LogFiles started, keeping last {0} days of log files.' -f ($DaysToKeep)))
    
  $AllExchangeServers = $null

  if($IsEdge) { 
    # Just run on local server
    $AllExchangeServers = $env:COMPUTERNAME 
  }
  else {
    # Get a list of all Exchange V15* servers, exclude server service the EDGE role
    $AllExchangeServers = Get-ExchangeServer | Where-Object {($_.IsE15OrLater -eq $true) -and ($_.ServerRole -notmatch 'Edge')} | Sort-Object -Property Name
  }

  $logger.WriteEventLog(('Script started. Script will purge log files on: {0}' -f $AllExchangeServers))

  # Lets count the steps for a nice progress bar
  $i = 1

  # Issue #12, Using Measure-Object to ensure that we have a Count property
  $PurgeActions = 3
  if($IncludeHttpErr) {$PurgeActions = 4}
  
  $PurgeActionsPerServer = ($AllExchangeServers | Measure-Object).Count * $PurgeActions # purge actions to execute per server

  # Some text
  $HeaderText = '<p>'
  if($UseDynamicExchangePaths) {
    $HeaderText += 'Log files purged using static IIS and dynamic Exchange installation path(s).'
  }
  else {
    $HeaderText += 'Log files purged using static Exchange and IIS UNC paths.'
  }
  $HeaderText += '</p>'

  # Prepare Output
  $Output = ('{0}{1}' -f $HtmlHeader, $HeaderText)

  # Call function for each server and each directory type
  foreach ($E15Server In $AllExchangeServers) {

    $E15Server = ([string]$E15Server).ToUpper()
        
    Write-Verbose -Message ('Working on: {0}' -f $E15Server)

    $Output += ('<h5>{0}</h5>
    <ul>' -f $E15Server)

    # Remove IIS log files
    if(-not $IsEdge) {
        $Output += Remove-LogFiles -Path $IisUncLogPath -Type IIS
        $i++
    }

    # Remove Exchange files
    $ExchangePurgeUncLogPath = $ExchangeUncLogPath

    if($UseDynamicExchangePaths) {
      $ExchangePurgeUncLogPath = Get-ExchangeServerDynamicPath -ServerName $E15Server 
    }

    if($ExchangePurgeUncLogPath -ne '') {
      # call function to purge Exchange Server logs
      $Output += Remove-LogFiles -Path $ExchangePurgeUncLogPath -Type Exchange
    }
    else {
      Write-Verbose -Message ('Failed to get Exchange Install Path for server {0}' -f $E15Server)
    }
    $i++

    # Ensure dedicated purging of ActiveMonitoringTrace logs
    # Check \bin\MSExchangeHMWorker.exe.config settings for trace log file settings
    $Output += Remove-LogFiles -Path $ExchangeUncActiveMonitoringLogPath -Type Exchange

    # Remove HttpErr files
    if($IncludeHttpErr) { 

      $Output += Remove-LogFiles -Path $HttpErrUncLogPath -Type Exchange
      $i++

    }

    $Output+='</ul>'

  }


  # Finalize Output
  $Output+='</body></html>'

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

  Return $ERR_OK
}
else {
  # Ooops, the admin did it again.
  Write-Output 'The script need to be executed in elevated mode. Start the Exchange Management Shell with administrative privileges.'

  Return $ERR_NONELEVATEDMODE
}
