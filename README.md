# IMPORTANT NOTE

Update to most current release if you are using _v2.1_

# Purge-LogFiles.ps1

PowerShell script for Exchange Server 2013 environments to clean up Exchange and IIS log files.

## Description

This script deletes all Exchange and IIS logs older than X days from all Exchange 2013 servers that are fetched using the Get-ExchangeServer cmdlet.

The Exchange log file location is read from the environment variable and used to build an adminstrative UNC path for file deletions.

It is assumed that the Exchange setup path is IDENTICAL across all Exchange servers.

The IIS log file location is read from the local IIS metabase of the LOCAL server and is used to build an administrative UNC path for IIS log file deletions.

It is assumed that the IIS log file location is identical across all Exchange servers

## Requirements

- Utilizes the global function library found here: [http://scripts.granikos.eu](http://scripts.granikos.eu)
- Exchange Server 2013+
- Exchange Management Shell (EMS)

## Parameters

### DaysToKeep

Number of days Exchange and IIS log files should be retained, default is 30 days

### Auto

Switch to use automatic detection of the IIS and Exchange log folder paths

### IsEdge

Indicates the the script is executed on an Exchange Server holding the EDGE role. Without the switch servers holding the EDGE role are excluded

### RepositoryRootPath

Absolute path to a repository folder for storing copied log files and compressed archives. Preferably an UNC path. A new subfolder will be created for each Exchange server.

### ArchiveMode

Log file copy and archive mode. Possible values

* _None_ = All log files will be purged without being copied 
* _CopyOnly_ = Simply copy log files to the RepositoryRootPath
* _CopyAndZip_ = Copy logfiles and send copied files to compressed archive
* _CopyZipAndDelete_ = Same as CopyAndZip, but delete copied log files from RepositoryRootPath

### SendMail

Switch to send an Html report

### MailFrom

Email address of report sender

### MailTo

Email address of report recipient

### MailServer

SMTP Server for email report

## Examples

``` PowerShell
.\Purge-LogFiles -DaysToKeep 14
```

Delete Exchange and IIS log files older than 14 days

``` PowerShell
.\Purge-LogFiles -DaysToKeep 7 -Auto
```

Delete Exchange and IIS log files older than 7 days with automatic discovery

``` PowerShell
.\Purge-LogFiles -DaysToKeep 7 -Auto -SendMail -MailFrom postmaster@sedna-inc.com -MailTo exchangeadmin@sedna-inc.com -MailServer mail.sedna-inc.com
```

Delete Exchange and IIS log files older than 7 days with automatic discovery and send email report

``` PowerShell
.\Purge-LogFiles -DaysToKeep 14 -RepositoryRootPath \\OTHERSERVER\OtherShare\LOGS -ArchiveMode CopyZipAndDelete`
```

Delete Exchange and IIS log files older than 14 days, but copy files to a central repository and compress the log files before final deletion

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery

Download and vote at TechNet Gallery

* [https://gallery.technet.microsoft.com/Purge-Exchange-Server-2013-c2e03e72](https://gallery.technet.microsoft.com/Purge-Exchange-Server-2013-c2e03e72)

## Credits

Written by: Thomas Stensitzki

Stay connected:

* My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
* Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
* LinkedIn:	[http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
* Github: [https://github.com/Apoc70](https://github.com/Apoc70)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
* Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
* Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)

Additional Credits:

* Is-Admin function (c) by Michel de Rooij, michel@eightwone.com