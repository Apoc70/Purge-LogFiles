# IMPORTANT NOTE

Update to most current release if you are using _v2.1_

# Purge-LogFiles.ps1

PowerShell script for modern Exchange Server environments to clean up Exchange Server and IIS log files.

## Description

This script deletes all Exchange Server and IIS logs older than X days from all Exchange 2013+ servers that are fetched using the Get-ExchangeServer cmdlet.

The Exchange log file location is read from a variable and used to build an administrative UNC path for file deletions. It is assumed that the Exchange setup path is IDENTICAL across all Exchange servers.

Optionally, you can use the Active Directory configuration partition to determine the Exchange install path dynamically, if supported in your Active Directory environment.

The IIS log file location is read from the local IIS metabase of the LOCAL server and is used to build an administrative UNC path for IIS log file deletions.

Currently, it is assumed that the IIS log file location is identical across all Exchange servers.

## Requirements

- Windows Server 2012 R2 or newer
- Utilizes the global function library found here: [http://scripts.granikos.eu](http://scripts.granikos.eu)
- AciveDirectory PowerShell module
- Exchange Server 2013+
- Exchange Management Shell (EMS)

## Updates

- 2020-05-14, v2.3.1, Issues #14, #15 fixed to properly support Edge Transport Role
- 2020-03-12, v2.3, Option for HTTPERR added, Option for dynamic Exchange install paths added, Html formatting added, tested with Exchange Server 2019

## Parameters

### DaysToKeep

Number of days Exchange and IIS log files should be retained, default is 30 days

### Auto

Switch to use automatic detection of the IIS and Exchange log folder paths

### IsEdge

Indicates the the script is executed on an Exchange Server holding the EDGE role. Without the switch servers holding the EDGE role are excluded.

### IncludeHttpErr

Include the HTTPERR log files in the purge routine. Those logs are normally stored at _C:\Windows\System32\LogFiles\HTTPERR_.

### UseDynamicExchangePaths

Determine the Exchange install path by querying the server object in AD configuration partition. This helps if your Exchange servers do not have a unified install path across all servers.

### RepositoryRootPath

Absolute path to a repository folder for storing copied log files and compressed archives. Preferably an UNC path. A new subfolder will be created for each Exchange server.

### ArchiveMode

Log file copy and archive mode. Possible values

- _None_ = All log files will be purged without being copied
- _CopyOnly_ = Simply copy log files to the RepositoryRootPath
- _CopyAndZip_ = Copy logfiles and send copied files to compressed archive
- _CopyZipAndDelete_ = Same as CopyAndZip, but delete copied log files from RepositoryRootPath

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

``` PowerShell
.\Purge-LogFiles.ps1 -DaysToKeep 7 -SendMail -MailFrom postmaster@sedna-inc.com -MailTo exchangeadmin@sedna-inc.com -MailServer mail.sedna-inc.com -UseDynamicExchangePaths -IncludeHttpErr
```

Delete Exchange Server, IIS, and HTTPERR log files older than 7 days, and send an HTML email. Identify Exchange file paths using AD configuration objects.

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

Stay connected:

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn:	[http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)

Additional Credits:

- Is-Admin function (c) by Michel de Rooij, michel@eightwone.com