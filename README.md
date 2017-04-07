# Purge-LogFiles.ps1
PowerShell script for Exchange Server 2013 environments to clean up Exchange and IIS log files.

## Description
This script deletes all Exchange and IIS logs older than X days from all Exchange 2013 servers that are fetched using the Get-ExchangeServer cmdlet.

The Exchange log file location is read from the environment variable and used to build an adminstrative UNC path for file deletions.

It is assumed that the Exchange setup path is IDENTICAL across all Exchange servers.

The IIS log file location is read from the local IIS metabase of the LOCAL server and is used to build an administrative UNC path for IIS log file deletions.

 
It is assumed that the IIS log file location is identical across all Exchange servers

## Requirements

- Exchange Server 2013+
- Exchange Management Shell (EMS)

## Parameters
### DaysToKeep
Number of days Exchange and IIS log files should be retained, default is 30 days

### Auto
Switch to use automatic detection of the IIS and Exchange log folder paths

### SendMail
Switch to send an Html report

### MailFrom
Email address of report sender

### MailTo
Email address of report recipient

### MailServer
SMTP Server for email report


## Examples
```
.\Purge-LogFiles -DaysToKeep 14
```
Delete Exchange and IIS log files older than 14 days

```
.\Purge-LogFiles -DaysToKeep 7 -Auto
```
Delete Exchange and IIS log files older than 7 days with automatic discovery

```
.\Purge-LogFiles -DaysToKeep 7 -Auto -SendMail -MailFrom postmaster@sedna-inc.com -MailTo exchangeadmin@sedna-inc.com -MailServer mail.sedna-inc.com 
```
Delete Exchange and IIS log files older than 7 days with automatic discovery and send email report

## TechNet Gallery
Find the script at TechNet Gallery
* https://gallery.technet.microsoft.com/Purge-Exchange-Server-2013-c2e03e72


## Credits
Written by: Thomas Stensitzki

## Social

* My Blog: https://www.granikos.eu/en/justcantgetenough
* Archived Blog:	http://www.sf-tools.net/
* Twitter:	https://twitter.com/apoc70
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github:	https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog:     http://blog.granikos.eu/
* Website:	https://www.granikos.eu/en/
* Twitter:	https://twitter.com/granikos_de

Additional Credits:
* Is-Admin function (c) by Michel de Rooij, michel@eightwone.com
