# Update-CASMailbox.ps1
Enables or disabled Exchange CASMailbox protocols based on AD group membership

##Description
This script fetches the members of an AD security group and compares the list of members with a list of mailbox users having the requested CAS feature configured. All mailbox users not having the feature configured will get the CAS feature configured.

Available options:
- POP enabled/disabled
- IMAP enabled/disabled
- Outlook on the Web enabled/disabled
- ActiveSync enabled/disabled

Configuration actions are logged using the global functions library https://gallery.technet.microsoft.com/Centralized-logging-64e20f97

##Inputs
.PARAMETER GroupName
Name of Active Directory security group with mailbox user accounts to configure CAS mailbox settings

.PARAMETER POP
Switch to enable/disable POP3

.PARAMETER IMAP
Switch to enable/disable IMAP4

.PARAMETER OWA
Switch to enable/disabled Outlook on the web (aka Outlook Web Access)

.PARAMETER ActiveSync
Switch to enable/disable Exchange Server ActiveSync

.PARAMETER FeatureEnabled
Boolean attribute to enable or disable a CAS mailbox feature

##Outputs
All outputs (cmdlet results) are written to a log file. A new log file is generated daily.

You can use the -Verbose switch to get more output to the command line.

##Examples
```
.\Update-CAS-Mailbox.ps1 -POP -FeatureEnabled $true -GroupName Exchange_POP_enabled
```
Enable POP3 for members of group Exchange_POP_enabled

```
.\Update-CAS-Mailbox.ps1 -OWA -FeatureEnabled $true -GroupName MyCompany_OWA_enabled - Verbose
```
Enable OWA for members of group MyCompany_OWA_enabled and getting verbose output

##TechNet Gallery
Find the script at TechNet Gallery
* 

##Credits
Written by: Thomas Stensitzki

Stay connected:

* My Blog: http://justcantgetenough.granikos.eu
* Archived Blog: http://www.sf-tools.net/
* Twitter:	https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github:	https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog:     http://blog.granikos.eu/
* Website:	https://www.granikos.eu/en/
* Twitter:	https://twitter.com/granikos_de