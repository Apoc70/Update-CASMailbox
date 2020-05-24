# Update-CASMailbox.ps1

Enables or disabled Exchange CASMailbox protocols based on AD group membership

## Description

This script fetches the members of an AD security group and compares the list of members with a list of mailbox users having the requested CAS feature configured. All mailbox users not having the feature configured will get the CAS feature configured.

Available options:

- POP enabled/disabled
- IMAP enabled/disabled
- Outlook on the Web enabled/disabled
- ActiveSync enabled/disabled

Configuration actions are logged using the global functions library https://gallery.technet.microsoft.com/Centralized-logging-64e20f97

## Parameters

### GroupName

Name of Active Directory security group with mailbox user accounts to configure CAS mailbox settings

### POP

Switch to enable/disable POP3

### IMAP

Switch to enable/disable IMAP4

### OWA

Switch to enable/disabled Outlook on the web (aka Outlook Web Access)

### ActiveSync

Switch to enable/disable Exchange Server ActiveSync

### FeatureEnabled

Boolean attribute to enable or disable a CAS mailbox feature

## Examples

``` PowerShell
.\Update-CAS-Mailbox.ps1 -POP -FeatureEnabled $true -GroupName Exchange_POP_enabled
```

Enable POP3 for members of group Exchange_POP_enabled

``` PowerShell
.\Update-CAS-Mailbox.ps1 -OWA -FeatureEnabled $true -GroupName MyCompany_OWA_enabled -Verbose
```

Enable OWA for members of group MyCompany_OWA_enabled and getting verbose output

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)