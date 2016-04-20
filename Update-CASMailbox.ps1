<#
    .SYNOPSIS
    Enables or disabled Exchange CASMailbox protocols based on AD group membership

   
   	Thomas Stensitzki

	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.0, 2016-04-18

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    More information can be found at http://www.granikos.eu/en/scripts 

    .LINK
    Global functions library https://gallery.technet.microsoft.com/Centralized-logging-64e20f97
	
    .DESCRIPTION
    This script fetches the members of an AD security group and compares the list of members
    with a list of mailbox users having the requested CAS feature configured. All mailbox
    users not having the feature configured will get the CAS feature configured.

    Available options:
    - POP enabled/disabled
    - IMAP enabled/disabled
    - Outlook on the Web enabled/disabled
    - ActiveSync enabled/disabled

    Configuration actions are logged using the global functions library
	
    .NOTES 
    Requirements 
    - Windows Server 2012 or Windows Server 2012 R2  
    - Utilizes global functions library

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
	
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
   
	.EXAMPLE
    Enable POP3 for members of group Exchange_POP_enabled 
    Update-CAS-Mailbox.ps1 -POP -FeatureEnabled $true -GroupName Exchange_POP_enabled

    .EXAMPLE
    Enable OWA for members of group MyCompany_OWA_enabled and getting verbose output
    Update-CAS-Mailbox.ps1 -OWA -FeatureEnabled $true -GroupName MyCompany_OWA_enabled -Verbose
#>

Param(
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [string]$GroupName = "",
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [switch]$POP,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [switch]$IMAP,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [switch]$OWA,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [switch]$ActiveSync,
    [parameter(Mandatory=$true,ValueFromPipeline=$false)]
        [boolean]$FeatureEnabled
)

Set-StrictMode -Version Latest

Import-Module GlobalFunctions
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write("Script started")

$groupDN = ""

function Get-GroupMembers {
    Write-Verbose "Loading group members for $($groupDN)"
    $error.clear()
    $rootDomain = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().RootDomain.Name
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"GC://$rootDomain")
    $objSearcher.PageSize = 100
    $objSearcher.Filter = "(Memberof=$groupDN)"
    $newUsers = $objSearcher.FindAll()
    if($error) {
        $logger.Write("Error loading $($groupDN)", 1)
    }
    else {
        $tmp = ($newUsers | Measure-Object).Count
        Write-Verbose "$($tmp) members in $($groupDN)"
        $logger.Write("$($tmp) members in $($groupDN)")
    }

    $newUsers
}

function Get-FeatureEnabledUsers {
param (
    [parameter(Mandatory=$false)] [string]$FeatureFilter
)

    [hashtable]$featureUsersHash = @{}

    $expr = 'Get-CASMailbox -ResultSize unlimited | ?{' + $($FeatureFilter) + '}'

    Write-Verbose "Fetching feature enabled users: $($expr)"

    $featureUsers = Invoke-Expression $expr

    $tmp = ($featureUsers | Measure-Object).Count
    Write-Verbose "$($tmp) feature enabled users"
    $logger.Write("$($tmp) feature enabled users")

    if($tmp -ne 0) {
        foreach($featureUser in $featureUsers) {
            $featureUsersHash.Add($featureUser.DistinguishedName.ToLower(),$featureUser.PrimarySmtpAddress)
        }
    }

    if($True) {
        $output = $featureUsersHash.GetEnumerator() | foreach{ New-Object psobject -Property ([ordered]@{DN=$_.Name;SMTP=$_.Value})}
        $output | Export-Csv -Path D:\AutomatedServices\Exchange-Skripte\Update-CASMailbox\hash.txt -Encoding UTF8 -NoTypeInformation -Force
    }

    $featureUsersHash
}

function Set-CASFeature {
    param(
        [parameter(Mandatory=$false)] [string]$CasUserDN,
        [parameter(Mandatory=$false)] [string]$FeatureAttribute
    )

    # enabled CAS mailbox feature
    $Error.Clear()

    $expr = "Set-CASMailbox -Identity $($CasUserDN) $($FeatureAttribute) -ErrorAction Continue"

    Write-Verbose "Invoke: $($expr)"

    # Set-CASMailbox -Identity $CasUserDN $FeatureAttribute -ErrorAction Continue   

    if(!$Error) {
        Write-Verbose "Set feature [$($FeatureAttribute)]: $($CasUserDN)"
        $logger.Write("Set feature [$($FeatureAttribute)]: $($CasUserDN)")
    }
    else {
        Write-Error "Error setting feature [$($FeatureAttribute)]: $($CasUserDN)"
        $logger.Write("Setting feature failed [$($FeatureAttribute)]: $($CasUserDN)", 1)
        $Error.Clear()
    }
}

if($GroupName -ne "") {
    $groupDN = (Get-Group $GroupName).DistinguishedName
    if($groupDN -ne "") {
        Write-Verbose "$($GroupName) group found"
        $logger.Write("Group DN: $($groupDN)")

        # check CAS feature to updated
        if($POP) {
            # $filter = '{popenabled -eq $' + $($FeatureEnabled) + '}'
            $filter = '$_.PopEnabled -eq $' + $($FeatureEnabled)
            $featureAttributeString = '-POPEnabled:$' + $($FeatureEnabled)
        }
        elseif($IMAP) {
            # $filter = '{ImapEnabled -eq $' + $($FeatureEnabled) + '}'
            $filter = '$_.ImapEnabled -eq $' + $($FeatureEnabled)
            $featureAttributeString = '-ImapEnable:$' + $($FeatureEnabled)
        }
        elseif($OWA) {
            # $filter = '{OWAEnabled -eq $' + $($FeatureEnabled) + '}'
            $filter = '$_.OWAEnabled -eq $' + $($FeatureEnabled)
            $featureAttributeString = '-OWAEnabled:$' + $($FeatureEnabled)
        }       
        elseif($ActiveSync) {
            # $filter = '{ActiveSyncEnabled -eq $' + $($FeatureEnabled) + '}'
            $filter = '$_.ActiveSyncEnabled -eq $' + $($FeatureEnabled)
            $featureAttributeString = '-ActiveSyncEnabled:$' + $($FeatureEnabled)
        } 

        # Fetch Active Directory group members
        $currentGroupMembers = Get-GroupMembers

        # Check if AD group contains members
        if(($currentGroupMembers | Measure-Object).Count -gt 0) {

            # Fetch Exchange feature enabled users
            [hashtable]$currentFeatureUsers = Get-FeatureEnabledUsers -FeatureFilter $filter

            # Change feature settings
            $newfeatureUsers =  [System.Collections.ArrayList]@()

            foreach($groupMember in $currentGroupMembers) {
                $groupMemberDN = $groupMember.properties.distinguishedname[0].ToLower()
                if($currentFeatureUsers.ContainsKey($groupMemberDN)) {
                    Write-Verbose "$($groupMemberDN) already feature enabled"
                    $currentFeatureUsers.Remove($groupMemberDN)
                }
                else {
                    # add Active Directory group member to hashtable for enabling feature
                    Write-Verbose "$($groupMemberDN) not feature enabled, added to list"
                    $logger.Write("$($groupMemberDN) not feature enabled, added to list")
                    $newfeatureUsers.Add($groupMemberDN)
                }
            }

            if($newfeatureUsers.Count -ne 0) {
                Write-Verbose "$($newfeatureUsers.Count) new features users. Start enabling users."
                $logger.Write("$($newfeatureUsers.Count) new features users. Start enabling users.")
                foreach($newFeatureUser in $newfeatureUsers) {
                
                    Set-CASFeature -CasUserDN $newfeatureUser -FeatureAttribute $featureAttributeString
                }
            }
            else {
                Write-Verbose "$($newfeatureUsers.Count) new features users. Nothing to do."
                $logger.Write("$($newfeatureUsers.Count) new features users. Nothing to do.")
            }
        }
        else {
            Write-Output "$($groupDN) does not contain any members"
            $logger.Write("$($groupDN) does not contain any members")
        }
    }
    else {
        Write-Output "$($GroupName) group not found"
    }
}

$logger.Write("Script finished")