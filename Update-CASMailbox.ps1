<#
    .SYNOPSIS
    NOT YET READY
   
   	Thomas Stensitzki

	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.0, 2016-04-15

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    More information can be found at http://www.granikos.eu/en/scripts 
	
    .DESCRIPTION
	
    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
    - Utilizes global functions library

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
	
	.PARAMETER 
   
	.EXAMPLE

#>

Param(
    [parameter(Mandatory=$false,ValueFromPipeline=$false,HelpMessage='SMTP Server address for sending result summary')]
        [string]$GroupName = "bkmail_POP",
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

Import-Module BDRFunctions
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write("Script started")

$groupDN = ""

function Get-GroupMembers {
    Write-Verbose "Loading group members for $($groupDN)"
    $error.clear()
    $rootDomain = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().RootDomain.Name
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"GC://$root")
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

    Write-Verbose "Fetching feature using: $($expr)"

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
        Write-Verbose "Feature enabled: $($CasUserDN)"
        $logger.Write("Feature enabled: $($CasUserDN)")
    }
    else {
        Write-Error "Error feature enabling: $($CasUserDN)"
        $logger.Write("Feature enable failed: $($CasUserDN)", 1)
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
        Write-Verbose "$($GroupName) group not found"
    }
}

$logger.Write("Script finished")