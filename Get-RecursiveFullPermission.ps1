<#	
	.NOTES
	===========================================================================
	 Created on:   	9-Aug-2018
	 Created by:   	Tito D. Castillote Jr.
					june.castillote@gmail.com
	 Filename:     	Get-RecursiveFullPermission.ps1
	 Version:		1.0 (9-Aug-2018)
	===========================================================================

	.LINK
		https://www.lazyexchangeadmin.com/2018/08/recursive-mailbox-full-permission.html

	.SYNOPSIS
		Get-RecursiveFullPermission.ps1 to extract the Full Mailbox Permission including Groups and Nested Groups

	.DESCRIPTION
		IMPORTANT: 	DO NOT USE WITH EXCHANGE MANAGEMENT SHELL OR REMOTE POWERSHELL.		
					USE ONLY WITH NORMAL POWERSHELL WITH EXCHANGE MANAGEMENT TOOLS
					THE SCRIPT WILL AUTOMATICALLY IMPORT THE EXCHANGE SNAPIN
	.PARAMETER

	.EXAMPLE
		Get-RecursiveFullPermission.ps1
#>

#>>Import Exchange 2010 Shell Snap-In if not already added-----------------------

	if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
	{
		try
		{
			Write-Host 'Add Exchange Snap-in' -ForegroundColor Green
			Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
		}
		catch
		{
			Write-Warning $_.Exception.Message
			EXIT
		}
	}

#>>------------------------------------------------------------------------------


#Function to recursively list group members (nested)
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$today = '{0:dd-MMM-yyyy_hh-mm-tt}' -f (Get-Date)
Start-Transcript -Path "$($script_root)\$($today)_debugLog.txt" -Append

Function Get-MembersRecursive ($groupName)
{
    $groupMembers = @()
    $groupName = Get-Group $groupName -ErrorAction SilentlyContinue
    foreach ($groupMember in $groupName.Members)
    {
        if (Get-Group $groupMember -ErrorAction SilentlyContinue)
        {
            $groupMembers += Get-MembersRecursive $groupMember
        } else {
            $groupMembers += ((get-user $groupMember.Name).UserPrincipalName)
        }
    }
    $groupMembers = $groupMembers | Select -Unique
    return $groupMembers
}

#Start Script
$mailboxFile = "$($script_root)\mailboxlist.txt"
$reportFile = "$($script_root)\report.csv"

Write-Host "Importing Mailbox List" -ForegroundColor Green

#================ OPTIONS
#Uncomment which option you wish to use, and comment out the ones that will not be used

#OPTION 1: If you want to process only a specific list of mailbox from the mailboxlist.txt
$mailboxList = Get-Content $mailboxFile

#OPTION 2: If you want to process ALL mailboxes
#$mailboxList = Get-Mailbox -resultsize Unlimited -RecipientTypeDetails UserMailbox

#OPTION 3: If you want to process just a single mailbox
#$mailboxList = Get-Mailbox User1
#================ OPTIONS

Write-Host "Total Number of Mailbox to Process: $($mailboxList.count)" -ForegroundColor Green

$finalReport = @()
foreach ($mailbox in $mailboxList)
{
Write-Host "Mailbox: $($mailbox)" -ForegroundColor Yellow
$mailboxPermissions = Get-MailboxPermission $mailbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.user.tostring() -notlike "S-*" -and $_.IsInherited -eq $false -and $_.AccessRights -match "FullAccess"}
$mailboxDetail = Get-Recipient $mailbox
	if ($mailboxPermissions.count -gt 0)
	{
		Write-Host "Access List: " -ForegroundColor Cyan
		foreach ($mailboxPermission in $mailboxPermissions)
		{
			
			$x = Get-Recipient $mailboxPermission.User
			
			#if the WhoHasAccessName is a group, recursively extract members
			if ($x.RecipientType -match 'group')
			{
				#Call function to recurse the group
				$members = Get-MembersRecursive $x.Identity
				
				#if the function returned a non ZERO result
				if ($members.count -gt 0)
				{
					#Write-Host "Access List: " -ForegroundColor Cyan -NoNewLine
					foreach ($member in $members)
					{
						Write-Host "     $($member)" -ForegroundColor Cyan
						$temp = "" | Select MailboxName,MailboxEmailAddress,WhoHasAccessName,WhoHasAccessEmailAddress,AccessType,ParentGroupName,ParentGroupEmailAddress
						$y = Get-Recipient $member
						$temp.MailboxName = $mailboxPermission.Identity.ToString().Split("/")[-1]
						$temp.MailboxEmailAddress = $mailboxDetail.PrimarySMTPAddress
						$temp.WhoHasAccessName = $y.Identity.ToString().Split("/")[-1]
						$temp.WhoHasAccessEmailAddress = $y.PrimarySMTPAddress
						$temp.AccessType = "InheritedFromGroup"
						$temp.ParentGroupName = $x.Name
						$temp.ParentGroupEmailAddress = $x.PrimarySMTPAddress
						$finalReport += $temp
					}				
				}				
			}
			else
			{				
				Write-Host "     $($x.PrimarySMTPAddress)" -ForegroundColor Cyan
				$temp = "" | Select MailboxName,MailboxEmailAddress,WhoHasAccessName,WhoHasAccessEmailAddress,AccessType,ParentGroupName,ParentGroupEmailAddress
				$temp.MailboxName = $mailboxPermission.Identity.ToString().Split("/")[-1]
				$temp.MailboxEmailAddress = $mailboxDetail.PrimarySMTPAddress
				$temp.WhoHasAccessName = $mailboxPermission.User.ToString().Split("\")[-1]
				$temp.WhoHasAccessEmailAddress = $x.PrimarySMTPAddress
				$temp.AccessType = "DirectUser"
				$temp.ParentGroupName = ""
				$temp.ParentGroupEmailAddress = ""
				$finalReport += $temp
			}
		}
	}
	else
	{
		Write-Host " -> Skipped" -ForegroundColor Yellow
	}
}
$finalReport | export-csv -nti $reportFile
Write-Host "Process Completed. Please see report for details - " -ForegroundColor Green -NoNewLine
Write-Host $reportFile -ForegroundColor Yellow
Stop-Transcript