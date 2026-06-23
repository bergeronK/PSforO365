<#
.SYNOPSIS
	Imports the CSV produced by Export-ADGroupsAndMembers.ps1 into a target Active Directory
	forest/domain: creates any missing groups, then adds their members.

.DESCRIPTION
	Reads the export CSV (one row per group+member, with empty-group placeholder rows) and:
		1. Creates each group that doesn't already exist in the target (matched by
		   GroupSamAccountName), placing it under an OU that mirrors GroupOUPath from the
		   source forest. Existing groups are left alone except their Description, which is
		   refreshed to match the source.
		2. Resolves GroupManagedBy (if present) against the target forest and sets it.
		3. Adds each member to its group by looking up MemberSamAccountName in the target
		   forest. Rows where MemberResolved was False in the export (e.g. unresolved foreign
		   security principals) are skipped and reported instead of guessed at.

	Run with -WhatIf first to preview every group/OU/membership change before committing.

.NOTES
	Author		: Claude Code
	File Name	: Import-ADGroupsAndMembers.ps1
	Requires	: RSAT "Active Directory module for Windows PowerShell" and write access
				  (create group/OU, modify group membership) in the target forest.
	Pairs with	: Export-ADGroupsAndMembers.ps1

.EXAMPLE
	.\Import-ADGroupsAndMembers.ps1 -CsvPath C:\Temp\ADGroupExport.csv -WhatIf

	Previews what would be created/changed in the current domain without making any changes.

.EXAMPLE
	.\Import-ADGroupsAndMembers.ps1 -CsvPath C:\Temp\ADGroupExport.csv -CreateMissingOUs -Server dc01.fabrikam.com -Credential (Get-Credential)

	Imports into the fabrikam.com forest, recreating any missing OU structure and
	authenticating with alternate credentials.

.EXAMPLE
	.\Import-ADGroupsAndMembers.ps1 -CsvPath C:\Temp\ADGroupExport.csv -TargetOUFallback "OU=MigratedGroups,DC=fabrikam,DC=com" -UnresolvedLogPath C:\Temp\UnresolvedMembers.csv

	Imports groups whose source OU can't be matched into a single fallback OU instead, and
	logs every member that couldn't be added so they can be handled manually.

.PARAMETER CsvPath
	Path to the CSV produced by Export-ADGroupsAndMembers.ps1.

.PARAMETER TargetOUFallback
	DistinguishedName of an OU to place groups in when their source GroupOUPath is empty, or
	doesn't exist in the target and -CreateMissingOUs was not specified. Defaults to the
	target domain root.

.PARAMETER CreateMissingOUs
	Switch. When set, recreates the source OU path under the target domain if it doesn't
	already exist there.

.PARAMETER Server
	Optional domain controller (or domain) in the target forest to write to. Defaults to the
	current logon domain.

.PARAMETER Credential
	Optional alternate credential used for all AD writes (e.g. when importing into a forest
	you are not logged into).

.PARAMETER UnresolvedLogPath
	Optional CSV path. When set, every member row that was skipped (unresolved in the export,
	or not found in the target forest) is written here for manual follow-up.
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param (
	[Parameter(Mandatory=$true)]
	[ValidateScript({Test-Path $_ -PathType Leaf})]
	[string] $CsvPath,

	[Parameter(Mandatory=$false)]
	[string] $TargetOUFallback,

	[Parameter(Mandatory=$false)]
	[switch] $CreateMissingOUs,

	[Parameter(Mandatory=$false)]
	[string] $Server,

	[Parameter(Mandatory=$false)]
	[System.Management.Automation.PSCredential] $Credential,

	[Parameter(Mandatory=$false)]
	[ValidateScript({Test-Path (Split-Path $_ -Parent) -PathType Container})]
	[string] $UnresolvedLogPath
)

#Region Variables
####################################################
# Variables
####################################################

$Stats = [ordered]@{
	GroupsCreated           = 0
	GroupsExisting          = 0
	GroupsFailed            = 0
	MembersAdded            = 0
	MembersAlreadyPresent   = 0
	MembersSkippedUnresolved = 0
	MembersFailed           = 0
}

$UnresolvedRows = [System.Collections.Generic.List[object]]::new()
$GroupDNMap     = @{}

#EndRegion

#Region Functions
####################################################
# Functions
####################################################

#---------------------------------------------------
# Builds a parameter hashtable with -Server/-Credential
# only included when they were actually supplied.
#---------------------------------------------------
function New-ADParamSplat
{
	param (
		[string] $Server,
		[System.Management.Automation.PSCredential] $Credential,
		[hashtable] $Additional
	)

	$Splat = @{}
	if ($Server)     { $Splat.Server     = $Server }
	if ($Credential) { $Splat.Credential = $Credential }
	if ($Additional)
	{
		foreach ($Key in $Additional.Keys) { $Splat[$Key] = $Additional[$Key] }
	}
	return $Splat
}

#---------------------------------------------------
# Resolves (and optionally creates) the target OU that a
# group's relative OUPath from the export maps to, falling
# back to -TargetOUFallback or the domain root when the path
# is empty or missing and -CreateMissingOUs was not set.
#---------------------------------------------------
function Resolve-TargetGroupPath
{
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (
		[string] $RelativeOUPath,
		[Parameter(Mandatory=$true)] [string] $DomainDN,
		[string] $Server,
		[System.Management.Automation.PSCredential] $Credential,
		[switch] $CreateMissingOUs,
		[string] $TargetOUFallback
	)

	$FallbackPath = if ($TargetOUFallback) { $TargetOUFallback } else { $DomainDN }

	if (-not $RelativeOUPath) { return $FallbackPath }

	$FullPath = "$RelativeOUPath,$DomainDN"

	$CheckSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{ Identity = $FullPath; ErrorAction = 'Stop' }
	try
	{
		Get-ADOrganizationalUnit @CheckSplat | Out-Null
		return $FullPath
	}
	catch
	{
		if (-not $CreateMissingOUs)
		{
			Write-Warning "OU path '$FullPath' was not found in the target and -CreateMissingOUs was not specified. Falling back to '$FallbackPath'."
			return $FallbackPath
		}

		$Segments = $RelativeOUPath -split '(?<!\\),'
		[array]::Reverse($Segments)

		$CurrentParent = $DomainDN
		foreach ($Segment in $Segments)
		{
			$Name      = $Segment -replace '^OU=', ''
			$CurrentDN = "$Segment,$CurrentParent"

			$ExistSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{ Identity = $CurrentDN; ErrorAction = 'SilentlyContinue' }
			$Existing = Get-ADOrganizationalUnit @ExistSplat

			if (-not $Existing -and $PSCmdlet.ShouldProcess($CurrentDN, 'Create organizational unit'))
			{
				$NewSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{ Name = $Name; Path = $CurrentParent; ErrorAction = 'Stop' }
				New-ADOrganizationalUnit @NewSplat
			}

			$CurrentParent = $CurrentDN
		}

		return $FullPath
	}
}

#---------------------------------------------------
# Creates the group in the target forest if it doesn't
# already exist (matched on SamAccountName), or refreshes
# its Description if it does. Returns the group's DN (or
# $null on failure) so membership can be added afterwards.
#---------------------------------------------------
function Import-TargetGroup
{
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (
		[Parameter(Mandatory=$true)] [object] $GroupDef,
		[Parameter(Mandatory=$true)] [string] $TargetPath,
		[string] $Server,
		[System.Management.Automation.PSCredential] $Credential
	)

	$FindSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{
		Filter      = "SamAccountName -eq '$($GroupDef.GroupSamAccountName)'"
		Properties  = @('Description')
		ErrorAction = 'Stop'
	}
	$Existing = Get-ADGroup @FindSplat

	if ($Existing)
	{
		$Stats.GroupsExisting++

		if ($GroupDef.GroupDescription -and ($Existing.Description -ne $GroupDef.GroupDescription))
		{
			if ($PSCmdlet.ShouldProcess($GroupDef.GroupSamAccountName, 'Update group description'))
			{
				$SetSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{
					Identity    = $Existing.DistinguishedName
					Description = $GroupDef.GroupDescription
					ErrorAction = 'Stop'
				}
				Set-ADGroup @SetSplat
			}
		}

		return $Existing.DistinguishedName
	}

	$NewGroupParams = @{
		Name           = $GroupDef.GroupName
		SamAccountName = $GroupDef.GroupSamAccountName
		GroupCategory  = $GroupDef.GroupCategory
		GroupScope     = $GroupDef.GroupScope
		Path           = $TargetPath
		ErrorAction    = 'Stop'
	}
	if ($GroupDef.GroupDescription) { $NewGroupParams.Description = $GroupDef.GroupDescription }

	$NewSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional $NewGroupParams
	$ExpectedDN = "CN=$($GroupDef.GroupName -replace ',', '\,'),$TargetPath"

	if ($PSCmdlet.ShouldProcess($GroupDef.GroupSamAccountName, "Create group under '$TargetPath'"))
	{
		try
		{
			New-ADGroup @NewSplat
			$Stats.GroupsCreated++
			return $ExpectedDN
		}
		catch
		{
			Write-Warning "Failed to create group '$($GroupDef.GroupSamAccountName)': $($_.Exception.Message)"
			$Stats.GroupsFailed++
			return $null
		}
	}

	# -WhatIf: report what would happen and return the DN it would have so
	# membership additions can still be previewed.
	return $ExpectedDN
}

#---------------------------------------------------
# Resolves -ManagedBy against the target forest by
# SamAccountName and applies it to the group, if found.
#---------------------------------------------------
function Set-TargetGroupManager
{
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (
		[Parameter(Mandatory=$true)] [string] $GroupDN,
		[Parameter(Mandatory=$true)] [string] $ManagerSamAccountName,
		[string] $Server,
		[System.Management.Automation.PSCredential] $Credential
	)

	$FindSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{
		Filter      = "SamAccountName -eq '$ManagerSamAccountName'"
		ErrorAction = 'SilentlyContinue'
	}
	$Manager = Get-ADObject @FindSplat

	if (-not $Manager)
	{
		Write-Warning "ManagedBy account '$ManagerSamAccountName' was not found in the target forest - leaving ManagedBy unset."
		return
	}

	if ($PSCmdlet.ShouldProcess($GroupDN, "Set ManagedBy to '$ManagerSamAccountName'"))
	{
		$SetSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{
			Identity    = $GroupDN
			ManagedBy   = $Manager.DistinguishedName
			ErrorAction = 'Stop'
		}
		Set-ADGroup @SetSplat
	}
}

#---------------------------------------------------
# Resolves a member by SamAccountName in the target forest
# and adds it to the given group. Tracks stats and returns
# whether the member is (now) in the group.
#---------------------------------------------------
function Add-TargetGroupMember
{
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (
		[Parameter(Mandatory=$true)] [string] $GroupSamAccountName,
		[Parameter(Mandatory=$true)] [string] $GroupDN,
		[Parameter(Mandatory=$true)] [string] $MemberSamAccountName,
		[string] $Server,
		[System.Management.Automation.PSCredential] $Credential
	)

	$FindSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{
		Filter      = "SamAccountName -eq '$MemberSamAccountName'"
		ErrorAction = 'SilentlyContinue'
	}
	$TargetMember = Get-ADObject @FindSplat

	if (-not $TargetMember)
	{
		$Stats.MembersSkippedUnresolved++
		return $false
	}

	if (-not $PSCmdlet.ShouldProcess("$MemberSamAccountName -> $GroupSamAccountName", 'Add group member')) { return $true }

	try
	{
		$AddSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{
			Identity    = $GroupDN
			Members     = $TargetMember.DistinguishedName
			ErrorAction = 'Stop'
		}
		Add-ADGroupMember @AddSplat
		$Stats.MembersAdded++
		return $true
	}
	catch
	{
		if ($_.Exception.Message -match 'already a member')
		{
			$Stats.MembersAlreadyPresent++
			return $true
		}

		Write-Warning "Failed to add '$MemberSamAccountName' to '$GroupSamAccountName': $($_.Exception.Message)"
		$Stats.MembersFailed++
		return $false
	}
}

#EndRegion

#Region Main
####################################################
# Main
####################################################

try
{
	Import-Module ActiveDirectory -ErrorAction Stop
}
catch
{
	throw "The ActiveDirectory PowerShell module is required (RSAT). $($_.Exception.Message)"
}

$DomainSplat   = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{ ErrorAction = 'Stop' }
$TargetDomainDN = (Get-ADDomain @DomainSplat).DistinguishedName

$Rows = Import-Csv -Path $CsvPath
$GroupedRows = $Rows | Group-Object -Property GroupSamAccountName

Write-Host "=> Importing $($GroupedRows.Count) group(s) into target domain (root: $TargetDomainDN)..."

foreach ($GroupRows in $GroupedRows)
{
	$GroupDef = $GroupRows.Group[0]

	$TargetPath = Resolve-TargetGroupPath -RelativeOUPath $GroupDef.GroupOUPath -DomainDN $TargetDomainDN `
		-Server $Server -Credential $Credential -CreateMissingOUs:$CreateMissingOUs -TargetOUFallback $TargetOUFallback

	Write-Host "`t- Group '$($GroupDef.GroupSamAccountName)' -> '$TargetPath'"

	$GroupDN = Import-TargetGroup -GroupDef $GroupDef -TargetPath $TargetPath -Server $Server -Credential $Credential
	if (-not $GroupDN) { continue }

	$GroupDNMap[$GroupDef.GroupSamAccountName] = $GroupDN

	if ($GroupDef.GroupManagedBy)
	{
		Set-TargetGroupManager -GroupDN $GroupDN -ManagerSamAccountName $GroupDef.GroupManagedBy -Server $Server -Credential $Credential
	}
}

Write-Host "=> Adding group members..."

foreach ($Row in $Rows)
{
	if (-not $Row.MemberSamAccountName) { continue }

	if ($Row.MemberResolved -ne 'True')
	{
		$Stats.MembersSkippedUnresolved++
		$UnresolvedRows.Add($Row)
		continue
	}

	$GroupDN = $GroupDNMap[$Row.GroupSamAccountName]
	if (-not $GroupDN) { continue }

	$Added = Add-TargetGroupMember -GroupSamAccountName $Row.GroupSamAccountName -GroupDN $GroupDN `
		-MemberSamAccountName $Row.MemberSamAccountName -Server $Server -Credential $Credential

	if (-not $Added) { $UnresolvedRows.Add($Row) }
}

if ($UnresolvedLogPath -and $UnresolvedRows.Count -gt 0)
{
	$UnresolvedRows | Export-Csv -Path $UnresolvedLogPath -NoTypeInformation -Encoding UTF8
}

Write-Host ""
Write-Host "Import complete."
Write-Host "`tGroups created            : $($Stats.GroupsCreated)"
Write-Host "`tGroups already existing   : $($Stats.GroupsExisting)"
Write-Host "`tGroups failed             : $($Stats.GroupsFailed)"
Write-Host "`tMembers added             : $($Stats.MembersAdded)"
Write-Host "`tMembers already present   : $($Stats.MembersAlreadyPresent)"
Write-Host "`tMembers skipped/unresolved: $($Stats.MembersSkippedUnresolved)"
Write-Host "`tMembers failed            : $($Stats.MembersFailed)"
if ($UnresolvedLogPath -and $UnresolvedRows.Count -gt 0)
{
	Write-Host "`tUnresolved/failed member rows logged to: $UnresolvedLogPath"
}

#EndRegion
