<#
.SYNOPSIS
	Exports Active Directory security groups (and their members) from a source forest/domain
	to a CSV file that is structured to be re-imported into a different target AD forest.

.DESCRIPTION
	For every group matched by -GroupCategory (Security by default), the script records the
	group's identity/attributes plus one row per member. Membership is captured by
	SamAccountName/UserPrincipalName/SID rather than DistinguishedName, since DNs from the
	source forest are meaningless in the target forest. This lets an import routine:
		1. Create/update each group (deduplicated on GroupSamAccountName).
		2. Resolve each MemberSamAccountName (or MemberUPN) against the target forest and add
		   it to the corresponding group.

	Members that cannot be resolved to an AD object in the source forest (most commonly
	ForeignSecurityPrincipals representing members from a trusted domain outside the forest
	being queried) are still written out with MemberResolved = False and whatever SID could
	be parsed from the member's DN, so they can be reviewed/handled manually instead of
	silently dropped.

	Groups with no members are written out as a single row with empty Member* fields so they
	are still created on import.

.NOTES
	Author		: Claude Code
	File Name	: Export-ADGroupsAndMembers.ps1
	Requires	: RSAT "Active Directory module for Windows PowerShell" and read access to the
				  source forest.
	Limitations	: Group membership is read directly from the group's 'member' attribute
				  (not recursively expanded), so nested groups are preserved as their own
				  group rows rather than flattened. Very large groups (thousands of members)
				  may require attribute range retrieval, which is not implemented here.

.EXAMPLE
	.\Export-ADGroupsAndMembers.ps1 -OutputPath C:\Temp\ADGroupExport.csv

	Exports all Security groups from the current domain.

.EXAMPLE
	.\Export-ADGroupsAndMembers.ps1 -OutputPath C:\Temp\ADGroupExport.csv -AllDomainsInForest -Server dc01.contoso.com

	Exports all Security groups from every domain in the forest reachable via dc01.contoso.com.

.EXAMPLE
	.\Export-ADGroupsAndMembers.ps1 -OutputPath C:\Temp\ADGroupExport.csv -SearchBase "OU=Groups,DC=contoso,DC=com" -GroupCategory All -Credential (Get-Credential)

	Exports both Security and Distribution groups under a specific OU, authenticating with
	alternate credentials (useful when running against a forest you don't have a logon in).

.PARAMETER OutputPath
	File path for the CSV export.

.PARAMETER SearchBase
	Optional DistinguishedName to limit the export to groups under a specific OU/container.

.PARAMETER GroupCategory
	Which groups to export: Security (default), Distribution, or All.

.PARAMETER Server
	Optional domain controller (or domain) to query. Defaults to the current logon domain.

.PARAMETER Credential
	Optional alternate credential used for all AD queries (e.g. when exporting from a forest
	you are not logged into).

.PARAMETER AllDomainsInForest
	Switch. When set, discovers every domain in the forest reachable from -Server (or the
	current forest if -Server is omitted) and exports groups from each of them.
#>

#Requires -Version 5.1

[CmdletBinding()]
param (
	[Parameter(Mandatory=$true)]
	[ValidateScript({Test-Path (Split-Path $_ -Parent) -PathType Container})]
	[string] $OutputPath,

	[Parameter(Mandatory=$false)]
	[string] $SearchBase,

	[Parameter(Mandatory=$false)]
	[ValidateSet('Security','Distribution','All')]
	[string] $GroupCategory = 'Security',

	[Parameter(Mandatory=$false)]
	[string] $Server,

	[Parameter(Mandatory=$false)]
	[System.Management.Automation.PSCredential] $Credential,

	[Parameter(Mandatory=$false)]
	[switch] $AllDomainsInForest
)

#Region Variables
####################################################
# Variables
####################################################

$ExportRows = [System.Collections.Generic.List[object]]::new()
$Stats = [ordered]@{
	GroupsExported      = 0
	MemberRowsExported  = 0
	UnresolvedMembers   = 0
}

#EndRegion

#Region Functions
####################################################
# Functions
####################################################

#---------------------------------------------------
# Builds a parameter hashtable with -Server/-Credential
# only included when they were actually supplied, so we
# don't pass empty values to AD cmdlets.
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
# Strips the leaf (CN=...) and domain (DC=...) components
# from a DistinguishedName, leaving just the OU path so it
# can be matched against (or recreated under) an equivalent
# OU structure in the target forest.
#---------------------------------------------------
function Get-RelativeOUPath
{
	param (
		[Parameter(Mandatory=$true)]
		[string] $DistinguishedName
	)

	$Segments  = $DistinguishedName -split '(?<!\\),'
	$OUSegments = $Segments | Select-Object -Skip 1 | Where-Object { $_ -notmatch '^DC=' }
	return ($OUSegments -join ',')
}

#---------------------------------------------------
# Resolves a member/manager DistinguishedName to the
# identity fields needed for a cross-forest import. Falls
# back to a parsed SID (e.g. ForeignSecurityPrincipal) when
# the object can't be read directly.
#---------------------------------------------------
function Resolve-ADMemberObject
{
	param (
		[Parameter(Mandatory=$true)]
		[string] $MemberDN,

		[string] $Server,

		[System.Management.Automation.PSCredential] $Credential
	)

	$Splat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{
		Identity     = $MemberDN
		Properties   = @('SamAccountName','objectClass','userPrincipalName','objectSid','CanonicalName')
		ErrorAction  = 'Stop'
	}

	try
	{
		$Obj = Get-ADObject @Splat

		return [PSCustomObject]@{
			SamAccountName = $Obj.SamAccountName
			Name           = $Obj.Name
			ObjectClass    = $Obj.objectClass
			UPN            = $Obj.userPrincipalName
			SID            = if ($Obj.objectSid) { $Obj.objectSid.Value } else { $null }
			Domain         = if ($Obj.CanonicalName) { ($Obj.CanonicalName -split '/')[0] } else { $null }
			Resolved       = $true
		}
	}
	catch
	{
		$ParsedSID = $null
		if ($MemberDN -match 'CN=(S-1-[\d-]+),') { $ParsedSID = $Matches[1] }

		Write-Warning "Could not resolve member '$MemberDN' - recording as unresolved (likely a foreign security principal). $($_.Exception.Message)"

		return [PSCustomObject]@{
			SamAccountName = $null
			Name           = $MemberDN
			ObjectClass    = 'foreignSecurityPrincipal'
			UPN            = $null
			SID            = $ParsedSID
			Domain         = $null
			Resolved       = $false
		}
	}
}

#---------------------------------------------------
# Exports security groups (and members) found via the
# given -Server (a specific DC/domain), appending rows
# to the shared $ExportRows list.
#---------------------------------------------------
function Export-GroupsFromDomain
{
	param (
		[string] $Server,
		[string] $SearchBase,
		[string] $GroupCategory,
		[System.Management.Automation.PSCredential] $Credential
	)

	$Filter = switch ($GroupCategory)
	{
		'Security'     { "GroupCategory -eq 'Security'" }
		'Distribution'  { "GroupCategory -eq 'Distribution'" }
		'All'           { '*' }
	}

	$GroupSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{
		Filter      = $Filter
		Properties  = @('Description','ManagedBy','member','GroupScope','GroupCategory','DistinguishedName','SamAccountName','Name')
		ErrorAction = 'Stop'
	}
	if ($SearchBase) { $GroupSplat.SearchBase = $SearchBase }

	$ServerLabel = if ($Server) { $Server } else { '(current domain)' }
	Write-Host "=> Querying groups on '$ServerLabel'..."
	$Groups = Get-ADGroup @GroupSplat

	foreach ($Group in $Groups)
	{
		Write-Host "`t- Exporting group '$($Group.SamAccountName)' ($($Group.member.Count) member(s))..."
		$Stats.GroupsExported++

		$ManagedBySam = $null
		if ($Group.ManagedBy)
		{
			$ManagedBySam = (Resolve-ADMemberObject -MemberDN $Group.ManagedBy -Server $Server -Credential $Credential).SamAccountName
		}

		$GroupOUPath = Get-RelativeOUPath -DistinguishedName $Group.DistinguishedName

		if (-not $Group.member -or $Group.member.Count -eq 0)
		{
			$Row = [PSCustomObject][ordered]@{
				GroupSamAccountName  = $Group.SamAccountName
				GroupName            = $Group.Name
				GroupDescription     = $Group.Description
				GroupCategory        = $Group.GroupCategory
				GroupScope           = $Group.GroupScope
				GroupOUPath          = $GroupOUPath
				GroupManagedBy       = $ManagedBySam
				MemberSamAccountName = $null
				MemberName           = $null
				MemberType           = $null
				MemberUPN            = $null
				MemberSID            = $null
				MemberResolved       = $null
				MemberDomain         = $null
			}
			$ExportRows.Add($Row)
			continue
		}

		foreach ($MemberDN in $Group.member)
		{
			$Member = Resolve-ADMemberObject -MemberDN $MemberDN -Server $Server -Credential $Credential

			$Row = [PSCustomObject][ordered]@{
				GroupSamAccountName  = $Group.SamAccountName
				GroupName            = $Group.Name
				GroupDescription     = $Group.Description
				GroupCategory        = $Group.GroupCategory
				GroupScope           = $Group.GroupScope
				GroupOUPath          = $GroupOUPath
				GroupManagedBy       = $ManagedBySam
				MemberSamAccountName = $Member.SamAccountName
				MemberName           = $Member.Name
				MemberType           = $Member.ObjectClass
				MemberUPN            = $Member.UPN
				MemberSID            = $Member.SID
				MemberResolved       = $Member.Resolved
				MemberDomain         = $Member.Domain
			}
			$ExportRows.Add($Row)

			$Stats.MemberRowsExported++
			if (-not $Member.Resolved) { $Stats.UnresolvedMembers++ }
		}
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

$ServersToQuery = @()

if ($AllDomainsInForest)
{
	$ForestSplat = New-ADParamSplat -Server $Server -Credential $Credential -Additional @{ ErrorAction = 'Stop' }
	$Forest = Get-ADForest @ForestSplat

	foreach ($DomainName in $Forest.Domains)
	{
		try
		{
			$DC = Get-ADDomainController -DomainName $DomainName -Discover -ErrorAction Stop
			$ServersToQuery += $DC.HostName[0]
		}
		catch
		{
			Write-Warning "Could not discover a domain controller for '$DomainName' - skipping. $($_.Exception.Message)"
		}
	}
}
else
{
	$ServersToQuery = @($Server)
}

foreach ($QueryServer in $ServersToQuery)
{
	Export-GroupsFromDomain -Server $QueryServer -SearchBase $SearchBase -GroupCategory $GroupCategory -Credential $Credential
}

$ExportRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Export complete: $OutputPath"
Write-Host "`tGroups exported       : $($Stats.GroupsExported)"
Write-Host "`tMember rows exported  : $($Stats.MemberRowsExported)"
Write-Host "`tUnresolved members    : $($Stats.UnresolvedMembers)"
if ($Stats.UnresolvedMembers -gt 0)
{
	Write-Warning "Rows with MemberResolved = False could not be matched to an AD object (commonly ForeignSecurityPrincipals from a trusted domain). Review these before importing into the target forest."
}

#EndRegion
