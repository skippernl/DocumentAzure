﻿<#
.SYNOPSIS
Azure2Word uses powershell to get Azure information and puts this in a Word document
.DESCRIPTION
Connects to Azure powershell to get (virtual server) information
.PARAMETER Customer
[REQUIRED] This is the companyname that is used in the documentation
.PARAMETER ReportPath
[REQUIRED] This is the path where the report is stored
.PARAMETER TenantId
[OPTIONAL] When you have mutiple Tennants (CSP) enter the Tenant GUID
.PARAMETER SubscriptionId
[OPTIONAL] When you have mutiple Subscriptions (CSP) enter the Subscription GUID
.\Azure2Word -Customer Contoso -ReportPath c:\reports
    Runs the script for customer Contoso and create the word file in c:\reports
.\Azure2Word -Customer Contoso -ReportPath c:\reports -TenantId xxxx-xxxx -SubscriptionId yyyy-yyyy
    Runs the script for customer Contoso and create the word file in c:\reports 
    Running the scipt on TenantID xxxx-xxxx and SubscriptionID yyyy-yyy
.PARAMETER SkipVaults
[OPTIONAL] set to false to skip Backup and Replication
    .\Azure2Word -Customer Contoso -ReportPath c:\reports -SkipVaults:$true
Author: Xander Angenent (@XaAng70)
Adapted and fixed errors by SkipperNL
Idea: Anders Bengtsson http://contoso.se/blog/?p=4286
The Word file generation part of the script is based upon the work done by:

Carl Webster  | http://www.carlwebster.com | @CarlWebster
Iain Brighton | http://virtualengine.co.uk | @IainBrighton
Jeff Wouters  | http://www.jeffwouters.nl  | @JeffWouters

Uses modules AZ and Az.Reservations
Install-module -Name az
Install-Module -Name Az.Accounts -RequiredVersion 1.9.2
Install-Module -Name Az.Reservations
Last Modified: 2020/12/18
#>

Param
(
    [Parameter(Mandatory = $true)]
    $Customer,
    [Parameter(Mandatory = $true)]
    $ReportPath,
    $SkipVaults,
    $TenantId,
    $SubscriptionId
)
<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This Function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this Function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	$WordTable = AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	$WordTable = AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	$WordTable = AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	$WordTable = AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	$WordTable = AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>
Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Columns = $null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string[]] $Headers = $null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $AutoFit = $wdAutoFitContent,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines=$false,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Format = '-231'
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($null -eq $Columns) -and ($null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $null;
		}
		ElseIf(($null -ne $Columns) -and ($null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
        [System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($null -ne $Headers) 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($null -ne $Headers) 
					{ 
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
                        $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
                    $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end Switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Document.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $true);
			$ConvertToTableArguments.Add("ApplyShading", $true);
			$ConvertToTableArguments.Add("ApplyFont", $true);
			$ConvertToTableArguments.Add("ApplyColor", $true);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $true); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $true);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $true);
			$ConvertToTableArguments.Add("ApplyLastColumn", $true);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$null,                                          # Modifiers
			$null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting)
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		#the next line causes the heading row to flow across page breaks
		$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}
Function ArrayToLine ($ImputArray) {
	
	$ArrayLine = "" 
	Foreach ($ImputMember in $ImputArray) {
		if ($ArrayLine -eq "") { $ArrayLine = $ImputMember }
		else { $ArrayLine = $ArrayLine + "," + $ImputMember }
	}
	Return $ArrayLine
}
Function ConvertArrayToLine ($ConvertArray) {
#This function converts an Array to a sinlge line of text)
if ($ConvertArray) {
        $TextLine = ""
        foreach ($Member in $ConvertArray) {
            $TextLine = $TextLine + $Member + ","
        }
    }
    $TextLine = $TextLine.Substring(0,$TextLine.Length-1)

    return $TextLine

}
Function FindVM ($VMNic) {
#This function searches the Virtual Machine that is connected to a Nic interface
    $VirtualMachine = "Unkown"
    foreach ($VM in $VMArray) {
        if ($VM.nic -eq $VMNic) {
            $VirtualMachine = $VM.VMName
            break
        }
    }
    return $VirtualMachine
}
Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Document.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Selection.EndKey($wdStory,$wdMove) | Out-Null
}
Function GroupArrayToLine ($ImputArray) {
	
	$ArrayLine = "" 
	Foreach ($ImputMember in $ImputArray) {
		$Parts = $ImputMember.Split("/")
		if ($ArrayLine -eq "") { $ArrayLine = $Parts[8] }
		else { $ArrayLine = $ArrayLine + "," + $Parts[8] }
	}
	Return $ArrayLine
}
Function InitFirewallRule {
    $InitRule = New-Object System.Object;
    $InitRule | Add-Member -type NoteProperty -name Destination -Value ""
	$InitRule | Add-Member -Type NoteProperty -Name destinationPorts -Value ""
	$InitRule | Add-Member -type NoteProperty -name Firewall -Value ""
	$InitRule | Add-Member -type NoteProperty -name FirewallPolicyName -Value ""
	$InitRule | Add-Member -type NoteProperty -name FirewallPolRuleName -Value ""
    $InitRule | Add-Member -type NoteProperty -name FirewallRuleName -Value ""
	$InitRule | Add-Member -type NoteProperty -name FirewallRulePrio -Value ""
    $InitRule | Add-Member -type NoteProperty -name NetworkRuleCollection -Value ""
	$InitRule | Add-Member -type NoteProperty -name Protocols -Value ""
	$InitRule | Add-Member -Type NoteProperty -Name ruleType -Value ""
    $InitRule | Add-Member -type NoteProperty -name Source -Value ""
	$InitRule | Add-Member -Type NoteProperty -Name translated -Value ""
	$InitRule | Add-Member -Type NoteProperty -Name translatedPort -Value ""
    
    return $InitRule      
}

$StartScriptTime = get-date 
Write-Output "Script Started."
if ($TenantId -and $SubscriptionId) {
    Connect-AzAccount -TenantId $TenantId -SubscriptionId $SubscriptionId | Out-Null
    if (!($?)) {
        Write-Output "Error logging in to Azure."
        Write-Output "Script stopped."
        exit
    }
}
elseif ($TenantId) {
    Connect-AzAccount -TenantId $TenantId | Out-Null
    if (!($?)) {
        Write-Output "Error logging in to Azure."
        Write-Output "Script stopped."
        exit
    }
}
elseif ($SubscriptionId) {
    Connect-AzAccount -SubscriptionId $SubscriptionId | Out-Null
    if (!($?)) {
        Write-Output "Error logging in to Azure."
        Write-Output "Script stopped."
        exit
    }
}
else {
    Connect-AzAccount | Out-Null
    if (!($?)) {
        Write-Output "Error logging in to Azure."
        Write-Output "Script stopped."
        exit
    }    
}
##Init Arrays and other default parameters
$NetworkGatewayConnections = [System.Collections.ArrayList]@()
$LocalEndPointResourceGroupNames = [System.Collections.ArrayList]@()
$LocalGatewayArray = [System.Collections.ArrayList]@()
[int]$wdMove = 0
[int]$wdSeekMainDocument = 0
#[int]$wdSeekPrimaryFooter = 4
[int]$wdStory = 6
#[long]$wdColorRed = 255
#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
#[int]$wdAutoFitFixed = 0
[int]$wdAutoFitContent = 1
#[int]$wdAutoFitWindow = 2
#[int]$wdLineStyleNone = 0
[int]$wdLineStyleSingle = 1
#Work-around for Powershell 7 not loading assembly files 
$RootAssemblyPathWord = $env:Systemroot + "\assembly\GAC_MSIL\Microsoft.Office.Interop.Word"
$WordAssemblyDirectory = Get-ChildItem -Path $RootAssemblyPathWord -Directory | Select-Object -First 1
$WordAssemblyFile = $WordAssemblyDirectory.FullName + "\Microsoft.Office.Interop.Word.dll"
add-type -literalPath $WordAssemblyFile
#End init
$Report = "$ReportPath\$Customer-Azure.docx"
# Creating the Word Object, set Word to not be visual and add document
$Word = New-Object -ComObject Word.Application
$Word.Visible = $false
$Document = $Word.Documents.Add()
#Switch to landscape
$Document.PageSetup.Orientation = 1
$Selection = $Word.Selection
Write-Output "Getting Word information."
$ALLStyles = $document.Styles | Select-Object NameLocal 
$Title = $AllStyles[360].Namelocal
$Heading1 = $AllStyles[149].Namelocal
$Heading2 = $AllStyles[150].Namelocal
$Heading3 = $AllStyles[151].Namelocal
$HeaderFooterIndex = "microsoft.office.interop.word.WdHeaderFooterIndex" -as [type]
$alignmentTab = "microsoft.office.interop.word.WdAlignmentTabAlignment" -as [type]
$section = $Document.sections.item(1)
$header = $section.headers.item($HeaderFooterIndex::wdHeaderFooterFirstPage)
$header.range.InsertAlignmentTab($alignmentTab::wdRight)
$header.range.InsertAfter("Azure documentation for $Customer")
## Add some text to start with
$Selection.Style = $Title
$Selection.TypeText("Azure documentation for $Customer")
$Selection.TypeParagraph()
$Selection.TypeParagraph()

### Add the TOC
$range = $Selection.Range
$toc = $Document.TablesOfContents.Add($range)
$Selection.TypeParagraph()

## Get all VMs from Azure
#Connect-AzAccount
Write-Output "Getting all Azure resources."
$ALLAzureResources = Get-AzResource

###
### VIRTUAL MACHINES
###

## Add some text
$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Virtual Machines")
$Selection.TypeParagraph()


Write-Output "Selecting VM's."
$VMs = Get-AzVM -Status | Sort-Object Name
$VMArray = [System.Collections.ArrayList]@()

## Values
Write-Output "Creating VM table."
Foreach ($VM in $VMs) {

    $TableMember = New-Object System.Object;
    $VMName = $VM.NetworkProfile.NetworkInterfaces.ID
    $Parts = $VMName.Split("/")
    $NICLabel = $PArts[8]
    
    $TableMember | Add-Member -type NoteProperty -name VMName -Value $VM.Name
    if ($null -eq $VM.Zones) 
        {$AvailablityZone = "-"
    }
    else {
        $AvailablityZone = $VM.Zones
    }
    if ($null -eq $VM.OSProfile.ComputerName) {
        #https://docs.microsoft.com/en-us/troubleshoot/azure/virtual-machines/computer-names-missing-blank
        $VirtualMachineName = "(" + $VM.Name + ")"
    }
    else {
        $VirtualMachineName = $VM.OSProfile.ComputerName
    }
    $TableMember | Add-Member -type NoteProperty -name Computername -Value $VirtualMachineName
    $TableMember | Add-Member -type NoteProperty -name Size -Value $VM.HardwareProfile.VmSize
    $TableMember | Add-Member -type NoteProperty -name RGN -Value $VM.ResourceGroupName
    $TableMember | Add-Member -type NoteProperty -name NIC -Value $NICLabel
    $TableMember | Add-Member -type NoteProperty -name Status -Value $VM.Powerstate
    $TableMember | Add-Member -type NoteProperty -name Zone -Value $AvailablityZone
    $VMArray.Add($TableMember) | Out-Null
}

FindWordDocumentEnd
$VMArray = $VMArray | Sort-Object VMName
$WordTable = AddWordTable -CustomObject $VMArray -Columns VMName, Computername, RGN, Size, NIC, Status, Zone -Headers  "VM Name", "Computer name", "Resource Group Name", "VM Size", "Network Interface", "Power Status", "Zone"
FindWordDocumentEnd
$Selection.TypeParagraph()
$Selection.Style = $Heading2
$Selection.TypeText("Virtual Machine Disks")
$Selection.TypeParagraph()
Write-Output "Getting disk information."
$Disks = get-Azdisk | Sort-Object Name

## Add a table for Disks
$TableArray = [System.Collections.ArrayList]@()

## Values
Write-Output "Creating disk table."
Foreach ($Disk in $Disks) {

    $TableMember = New-Object System.Object;
    $DiskID = $Disk.ManagedBy
    if ($null -eq $DiskID) {
        $Server = "Unknown"
    }
    else {
        $Parts = $DiskID.Split("/")
        $Server = $Parts[8]
    }
    
    $TableMember | Add-Member -type NoteProperty -name DiskName -Value $Disk.Name
    $TableMember | Add-Member -type NoteProperty -name ServerName -Value $Server
    $TableMember | Add-Member -type NoteProperty -name DiskIO -Value $Disk.DiskIOPSReadWrite.ToString()
    $TableMember | Add-Member -type NoteProperty -name DiskMB -Value $Disk.DiskMBpsReadWrite.ToString()
    $TableMember | Add-Member -type NoteProperty -name RGN -Value $Disk.ResourceGroupName
    $TableMember | Add-Member -type NoteProperty -name Size -Value $Disk.DiskSizeGB.ToString()
    $TableArray.Add($TableMember) | Out-Null

}
FindWordDocumentEnd
$WordTable = AddWordTable -CustomObject $TableArray -Columns DiskName, ServerName, RGN, DiskIO, DiskMB, Size -Headers "Disk Name", "Server Name", "Resource Group Name", "IOPS ReadWrite", "MBps ReadWrite", "Size (GB)"
FindWordDocumentEnd
$Selection.TypeParagraph()

########
######## NETWORK INTERFACE
########

$Selection.Style = $Heading2
$Selection.TypeText("Network interfaces")
$Selection.TypeParagraph()

Write-Output "Getting network interfaces."
$NICs = Get-AzNetworkInterface | Sort-Object Name
$TableArray = [System.Collections.ArrayList]@()

## Write NICs to NIC table 

Write-Output "Creating NIC table"
Foreach ($NIC in $NICs) {

    $TableMember = New-Object System.Object;
    ## Get connected VM, if there is one connected to the network interface
    If (!$NIC.VirtualMachine.id) {
        $VMLabel = "-"
    }
    Else
    {
        $VMName = $NIC.VirtualMachine.id
        $Parts = $VMName.Split("/")
        $VMLabel = $PArts[8]
    }
    ## GET VNET and SUBNET
    $NETCONF = $NIC.IPconfigurations.subnet.id
    $Parts = $NETCONF.Split("/")
    $VNETNAME = $PArts[8]
    $SUBNETNAME = $PArts[10]

    $TableMember | Add-Member -type NoteProperty -name VMName -Value $VMLabel
    $TableMember | Add-Member -type NoteProperty -name NIC -Value $NIC.Name
    $TableMember | Add-Member -type NoteProperty -name RGN -Value $NIC.ResourceGroupName
    $TableMember | Add-Member -type NoteProperty -name VNet -Value $VNETNAME 
    $TableMember | Add-Member -type NoteProperty -name Subnet -Value $SUBNETNAME
    $TableMember | Add-Member -type NoteProperty -name IP -Value $NIC.IPconfigurations.PrivateIpAddress
    $TableMember | Add-Member -type NoteProperty -name Alloc -Value $NIC.IPconfigurations.PrivateIpAllocationMethod
    $TableArray.Add($TableMember) | Out-Null
}
FindWordDocumentEnd
$WordTable = AddWordTable -CustomObject $TableArray -Columns VMName, NIC, RGN, VNet, Subnet, IP, Alloc -Headers  "Virtual Machine", "Network Card Name", "Resource Group Name", "VNET", "Subnet", "IP Address", "Allocation Method"
FindWordDocumentEnd
$Selection.TypeParagraph()

##Find reservations
$Selection.Style = $Heading2
$Selection.TypeText("Reservations")
$Selection.TypeParagraph()
Write-Output "Getting reservations."
#If there are no reservations - Operation returned an invalid status code 'Forbidden' is being displayed
$ALLReservationOrders = Get-AzReservationOrder | Sort-Object Name

Write-Output "Creating reservation table."
$TableArray = [System.Collections.ArrayList]@()
Foreach ($ReservationOrder in $ALLReservationOrders) {  
    $AllReservations = Get-AzReservation -ReservationOrderId $ReservationOrder.Name
    $StartTime = $ReservationOrder.CreatedDateTime
    $Term = $ReservationOrder.Term
    $Duration = $Term.substring(1,1)
    $LastChar = $Term.substring(2,1)
    #Using switch to be flexible
    switch ($LastChar) {
        "Y" {
            $EndTime = $StartTime.AddYears($Duration)
        }
    }
    foreach ($Reservation in $AllReservations) {
        $TableMember = New-Object System.Object;
        $TableMember | Add-Member -type NoteProperty -name DisplayName $Reservation.DisplayName
        $TableMember | Add-Member -type NoteProperty -name sku -Value $Reservation.Sku
        $TableMember | Add-Member -type NoteProperty -name Quantity -Value $Reservation.Quantity
        $TableMember | Add-Member -type NoteProperty -name StartTime -Value $StartTime.ToString()
        $TableMember | Add-Member -type NoteProperty -name Term -Value $Term
        $TableMember | Add-Member -type NoteProperty -name EndTime -Value $EndTime.ToString()
        $TableMember | Add-Member -type NoteProperty -name State -Value $Reservation.ProvisioningState
        $TableArray.Add($TableMember) | Out-Null
    }
}
FindWordDocumentEnd
if ($TableArray) {
    $WordTable = AddWordTable -CustomObject $TableArray -Columns Displayname, Sku, Quantity, StartTime, Term, EndTime, State -Headers "Displayname", "VMType", "Quantity", "Start", "Term", "End", "Status"
}
else {
    $Selection.TypeParagraph()
    $Selection.TypeText("No reservations found.")  
}
FindWordDocumentEnd
$Selection.TypeParagraph()

########
######## Create a table for NSG
########

$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Network security groups")
$Selection.TypeParagraph()
Write-Output "Getting NSGs"
$NSGs = Get-AzNetworkSecurityGroup | Sort-Object Name

## Write NICs to NIC table 

Write-Output "Creating NSG table."
$TableArray = [System.Collections.ArrayList]@()

Foreach ($NSG in $NSGs) {
    $TableMember = New-Object System.Object;
    ## Get connected NIC, if there is one connected 
    If (!$NSG.NetworkInterfaces.Id) { 
        $NICLabel = " "
    }
    Else {
        $SubnetName = $NSG.NetworkInterfaces.Id
        $Parts = $SubnetName.Split("/")
        $NICLabel = $PArts[8]
    }
    ## Get connected SUBNET, if there is one connected 
    If (!$NSG.Subnets.Id) { 
        $SubnetLabel = " "
    }
    Else
    {
        $SUBNETName = $NSG.Subnets.Id
        $Parts = $SUBNETName.Split("/")
        $SUBNETLabel = $PArts[10]
    }
    $TableMember | Add-Member -type NoteProperty -name NSG -Value $NSG.Name
    $TableMember | Add-Member -type NoteProperty -name RGN -Value $NSG.ResourceGroupName
    $TableMember | Add-Member -type NoteProperty -name NI -Value $NICLabel
    $TableMember | Add-Member -type NoteProperty -name Subnets -Value $SUBNETLabel
    $TableArray.Add($TableMember) | Out-Null
}
FindWordDocumentEnd
$WordTable = AddWordTable -CustomObject $TableArray -Columns NSG, RGN, NI, Subnets -Headers "NSG Name", "Resource Group Name", "Network Interfaces", "Subnets"
FindWordDocumentEnd
$Selection.TypeParagraph()


########
######## Create a table for each NSG
########

Write-Output "Creating rule table."
ForEach ($NSG in $NSGs) {

    ## Add Heading for each NSG
    $Selection.Style = $Heading2
    $Selection.TypeText($NSG.Name)
    $Selection.TypeParagraph()

    $TableArray = [System.Collections.ArrayList]@()

	$NSGRulesCustom = Get-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $NSG | Sort-Object Name
	$NSGRulesDefault = Get-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $NSG -DefaultRules | Sort-Object Name
    ### Add a table for each NSG, the NSg has custom rules

    ForEach ($NSGRULE in $NSGRulesCustom) {
        $TableMember = New-Object System.Object;
        $TableMember | Add-Member -type NoteProperty -name RuleName -Value $NSGRule.Name
        $TableMember | Add-Member -type NoteProperty -name Protocol -Value $NSGRule.Protocol
        $Value = ArrayToLine $NSGRule.SourcePortRange
        $TableMember | Add-Member -type NoteProperty -name SourcePort -Value $Value
        $Value = ArrayToLine $NSGRule.DestinationPortRange
        $TableMember | Add-Member -type NoteProperty -name DestPort -Value $Value
        $Value = ArrayToLine $NSGRule.SourceAddressPrefix
        $TableMember | Add-Member -type NoteProperty -name SourcePref -Value $Value
        $Value = ArrayToLine $NSGRule.DestinationAddressPrefix
        $TableMember | Add-Member -type NoteProperty -name DestPref -Value $Value
        $TableMember | Add-Member -type NoteProperty -name Access -Value $NSGRule.Access
        $TableMember | Add-Member -type NoteProperty -name Prio -Value $NSGRule.Priority.ToString()
        $TableMember | Add-Member -type NoteProperty -name Direction -Value $NSGRule.Direction
        $TableArray.Add($TableMember) | Out-Null
    }

    ForEach ($NSGRULE in $NSGRulesDefault) {
        $TableMember = New-Object System.Object;
        $TableMember | Add-Member -type NoteProperty -name RuleName -Value $NSGRule.Name
        $TableMember | Add-Member -type NoteProperty -name Protocol -Value $NSGRule.Protocol
        $Value = ArrayToLine $NSGRule.SourcePortRange
        $TableMember | Add-Member -type NoteProperty -name SourcePort -Value $Value
        $Value = ArrayToLine $NSGRule.DestinationPortRange
        $TableMember | Add-Member -type NoteProperty -name DestPort -Value $Value
        $Value = ArrayToLine $NSGRule.SourceAddressPrefix
        $TableMember | Add-Member -type NoteProperty -name SourcePref -Value $Value
        $Value = ArrayToLine $NSGRule.DestinationAddressPrefix
        $TableMember | Add-Member -type NoteProperty -name DestPref -Value $Value
        $TableMember | Add-Member -type NoteProperty -name Access -Value $NSGRule.Access
        $TableMember | Add-Member -type NoteProperty -name Prio -Value $NSGRule.Priority.ToString()
        $TableMember | Add-Member -type NoteProperty -name Direction -Value $NSGRule.Direction
        $TableArray.Add($TableMember) | Out-Null
    }

    ### Close the NSG table
    FindWordDocumentEnd
    $WordTable = AddWordTable -CustomObject $TableArray -Columns RuleName, Protocol, SourcePort, DestPort, SourcePref, DestPref, Access, Prio, Direction -Headers "Rule Name","Protocol", "Source Port Range", "Destination Port Range", "Source Address Prefix", "Destination Address Prefix", "Access", "Priority", "Direction"
    FindWordDocumentEnd
    $Selection.TypeParagraph()
    

}

##Get al Azure VPNs
$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("VPN information")
$Selection.TypeParagraph()
$Selection.Style = $Heading2
$Selection.TypeText("VPN gateway connections")
$Selection.TypeParagraph()
Write-Output "Getting VPN gateway connections"
$NetworkConnections = $ALLAzureResources | Where-Object {$_.ResourceType -eq "Microsoft.Network/connections" } | Sort-Object Name
Foreach ($NetworkConnection in $NetworkConnections) {
    $NGC = Get-AzVirtualNetworkGatewayConnection -ResourceName $NetworkConnection.ResourceName -ResourceGroupName $NetworkConnection.ResourceGroupName
    $NetworkGatewayConnections.Add($NGC) | Out-Null
}
$NetworkGatewayConnections = $NetworkGatewayConnections | Sort-Object Name
$TableArray = [System.Collections.ArrayList]@()

########
######## Create a table for VPN GatewayConnections
########
## Values
Write-Output "Creating VPN table."
$ExpressRouteFound=$False
Foreach ($NGC in $NetworkGatewayConnections) {
    $TableMember = New-Object System.Object;
    $ResourceGroupName = $NGC.ResourceGroupName
    $Parts = $NGC.VirtualNetworkGateway1.id.Split("/")
    $AzEndpoint = $Parts[8]
    if ($NGC.VirtualNetworkGateway2.id) {
        $Parts = $NGC.LocalNetworkGateway2.id.Split("/")
        $LocalEndpoint = $Parts[8]
        $Expressroute = "No"
    }
    else {
        #no endpoint Could it be an Expressroute?
        if ($NGC.Peer) {
            $Parts = $NGC.Peer.ID.Split("/")
            if ($Parts[7] -eq "expressRouteCircuits") {
                $Expressroute = "Yes"
                $LocalEndpoint = $Parts[8]
                $ExpressRouteFound=$True
            }
            Else {
                $Expressroute = "No"
                $LocalEndpoint = "Empty"
            }
        }
        Else {
            $Expressroute = "No"
            $LocalEndpoint = "Empty"
        }
    }
    if (!($LocalEndPointResourceGroupNames.Contains($ResourceGroupName))) { $LocalEndPointResourceGroupNames.Add($ResourceGroupName) | Out-Null }
    $TableMember | Add-Member -type NoteProperty -name VPN -Value $NGC.Name
    $TableMember | Add-Member -type NoteProperty -name RGN -Value $ResourceGroupName
    $TableMember | Add-Member -type NoteProperty -name AzEndpoint -Value $AzEndpoint
    $TableMember | Add-Member -type NoteProperty -name LocalEndpoint -Value $LocalEndpoint
    $TableMember | Add-Member -type NoteProperty -name Status -Value $NGC.ConnectionStatus
    $DataGB = $NGC.EgressBytesTransferred/1048576
    $DataGB = [math]::Round($DataGB)
    $TableMember | Add-Member -type NoteProperty -name Egress -Value $DataGB.ToString()
    $DataGB = $NGC.IngressBytesTransferred/1048576
    $DataGB = [math]::Round($DataGB)
    $TableMember | Add-Member -type NoteProperty -name Ingress -Value $DataGB.ToString()
    $TableArray.Add($TableMember) | Out-Null
}
FindWordDocumentEnd
if ($TableArray) {
    $WordTable = AddWordTable -CustomObject $TableArray -Columns VPN, RGN, AzEndpoint, LocalEndpoint, Status, Egress, Ingress -Headers "VPN", "Resource Group", "Azure Endpoint", "Local Endpoint", "Status", "EgressBytesTransferred (GB)", "IngressBytesTransferred (GB)"
}
else {
    $Selection.TypeParagraph()
    $Selection.TypeText("No VPN found.")  
}
FindWordDocumentEnd
$Selection.TypeParagraph()
if ($ExpressRouteFound) {
    $Selection.Style = $Heading2
    $Selection.TypeText("ExpressRoutes")
    $Selection.TypeParagraph()
    $Selection.TypeParagraph()
    $ExpressRoutes = Get-AzExpressRouteCircuit
    foreach ($ExpressRoute in $ExpressRoutes) {
        $Selection.TypeText("Name :$($ExpressRoute.Name)")
        $Selection.TypeParagraph()
        $Selection.TypeText("Location :$($ExpressRoute.Location)")
        $Selection.TypeParagraph()
        $Selection.TypeText("ProvisioningState :$($ExpressRoute.ProvisioningState)")
        $Selection.TypeParagraph()
        $Selection.TypeText("ServiceProviderProperties")
        $Selection.TypeParagraph()
        $Selection.TypeText("ServiceProviderName $($ExpressRoute.ServiceProviderProperties.ServiceProviderName)")
        $Selection.TypeParagraph()        
        $Selection.TypeText("PeeringLocation $($ExpressRoute.ServiceProviderProperties.PeeringLocation)")
        $Selection.TypeParagraph() 
        $Selection.TypeText("BandwidthInMbps $($ExpressRoute.ServiceProviderProperties.BandwidthInMbps)")
        $Selection.TypeParagraph() 
        $Selection.TypeText("AzureASN $($ExpressRoute.Peerings.AzureASN)")
        $Selection.TypeParagraph() 
        $Selection.TypeText("PeerASN $($ExpressRoute.Peerings.PeerASN)")
        $Selection.TypeParagraph()    
        $Selection.TypeText("PrimaryPeerAddressPrefix $($ExpressRoute.Peerings.PrimaryPeerAddressPrefix)")
        $Selection.TypeParagraph()    
        $Selection.TypeText("SecondaryPeerAddressPrefix $($ExpressRoute.Peerings.SecondaryPeerAddressPrefix)")
        $Selection.TypeParagraph()    
        $Selection.TypeText("PrimaryAzurePort $($ExpressRoute.Peerings.PrimaryAzurePort)")
        $Selection.TypeParagraph()    
        $Selection.TypeText("SecondaryAzurePort $($ExpressRoute.Peerings.SecondaryAzurePort)")
        $Selection.TypeParagraph()    
        $Selection.TypeText("SecondaryPeerAddressPrefix $($ExpressRoute.Peerings.SecondaryPeerAddressPrefix)")
        $Selection.TypeParagraph()         
    }
}

$Selection.Style = $Heading2
$Selection.TypeText("VPN local gateways")
$Selection.TypeParagraph()
foreach ($LocalEndPointResourceGroupName in $LocalEndPointResourceGroupNames) {
    $LocalGateway = Get-AzLocalNetworkGateway -ResourceGroupName $LocalEndPointResourceGroupName
    if ($LocalGateway -is [array]) {
        Foreach ($LocalGatewayMember in $LocalGateway) { $LocalGatewayArray.Add($LocalGatewayMember) | Out-Null }
    }
    else {
        $LocalGatewayArray.Add($LocalGateway) | Out-Null
    }
}
$LocalGatewayArray = $LocalGatewayArray | Sort-Object Name
########
######## Create a table for VPN LocalGateways
########

## Values
$TableArray = [System.Collections.ArrayList]@()
Write-Output "Creating VPN localgateway table"
Foreach ($LocalGateway in $LocalGatewayArray) {
    $TableMember = New-Object System.Object;

    $TableMember | Add-Member -type NoteProperty -name Name -Value $LocalGateway.Name
    $TableMember | Add-Member -type NoteProperty -name RGN -Value $LocalGateway.ResourceGroupName
    $Gateway = $LocalGateway.GatewayIpAddress
    if (!($Gateway)) {
        #There is no IP adrress -> Try FQDN
        $Gateway = $LocalGateway.fqdn
        if (!($Gateway)) {
            #Also not found
            $Gateway = "Unkown"
        }
    }
    $TableMember | Add-Member -type NoteProperty -name GateIP -Value $Gateway
    if ($LocalGateway.LocalNetworkAddressSpace.AddressPrefixes) {
        $LocalAddressSpace = ConvertArrayToLine $LocalGateway.LocalNetworkAddressSpace.AddressPrefixes
    }
    else {
        $LocalAddressSpace = "NOT defined"
    }
    $TableMember | Add-Member -type NoteProperty -name LocalNetwork $LocalAddressSpace
    $TableArray.Add($TableMember) | Out-Null
}
FindWordDocumentEnd
if ($TableArray) { 
    $WordTable = AddWordTable -CustomObject $TableArray -Columns Name, RGN, GateIP, LocalNetwork -Headers "Name", "Resource Group", "FQDN/Gateway Address", "Local Network Address Space"
}
else {
    $Selection.TypeParagraph()
    $Selection.TypeText("No local gateway found.")  
}
FindWordDocumentEnd
$Selection.TypeParagraph()

$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Public IP's")
$Selection.TypeParagraph()
$AllPublicIPs = Get-AzPublicIpAddress | Sort-Object Name
########
######## Create a table for Public IP addresses
########


## Values
Write-Output "Creating public IP table."
$TableArray = [System.Collections.ArrayList]@()
Foreach ($PublicIP in $AllPublicIPs) {
    $TableMember = New-Object System.Object;

    $TableMember | Add-Member -type NoteProperty -name Name -Value $PublicIP.Name
    $TableMember | Add-Member -type NoteProperty -name RGN -Value  $PublicIP.ResourceGroupName
    $TableMember | Add-Member -type NoteProperty -name IP -Value  $PublicIP.IpAddress
    if ($PublicIP.IpConfiguration.id) {
        $Parts = $PublicIP.IpConfiguration.id.Split("/")
        $Endpoint = $Parts[8]
    }
    else {
        $Endpoint = "Unused or NAT"
    }
    $TableMember | Add-Member -type NoteProperty -name Endpoint -Value  $Endpoint
    $TableArray.Add($TableMember) | Out-Null
}
FindWordDocumentEnd
if ($TableArray) { 
    $WordTable = AddWordTable -CustomObject $TableArray -Columns Name, RGN, IP, Endpoint -Headers "Name", "Resource Group", "IP Address", "Usedby"
}
else {
    $Selection.TypeParagraph()
    $Selection.TypeText("No public IP found.")  
}
FindWordDocumentEnd
$Selection.TypeParagraph()

#Adding recovery information
$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Backup and Replication")
$Selection.TypeParagraph()
FindWordDocumentEnd
if ($SkipVaults) {
    $Selection.TypeText("-SkipVaults option selected.")
    $Selection.TypeParagraph()
    $Selection.TypeText("No information documented.") 
    $Selection.TypeParagraph()
    FindWordDocumentEnd
    Write-Output "Skipping vault information"
}
else {
    $Vaults = Get-AzRecoveryServicesVault | Sort-Object Name
    ########
    ######## Create a table the backupjobs found
    ########

    ## Values
    #Get Only Restore points of the last week.
    $startDate = (Get-Date).AddDays(-7)
    $endDate = Get-Date
    Write-Output "Creating backup job table."
    $Selection.Style = $Heading2
    $Selection.TypeText("Backup")
    $Selection.TypeParagraph()
    $TableArray = [System.Collections.ArrayList]@()
    $BackupPolicies = [System.Collections.ArrayList]@()
    $ASRPolicies = [System.Collections.ArrayList]@()
    $BackupFailed = 0
    $BackupJobFailed = $null
    $CounterVault = 1
    $MaxCounterVault=$Vaults.Count
    Foreach ($Vault in $Vaults) {
        #Fill the Replication fabric (No need to run to all the vaults twice)
        Set-AzRecoveryServicesAsrVaultContext -Vault $vault | Out-Null
        $ASRFabrics += Get-AzRecoveryServicesAsrFabric
        $BackupPolicy = Get-AzRecoveryServicesBackupProtectionPolicy -VaultID $Vault.ID
        $BackupPolicies.Add($BackupPolicy) | Out-Null
        $ASRPolicy = Get-AzRecoveryServicesAsrPolicy
        if ($null -ne $ASRPolicy) { $ASRPolicies.Add($ASRPolicy) | Out-Null }
        #End fill Replication Fabric
        $ProcVault = $CounterVault/$MaxCounterVault*100
        $ProcVaultString = $ProcVault.ToString("0.00")
        Write-Progress -ID 0 -Activity "Checking vault $CounterVault/$MaxCounterVault $($Vault.Name) ($ProcVaultString%)" -PercentComplete ($ProcVault)
        $BackupJobs = Get-AzRecoveryServicesBackupJob -VaultId $Vault.ID
        $namedContainerVMs = Get-AzRecoveryServicesBackupContainer  -ContainerType "AzureVM" -Status "Registered" -VaultId $Vault.ID
        $CounterVault++
        $CounterBackupJob = 1
        $MaxCounterBackupJob = $BackupJobs.Count
        $Activity = $null
        ForEach ($BackupJob in $BackupJobs) {
            $TableMember = New-Object System.Object;
            $ProcBackup = $CounterBackupJob/$MaxCounterBackupJob*100
            $ProcBackupString = $ProcBackup.ToString("0.00")
            $Activity = "Checking backup job $CounterBackupJob/$MaxCounterBackupJob $($BackupJob.WorkloadName.ToUpper()) ($ProcBackupString%)" 
            Write-Progress -ID 1 -Activity $Activity -PercentComplete ($ProcBackup)
            #There can be multiple Restore Points due to the fact that there could be more Jobs (after changing resource group etc)
            $CounterBackupJob++
            $rp = @()
            switch ($BackupJob.BackupManagementType) {
                "AzureVM" {
                    foreach ($namedContainer in $namedContainerVMs) {
                        #Friendly name can be in multiple namedcontainers
                        if ($BackupJob.workloadname.ToUpper() -eq $namedContainer.FriendlyName.ToUpper()) {
                            $BackupnamedContainer = $namedContainer
                            $backupitem = Get-AzRecoveryServicesBackupItem -Container $BackupnamedContainer -WorkloadType $BackupJob.BackupManagementType -VaultId $Vault.ID
                            $rp = Get-AzRecoveryServicesBackupRecoveryPoint -Item $backupitem -StartDate $startdate.ToUniversalTime() -EndDate $enddate.ToUniversalTime() -VaultId $Vault.ID
                        }
                    }
                    $WorkloadName = $BackupJob.workloadname.ToUpper()
                }
                "AzureWorkload" {
                    $WorkloadArray = $BackupJob.workloadname.Split(" ")
                    if ($WorkloadArray.Count -eq 2) {
                        $SQLDatabase = $WorkloadArray[0]
                        $SQLServer = $WorkloadArray[1].Substring(1,$WorkloadArray[1].Length-2)
                    }
                    else {
                        $SQLDatabase = $BackupJob.workloadname
                        $SQLServer = "FirstBackup"                 
                    }
                    #As we are checking for a database we might get more than one result. 
                    $bkpItems = Get-AzRecoveryServicesBackupItem -BackupManagementType $BackupJob.BackupManagementType -WorkloadType MSSQL -Name $SQLDatabase -VaultId $Vault.ID
                    if ($BackupJob.Operation -eq "ConfigureBackup") {
                        $LatestRestorePoint = "Configuring"
                    }
                    else {
                        $LatestRestorePoint = "Unkown"
                    }
                    foreach ($backupitem in $bkpItems) {
                        if ($backupitem.ParentType -eq "AzureVmWorkloadSQLAvailabilityGroup") {
                            $SQLServer = $backupitem.ParentName
                        }
                        $ServerNameArray=$backupitem.ServerName.Split(".")
                        $BackupSQL = $ServerNameArray[0]
                        if ($BackupSQL -eq $SQLServer){ 
                            $rp += Get-AzRecoveryServicesBackupRecoveryPoint -Item $backupitem -StartDate $startdate.ToUniversalTime() -EndDate $enddate.ToUniversalTime() -VaultId $Vault.ID
                        }
                    }  
                    $WorkloadName = $SQLServer.ToUpper() + " " + $SQLDatabase.ToUpper()
                }
                default {
                    if ($BackupJob.Operation -eq "ConfigureBackup") {
                        $LatestRestorePoint = "Configuring"
                    }
                    else {
                        $LatestRestorePoint = "Unkown"
                    }
                }
            }
            if ($rp) {
                $rp = $rp | Sort-Object RecoveryPointTime -Descending
                $LatestRestorePoint = $rp[0].RecoveryPointTime.ToString()
            }
            else {
                if ($BackupJob.Operation -eq "ConfigureBackup") {
                    $LatestRestorePoint = "Configuring"
                }
                else {
                    $LatestRestorePoint = "Unkown"
                }
            }
            $TableMember | Add-Member -type NoteProperty -name Name -Value $Vault.Name
            $TableMember | Add-Member -type NoteProperty -name Workload -Value $WorkloadName
            $TableMember | Add-Member -type NoteProperty -name Status -Value $BackupJob.Status
            $TableMember | Add-Member -type NoteProperty -name StartTime -Value $BackupJob.StartTime.ToString()
            if ($BackupJob.Status -eq "InProgress") {
                $BackupJobEndTime = "-"
            }
            else {
                $BackupJobEndTime = $BackupJob.EndTime.ToString()
            }
            if ($BackupJob.Status -eq "failed") {
                $BackupFailed++
                if ($BackupJobFailed -eq 1) {
                    $BackupJobFailed = $BackupJobFailed + ", $WorkloadName"
                }
                else {
                    $BackupJobFailed = $WorkloadName
                }
            }
            $TableMember | Add-Member -type NoteProperty -name EndTime -Value $BackupJobEndTime
            $TableMember | Add-Member -type NoteProperty -name RP -Value $LatestRestorePoint
            $TableArray.Add($TableMember) | Out-Null
        }
        if ($Activity) { Write-Progress -ID 1 -Activity $Activity -Status "Ready" -Completed }
    }
    Write-Progress -ID 0 -Activity $Activity -Status "Ready" -Completed
    FindWordDocumentEnd
    if ($TableArray){ 
        $TableArray = $TableArray | Sort-Object Name, Workload
        $WordTable = AddWordTable -CustomObject $TableArray -Columns Name, Workload, Status, StartTime, EndTime, RP -Headers "Vault", "Backup Item", "Status", "Start Time (UTC)", "End Time (UTC)", "Latest RestorePoint (UTC)"
        FindWordDocumentEnd
        $Selection.TypeParagraph()
        switch ($backupFailed) {
            1 { 
                $Selection.TypeText("One failed backup was found!")
                $Selection.TypeParagraph()
                $Selection.TypeText("The job that failed is: $BackupJobFailed.")   
                $Selection.TypeParagraph()
            }
            2 {
                $Selection.TypeText("Two or more failed backups where found!")
                $Selection.TypeParagraph()
                $Selection.TypeText("The jobs that failed are $BackupJobFailed.")   
                $Selection.TypeParagraph()
            }
            Default {}
        }
    }
    else {
        $Selection.TypeParagraph()
        $Selection.TypeText("No backups found.") 
        $Selection.TypeParagraph()
    }
    FindWordDocumentEnd
    $Selection.Style = $Heading3
    $Selection.TypeText("Backup policy")
    $Selection.TypeParagraph()
    If ($BackupPolicies) {
        $TableArray = [System.Collections.ArrayList]@()
        foreach ($BackupPolicy in $BackupPolicies) {
            foreach ($BackupSchedule in $BackupPolicy) {
                $Selection.TypeParagraph()
                $Selection.TypeText("$($BackupSchedule.Name)")
                $Selection.TypeParagraph()
                $Selection.TypeText("$($BackupSchedule.WorkloadType)")
                $Selection.TypeParagraph()
                switch ($BackupSchedule.WorkloadType) {
                    "AzureVM" {     
                        $BackupSchRP = $BackupSchedule.RetentionPolicy
                        if ($BackupSchRP.IsDailyScheduleEnabled) {
                            $Selection.TypeText("Daily schedule is enabled.")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  Duration count in Days   : $($BackupSchRP.DailySchedule.DurationCountInDays)")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.DailySchedule.RetentionTimes
                            $Selection.TypeText("  RetentionTime            : $Value")
                            $Selection.TypeParagraph()
                        }
                        else  {
                            $Selection.TypeText("Daily schedule is disabled.")
                            $Selection.TypeParagraph()
                        }
                        if ($BackupSchRP.IsWeeklyScheduleEnabled) { 
                            $Selection.TypeText("Weekly schedule is enabled.")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  Duration count in Weeks  : $($BackupSchRP.WeeklySchedule.DurationCountInWeeks)")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.WeeklySchedule.RetentionTimes
                            $Selection.TypeText("  RetentionTime            : $Value")
                            $Selection.TypeParagraph()
                            $DayOfWeek = ArrayToLine $BackupSchRP.WeeklySchedule.DaysOfTheWeek
                            $Selection.TypeText("  Days of Week             : $DayOfWeek")
                            $Selection.TypeParagraph()
                        }
                        else  {
                            $Selection.TypeText("Weekly schedule is disabled.")
                            $Selection.TypeParagraph()
                        }
                        if ($BackupSchRP.IsMonthlyScheduleEnabled) {
                            $Selection.TypeText("Monthly schedule is enabled.")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  Duration count in Months : $($BackupSchRP.MonthlySchedule.DurationCountInMonths)")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  RetentionType            : $($BackupSchRP.MonthlySchedule.RetentionScheduleFormatType)")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  RetentionDaily           : $($BackupSchRP.MonthlySchedule.RetentionScheduleDaily)")
                            $Selection.TypeParagraph()
                            $DayOfWeek = ArrayToLine $BackupSchRP.MonthlySchedule.DaysOfTheWeek
                            $Week = ArrayToLine $BackupSchRP.MonthlySchedule.WeeksOfTheMonth
                            $Selection.TypeText("  RetentionWeekly          : $DayOfWeek $Week")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.MonthlySchedule.RetentionTimes
                            $Selection.TypeText("  RetentionTime            : $Value")
                            $Selection.TypeParagraph()
                        }
                        else  {
                            $Selection.TypeText("Monthly schedule is disabled.")
                            $Selection.TypeParagraph()
                        }
                        if ($BackupSchRP.IsYearlyScheduleEnabled) {
                            $Selection.TypeText("Yearly schedule is enabled.")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  Duration count in Years  : $($BackupSchRP.YearlySchedule.DurationCountInYears)")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  RetentionType            : $($BackupSchRP.YearlySchedule.RetentionScheduleFormatType)")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  RetentionDaily           : $($BackupSchRP.YearlySchedule.RetentionScheduleDaily)")
                            $Selection.TypeParagraph()
                            $DayOfWeek = ArrayToLine $BackupSchRP.YearlySchedule.RetentionScheduleWeekly.DaysOfTheWeek
                            $Week = ArrayToLine $BackupSchRP.YearlySchedule.RetentionScheduleWeekly.WeeksOfTheMonth
                            $Selection.TypeText("  RetentionWeekly          : $DayOfWeek $Week")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.YearlySchedule.MonthsOfYear
                            $Selection.TypeText("  RetentionYearly          : $Value")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.YearlySchedule.RetentionTimes
                            $Selection.TypeText("  RetentionTime            : $Value")
                            $Selection.TypeParagraph()
                        }
                        else  {
                            $Selection.TypeText("Yearly schedule is disabled.")
                            $Selection.TypeParagraph()
                        }
                    }
                    "MSSQL" {
                        $Selection.TypeText("Full Backup.")
                        $Selection.TypeParagraph()
                        $BackupSchRP = $BackupSchedule.FullBackupRetentionPolicy
                        if ($BackupSchRP.IsDailyScheduleEnabled) {
                            $Selection.TypeText("Daily schedule enabled.")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  Duration count in Days   : $($BackupSchRP.DailySchedule.DurationCountInDays)")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.DailySchedule.RetentionTimes
                            $Selection.TypeText("  RetentionTime            : $Value")
                            $Selection.TypeParagraph()
                        }
                        else  {
                            $Selection.TypeText("Daily schedule is disabled.")
                            $Selection.TypeParagraph()
                        }
                        if ($BackupSchRP.IsWeeklyScheduleEnabled) { 
                            $Selection.TypeText("Weekly Schedule is enabled.")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  Duration count in Weeks  : $($BackupSchRP.WeeklySchedule.DurationCountInWeeks)")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.WeeklySchedule.RetentionTimes
                            $Selection.TypeText("  RetentionTime            : $Value")
                            $Selection.TypeParagraph()
                            $DayOfWeek = ArrayToLine $BackupSchRP.WeeklySchedule.DaysOfTheWeek
                            $Selection.TypeText("  Days of Week             : $DayOfWeek")
                            $Selection.TypeParagraph()
                        }
                        else  {
                            $Selection.TypeText("Weekly schedule is disabled.")
                            $Selection.TypeParagraph()
                        }
                        if ($BackupSchRP.IsMonthlyScheduleEnabled) {
                            $Selection.TypeText("Monthly schedule is enabled.")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  Duration count in Months : $($BackupSchRP.MonthlySchedule.DurationCountInMonths)")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  RetentionType            : $($BackupSchRP.MonthlySchedule.RetentionScheduleFormatType)")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  RetentionDaily           : $($BackupSchRP.MonthlySchedule.RetentionScheduleDaily)")
                            $Selection.TypeParagraph()
                            $DayOfWeek = ArrayToLine $BackupSchRP.MonthlySchedule.DaysOfTheWeek
                            $Week = ArrayToLine $BackupSchRP.MonthlySchedule.WeeksOfTheMonth
                            $Selection.TypeText("  RetentionWeekly          : $DayOfWeek $Week")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.MonthlySchedule.RetentionTimes
                            $Selection.TypeText("  RetentionTime            : $Value")
                            $Selection.TypeParagraph()
                        }
                        else  {
                            $Selection.TypeText("Monthly schedule is disabled.")
                            $Selection.TypeParagraph()
                        }
                        if ($BackupSchRP.IsYearlyScheduleEnabled) {
                            $Selection.TypeText("Yearly schedule is enabled.")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  Duration count in Years  : $($BackupSchRP.YearlySchedule.DurationCountInYears)")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  RetentionType            : $($BackupSchRP.YearlySchedule.RetentionScheduleFormatType)")
                            $Selection.TypeParagraph()
                            $Selection.TypeText("  RetentionDaily           : $($BackupSchRP.YearlySchedule.RetentionScheduleDaily)")
                            $Selection.TypeParagraph()
                            $DayOfWeek = ArrayToLine $BackupSchRP.YearlySchedule.RetentionScheduleWeekly.DaysOfTheWeek
                            $Week = ArrayToLine $BackupSchRP.YearlySchedule.RetentionScheduleWeekly.WeeksOfTheMonth
                            $Selection.TypeText("  RetentionWeekly          : $DayOfWeek $Week")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.YearlySchedule.MonthsOfYear
                            $Selection.TypeText("  RetentionYearly          : $Value")
                            $Selection.TypeParagraph()
                            $Value = ArrayToLine $BackupSchRP.YearlySchedule.RetentionTimes
                            $Selection.TypeText("  RetentionTime            : $Value")
                            $Selection.TypeParagraph()
                        }
                        else  {
                            $Selection.TypeText("Yearly schedule is disabled.")
                            $Selection.TypeParagraph()
                        }
                        if ($BackupSchedule.IsLogBackupEnabled) {
                            $Selection.TypeText("Log backup is enabled.")
                            $Selection.TypeParagraph()   
                            $Selection.TypeText("Log backups are made every $($BackupSchedule.LogBackupSchedulePolicy.ScheduleFrequencyInMins) minutes.")
                            $Selection.TypeParagraph() 
                            $RetentionCount = $BackupSchedule.LogBackupRetentionPolicy.RetentionCount  
                            $RetentionType = $BackupSchedule.LogBackupRetentionPolicy.RetentionDurationType  
                            $Selection.TypeText("Log backups are kept for $RetentionCount $RetentionType.")
                            $Selection.TypeParagraph()            
                        }
                        else {
                            $Selection.TypeText("Log backup is disabled.")
                            $Selection.TypeParagraph()
                        }
                        if ($BackupSchedule.IsDifferentialBackupEnabled) {
                            $Selection.TypeText("Differential backup is enabled.")
                            $Selection.TypeParagraph()          
                        }
                        else {
                            $Selection.TypeText("Differential backup is disabled.")
                            $Selection.TypeParagraph()
                        }                        
                    }
                    Default {}
                }
            }
        }
    }
    else {
        $Selection.TypeParagraph()
        $Selection.TypeText("No backup policies found.") 
        FindWordDocumentEnd
    }
    Write-Output "Creating replication job table."
    #Adding recovery information
    #Only get information from last 24h
    $ASRFabrics += Get-AzRecoveryServicesAsrFabric
    $startDate = (Get-Date).AddDays(-1)
    $endDate = Get-Date
    $Selection.InsertNewPage()
    $Selection.Style = $Heading2
    $Selection.TypeText("Replication")
    $Selection.TypeParagraph()
    $TableArray = [System.Collections.ArrayList]@()
    $BackupFailed = 0
    $BackupJobFailed = $null
    $ASRFabrics = [System.Collections.ArrayList]@()
    $ASRFabrics = Get-AzRecoveryServicesAsrFabric
    $CounterASRFabrics = 1
    $MaxCounterASRFabrics = $ASRFabrics.Count
    Foreach ($ASRFabric in $ASRFabrics) {  
        $ProcASRFabric = $CounterASRFabrics/$MaxCounterASRFabrics*100
        $ProcASRFabricString = $ProcASRFabric.ToString("0.00") 
        $ASRActivity = "ASR Fabric $($ASRFabric.FriendlyName) ($ProcASRFabricString%)"
        Write-Progress -ID 0 -Activity $ASRActivity -PercentComplete ($ProcASRFabric)
        $ProtectionContainers = Get-AzRecoveryServicesAsrProtectionContainer -Fabric $ASRFabric
        $Activity = $null
        if ($ProtectionContainers) {
            foreach ($ProtectionContainer in $ProtectionContainers) {
                $ProtectedItems = Get-AzRecoveryServicesAsrReplicationProtectedItem -ProtectionContainer $ProtectionContainer
                $CounterProtectedItems = 1
                $MaxCounterProtectedItems = $ProtectedItems.Count
                foreach ($ProtectedItem in $ProtectedItems) {
                    $ProcProtectedItems = $CounterProtectedItems/$MaxCounterProtectedItems*100
                    $ProcPProtectedItemString = $ProcProtectedItems.ToString("0.00")
                    $TableMember = New-Object System.Object;
                    $ReplVM = Get-AzRecoveryServicesAsrReplicationProtectedItem -FriendlyName $ProtectedItem.FriendlyName -ProtectionContainer $ProtectionContainer
                    $Activity = "Checking ASR $CounterProtectedItems/$MaxCounterProtectedItems for $($ReplVM.RecoveryAzureVMName) ($ProcPProtectedItemString%)"
                    Write-Progress -ID 1 -Activity $Activity -PercentComplete ($ProcProtectedItems)
                    $CounterProtectedItems++
                    $RecoveryPoints = Get-AzRecoveryServicesAsrRecoveryPoint -ReplicationProtectedItem $ReplVM
                    $LastRecoveryPoint=$RecoveryPoints[$RecoveryPoints.count-1]
                    $TableMember | Add-Member -type NoteProperty -name Location -Value $ReplVM.RecoveryFabricFriendlyName
                    $TableMember | Add-Member -type NoteProperty -name Server -Value $ReplVM.RecoveryAzureVMName
                    $TableMember | Add-Member -type NoteProperty -name State -Value $LastRecoveryPoint.RecoveryPointType
                    $TableMember | Add-Member -type NoteProperty -name RPTime -Value $LastRecoveryPoint.RecoveryPointTime
                    $TableMember | Add-Member -type NoteProperty -name Policy -Value $ReplVM.PolicyFriendlyName
                    $TableMember | Add-Member -type NoteProperty -name LastOKTest -Value $ReplVM.LastSuccessfulTestFailoverTime
                    $TableArray.Add($TableMember) | Out-Null
                }
            }   
        }
        if ($Activity) { Write-Progress -ID 1 -Activity $Activity -Status "Ready" -Completed }
        if ($Activity) { Write-Progress -ID 0 -Activity $ASRActivity -Status "Ready" -Completed }
    }
    FindWordDocumentEnd
    if ($TableArray){ 
        $TableArray = $TableArray | Sort-Object Location,Server
        $WordTable = AddWordTable -CustomObject $TableArray -Columns Location, Server, State, RPTime, Policy -Headers "Location", "Server", "Status", "Last Restore Point Time (UTC)", "Policy"
    }
    else {
        $Selection.TypeText("No replication jobs found.") 
        $Selection.TypeParagraph()
    }
    FindWordDocumentEnd
    $Selection.Style = $Heading3
    $Selection.TypeText("Replication polcies")
    $Selection.TypeParagraph()   
    if ($ASRPolicies) {
        $TableArray = [System.Collections.ArrayList]@()
        foreach ($ASRPolicy in $ASRPolicies) {
            $TableMember = New-Object System.Object;
            $TableMember | Add-Member -type NoteProperty -name Name -Value $ASRPolicy.Name
            $TableMember | Add-Member -type NoteProperty -name ACFM -Value $ASRPolicy.ReplicationProviderSettings.AppConsistentFrequencyInMinutes
            $TableMember | Add-Member -type NoteProperty -name CCFIM -Value $ASRPolicy.ReplicationProviderSettings.CrashConsistentFrequencyInMinutes
            $TableMember | Add-Member -type NoteProperty -name MVMSS -Value $ASRPolicy.ReplicationProviderSettings.MultiVmSyncStatus
            $TableMember | Add-Member -type NoteProperty -name RPH -Value $ASRPolicy.ReplicationProviderSettings.RecoveryPointHistory
            $TableMember | Add-Member -type NoteProperty -name RPTIM -Value $ASRPolicy.ReplicationProviderSettings.RecoveryPointThresholdInMinutes
            $TableArray.Add($TableMember) | Out-Null
        }
        $TableArray = $TableArray | Sort-Object Name
        $WordTable = AddWordTable -CustomObject $TableArray -Columns Name, ACFM, CCFIM, MVMSS, RPH, RPTIM -Headers "Name", "AppConst (Min)", "CrashConst (Min)", "MulitVMSync", "RP History", "RP Thres. (Min)"
    }
    else {
        
        $Selection.TypeText("No replication policies found.") 
        $Selection.TypeParagraph()
    }   
    FindWordDocumentEnd    
}

Write-Output "Creating Azure firewall table."

$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Azure firewall")
$Selection.TypeParagraph()
$TableArray = [System.Collections.ArrayList]@()

$FWs = get-azfirewall
$TableArray = [System.Collections.ArrayList]@()
foreach ($FW in $FWs) {
	$FWPolicy = Get-AzFirewallPolicy -ResourceId $fw.FirewallPolicy.id
	$FWRCGroups = $FWPolicy.RuleCollectionGroups
	foreach ($FWRCGroup in $FWRCGroups) {
		$Parts = $FWRCGroup.ID.Split("/")
		$FirewallPolicyRuleName =  $Parts[10]
		$FirewallResourceGroup = $Parts[4]
		$FirewallPolicyName = $Parts[8]
		$FirewallPRCG = Get-AzFirewallPolicyRuleCollectionGroup -AzureFirewallPolicyName $FirewallPolicyName -Name $FirewallPolicyRuleName -ResourceGroupName $FirewallResourceGroup
		$FWRuleCollections = $FirewallPRCG.Properties.RuleCollection
		Foreach ($FWRuleCollection in $FWRuleCollections) {
            foreach ($FWRule in $FWRuleCollection.Rules) {
				$TableMember = InitFirewallRule
				$TableMember | Add-Member -MemberType NoteProperty -name Firewall -Value $FW.Name -force
				$TableMember | Add-Member -MemberType NoteProperty -name FirewallPolicyName -Value $FirewallPolicyName -force
				$TableMember | Add-Member -MemberType NoteProperty -name FirewallPolRuleName -Value $FirewallPolicyRuleName -force
				$TableMember | Add-Member -MemberType NoteProperty -name FirewallRulePrio -Value $FWRuleCollection.Priority -force
                $TableMember | Add-Member -MemberType NoteProperty -name NetworkRuleCollection $FWRuleCollection.Name -force
				$TableMember | Add-Member -MemberType NoteProperty -Name ruleType -Value $FWRule.RuleType -force
				$Value = ArrayToLine $FWRule.Protocols
				$TableMember | Add-Member -MemberType NoteProperty -Name Protocols -Value $Value -force
				$Value = ArrayToLine $FWRule.DestinationPorts
				$TableMember | Add-Member -MemberType NoteProperty -Name DestinationPorts -Value $Value -force
				$TableMember | Add-Member -MemberType NoteProperty -Name FirewallRuleName -Value $FWRule.Name -force
				if ($FWRule.SourceAddresses) {
					$Value = ArrayToLine $FWRule.SourceAddresses
					$TableMember | Add-Member -MemberType NoteProperty -Name Source -Value $Value -force
				}
				if ($FWRule.SourceIpGroups) {
					$Value = GroupArrayToLine $FWRule.SourceIpGroups
					$TableMember | Add-Member -MemberType NoteProperty -Name Source -Value $Value -force
				}					
				if ($FWRule.DestinationAddresses) {
					$Value = ArrayToLine $FWRule.DestinationAddresses
					$TableMember | Add-Member -MemberType NoteProperty -Name Destination -Value $Value -force
				}
				if ($FWRule.DestinationIpGroups) {
					$Value = GroupArrayToLine $FWRule.DestinationIpGroups
					$TableMember | Add-Member -MemberType NoteProperty -Name Destination -Value $Value -force
				}			
				if ($FWRule.DestinationFqdns) {
					$Value = ArrayToLine $FWRule.DestinationFqdns
					$TableMember | Add-Member -MemberType NoteProperty -Name Destination -Value $Value -force
				}	
				if ($FWRule.TranslatedAddress) { $TableMember | Add-Member -MemberType NoteProperty -Name translated -Value $FWRule.TranslatedAddress -force }
				if ($FWRule.TranslatedFqdn) { $TableMember | Add-Member -MemberType NoteProperty -Name translated -Value $FWRule.TranslatedFqdn -force }
				if ($FWRule.translatedPort) { $TableMember | Add-Member -MemberType NoteProperty -Name translatedPort -Value $FWRule.translatedPort -force }
                $TableArray.Add($TableMember) | Out-Null
			}
		}
	}
}
if ($TableArray) {
    $DNATTable = $TableArray | where-object { $_.RuleType -eq "NatRule" }
    $RuleTable = $TableArray | where-object { $_.RuleType -ne "NatRule" }
    if ($DNATTable) {
        $Selection.Style = $Heading2
        $Selection.TypeText("Incoming Rules")
        $Selection.TypeParagraph()
        FindWordDocumentEnd
        $DNATTable = $DNATTable | Sort-Object Firewall,FirewallRulePrio
        $WordTable = AddWordTable -CustomObject $DNATTable -Columns Firewall, FirewallPolicyName, FirewallPolRuleName, NetworkRuleCollection, FirewallRuleName, FirewallRulePrio, Protocols, Source, Destination, DestinationPorts, translated, translatedPort -Headers "Firewall", "Policy Name", "Rule Collection Group", "Network Rule Collection","Rule Name", "Priority", "Protocols", "Source", "Destination", "Destination Ports", "Translated", "Translated Port"
        FindWordDocumentEnd
    }
    if ($RuleTable) {
        $Selection.TypeParagraph()
        $Selection.Style = $Heading2
        $Selection.TypeText("Outgoing Rules")
        $Selection.TypeParagraph()
        FindWordDocumentEnd
        $RuleTable = $RuleTable | Sort-Object Firewall, FirewallRulePrio
        $WordTable = AddWordTable -CustomObject $RuleTable -Columns Firewall, FirewallPolicyName, FirewallPolRuleName, NetworkRuleCollection, FirewallRuleName, FirewallRulePrio, Protocols, Source, Destination, DestinationPorts -Headers "Firewall", "Policy Name", "Rule Collection Group", "Network Rule Collection","Rule Name", "Priority", "Protocols", "Source", "Destination", "Destination Ports"
    }
}
Else { 
    $Selection.TypeText("No Azure firewall found.")    
    $Selection.TypeParagraph() 
}
FindWordDocumentEnd

Write-Output "Creating IP groups table."
$AllIPGroups = get-azipgroup
if ($ALLIPGroups) {
    $TableArray = [System.Collections.ArrayList]@()
    $Selection.Style = $Heading2
    $Selection.TypeText("Groups")
    $Selection.TypeParagraph()
    FindWordDocumentEnd
    $TableArray = [System.Collections.ArrayList]@()
    foreach ($IPGroup in $ALLIPGroups) {
        $TableMember = New-Object System.Object;
        $TableMember | Add-Member -type NoteProperty -name Name -Value $IPGroup.Name
        $IpAddressList = ""
        foreach ($IpAddress in $IPgroup.IpAddresses) {
            if ($IpAddressList) {
                $IpAddressList = $IpAddressList + "," + $IpAddress
            }
            else {
                $IpAddressList = $IpAddress
            } 
        }
        $TableMember | Add-Member -type NoteProperty -name Members -Value $IpAddressList
        $TableArray.Add($TableMember) | Out-Null
    }
    $TableArray  = $TableArray | Sort-Object Name
    $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, Members -Headers "Name", "Members"
    FindWordDocumentEnd
}
### Get NAT gateway information
Write-Output "Creating NAT gateway table."
$NatGateways = Get-AzNatGateway

if ($NatGateways) {
    $TableArray = [System.Collections.ArrayList]@()
    $Selection.Style = $Heading1
    $Selection.TypeText("NAT gateway")
    $Selection.TypeParagraph()
    FindWordDocumentEnd
    foreach ($NatGateway in $NatGateways) {
        $TableMember = New-Object System.Object;
        $TableMember | Add-Member -type NoteProperty -name Name -Value $NatGateway.Name
        $TableMember | Add-Member -type NoteProperty -name RG -Value $NatGateway.ResourceGroupName
        $Parts = $NatGateway.PublicIpAddresses.ID.Split("/")
        $Value = $Parts[8]
        $TableMember | Add-Member -type NoteProperty -name IP -Value $Value
        $TableArray.Add($TableMember) | Out-Null
    }
    $TableArray  = $TableArray | Sort-Object Name
    $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, RG, IP -Headers "Name", "ResourceGroup", "PublicIP"
    FindWordDocumentEnd
}
Write-Output "Creating Bastion table."
$Bastions = Get-AzBastion
$Selection.Style = $Heading1
$Selection.TypeText("Bastion")
$Selection.TypeParagraph()
FindWordDocumentEnd
if ($Bastions) {
    $TableArray = [System.Collections.ArrayList]@()
    foreach ($Bastion in $Bastions) {
        $TableMember = New-Object System.Object;
        $TableMember | Add-Member -type NoteProperty -name Name -Value $Bastion.Name
        $TableMember | Add-Member -type NoteProperty -name RG -Value $Bastion.ResourceGroupName
        $TableMember | Add-Member -type NoteProperty -name PrivAlloc -Value $Bastion.IpConfigurations.PrivateIpAllocationMethod
        $TableArray.Add($TableMember) | Out-Null
    }
    $TableArray  = $TableArray | Sort-Object Name
    $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, RG, PrivAlloc -Headers "Name", "ResourceGroup", "PrivAllocMethode"
}
else {
    $Selection.TypeText("No bastion found.")  
    $Selection.TypeParagraph() 
            
}
FindWordDocumentEnd 
$ALLVirtualNetworks = Get-AzVirtualNetwork
$Selection.Style = $Heading1
$Selection.TypeText("Network")
$Selection.TypeParagraph()
FindWordDocumentEnd 
if ($ALLVirtualNetworks) {
    $TableArray = [System.Collections.ArrayList]@()
    foreach ($VirtualNetwork in $ALLVirtualNetworks) {
        foreach ($Subnet in $VirtualNetwork.Subnets) {
            $TableMember = New-Object System.Object;
            $TableMember | Add-Member -type NoteProperty -name VNName -Value $VirtualNetwork.Name
            $TableMember | Add-Member -type NoteProperty -name SubName -Value $Subnet.Name
            $TableMember | Add-Member -type NoteProperty -name Address -Value $Subnet.AddressPrefix
            if ($Subnet.NatGateway.Id) {
                $Parts = $Subnet.NatGateway.Id.Split("/")
                $Value = $Parts[8]
                $TableMember | Add-Member -type NoteProperty -name NAT -Value $Value
            }
            else {
                $TableMember | Add-Member -type NoteProperty -name NAT -Value "None"
            }
            $TableArray.Add($TableMember) | Out-Null
        }
    }
    $TableArray  = $TableArray | Sort-Object Name
    $WordTable = AddWordTable -CustomObject $TableArray  -Columns VNName, RG, SubName, Address, NAT -Headers "Virtual Network", "ResourceGroup", "Subnet Name", "IP Address", "NAT"       
}
else {
    $Selection.TypeText("No network found.")        
    $Selection.TypeParagraph()   
}
FindWordDocumentEnd 
Write-Output "Getting load balancers."
$ALLLBs = Get-AzLoadBalancer
$Selection.Style = $Heading1
$Selection.TypeText("Loadbalancers")
$Selection.TypeParagraph()
FindWordDocumentEnd 
if ($ALLLBs) {
    $TableArray = [System.Collections.ArrayList]@()
    foreach ($LB in $ALLLBs) { 
        $TableMember = New-Object System.Object;
        $TableMember | Add-Member -type NoteProperty -name Name -Value $LB.Name
        $TableMember | Add-Member -type NoteProperty -name RG -Value $LB.ResourceGroupName
        $TableArray.Add($TableMember) | Out-Null
    }
    $TableArray  = $TableArray | Sort-Object Name
    $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, RG -Headers "LB Name", "ResourceGroup"
    FindWordDocumentEnd
    $Selection.TypeParagraph()
    foreach ($LB in $ALLLBs) {
        ## Add Heading for each NSG
        $Selection.Style = $Heading2
        $Selection.TypeText($LB.Name)
        $Selection.TypeParagraph()
        $TableArray = [System.Collections.ArrayList]@()
        $Selection.Style = $Heading3
        $Selection.TypeText("Frontend")
        $Selection.TypeParagraph()
        foreach ($FrontendIPConfig in  $LB.FrontendIpConfigurations) {
            $TableMember = New-Object System.Object;
            $TableMember | Add-Member -type NoteProperty -name Name -Value $FrontendIPConfig.Name
            $TableMember | Add-Member -type NoteProperty -name PrivIP -Value $FrontendIPConfig.PrivateIpAddress
            if ($FrontendIPConfig.PublicIpAddress) {
                $TableMember | Add-Member -type NoteProperty -name PubIP -Value $FrontendIPConfig.PublicIpAddress
            }
            else {
                $TableMember | Add-Member -type NoteProperty -name PubIP -Value "none"
            }
            $Parts = $FrontendIPConfig.subnet.ID.Split("/")
            $TableMember | Add-Member -type NoteProperty -name subnet -Value $Parts[8]
            $Parts = $FrontendIPConfig.loadbalancingrules.ID.Split("/")
            $TableMember | Add-Member -type NoteProperty -name LBRule -Value $Parts[10]
            $TableArray.Add($TableMember) | Out-Null
        }
        $TableArray  = $TableArray | Sort-Object Name
        $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, PrivIP, PubIP, subnet, LBRule -Headers "Name", "Private IP", "Public IP", "subnet", "LB Rule"
        FindWordDocumentEnd
        $TableArray = [System.Collections.ArrayList]@()
        $Selection.Style = $Heading3
        $Selection.TypeText("Backend")
        $Selection.TypeParagraph()
        foreach ($BackendAddressPool in $LB.BackendAddressPools) {
            foreach ($BackendIPConfig in $BackendAddressPool.BackendIpConfigurations) {
                $Parts = $BackendIPConfig.ID.Split("/")
                $VirtualMachineName = FindVM $Parts[8]
                $TableMember = New-Object System.Object;
                $TableMember | Add-Member -type NoteProperty -name Name -Value $BackendAddressPool.Name
                $TableMember | Add-Member -type NoteProperty -name Nic -Value $Parts[10]
                $TableMember | Add-Member -type NoteProperty -name VM -Value $VirtualMachineName
                if ($BackendIPConfig.PrivateIpAddress) {
                    $Value = $BackendIPConfig.PrivateIpAddress
                }
                else {
                    $Value = "none"
                }
                if ($value -eq "none") {
                    #BackendIPconfig has no data. Trying alternative mode
                    $Nic = Get-AzNetworkInterface -Name $Parts[8]
                    $Value = $nic.IpConfigurations[0].PrivateIpAddress
                }
                $TableMember | Add-Member -type NoteProperty -name PrivIP -Value $Value
                $TableArray.Add($TableMember) | Out-Null
            } 
        }
        $TableArray  = $TableArray | Sort-Object Name
        $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, PrivIP, Nic, VM -Headers "Name", "Private IP", "nic", "Virtual Machine"
        FindWordDocumentEnd
        $TableArray = [System.Collections.ArrayList]@()
        $Selection.Style = $Heading3
        $Selection.TypeText("Loadbalancer probes")
        $Selection.TypeParagraph()
        foreach ($LBProbe in $LB.Probes) {
            $TableMember = New-Object System.Object;
            $TableMember | Add-Member -type NoteProperty -name Name -Value $LBProbe.Name
            $TableMember | Add-Member -type NoteProperty -name Prot -Value $LBProbe.Protocol
            $TableMember | Add-Member -type NoteProperty -name Port -Value $LBProbe.Port
            $TableMember | Add-Member -type NoteProperty -name Int -Value $LBProbe.IntervalInSeconds
            $TableArray.Add($TableMember) | Out-Null
        }
        $TableArray  = $TableArray | Sort-Object Name
        $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, Prot, Port, Int -Headers "Name", "Protocol", "Port", "Interval (s)"
        FindWordDocumentEnd
        $TableArray = [System.Collections.ArrayList]@()
        $Selection.Style = $Heading3
        $Selection.TypeText("Loadbalancer rules")
        $Selection.TypeParagraph()
        foreach ($LBRule in $LB.loadbalancingrules) {
            $TableMember = New-Object System.Object;
            $TableMember | Add-Member -type NoteProperty -name Name -Value $LBRule.Name
            $TableMember | Add-Member -type NoteProperty -name Prot -Value $LBRule.Protocol
            $TableMember | Add-Member -type NoteProperty -name FP -Value $LBRule.FrontendPort
            $TableMember | Add-Member -type NoteProperty -name BP -Value $LBRule.BackendPort
            $TableMember | Add-Member -type NoteProperty -name Timeout -Value $LBRule.IdleTimeoutInMinutes
            $Parts = $LBRule.Probe.ID.Split("/")
            $TableMember | Add-Member -type NoteProperty -name Probe -Value $Parts[10]
            $TableArray.Add($TableMember) | Out-Null
        }
        $TableArray  = $TableArray | Sort-Object Name
        $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, Prot, Fp, BP, Timeout, Probe -Headers "Name", "Protocol", "Frontend Port", "Backend Port", "Timeout", "Probe"
    }
}
else {
    $Selection.TypeText("No load balancers found.") 
    $Selection.TypeParagraph()  
}
FindWordDocumentEnd 
Write-Output "Getting keyvault information."
$KeyVaults = Get-AzKeyVault
$Selection.Style = $Heading1
$Selection.TypeText("KeyVaults")
$Selection.TypeParagraph()
FindWordDocumentEnd
if ($KeyVaults) {
    $TableArray = [System.Collections.ArrayList]@()
    foreach ($KeyVault in  $KeyVaults) {
        $TableMember = New-Object System.Object; 
        $TableMember | Add-Member -type NoteProperty -name Name -Value $KeyVault.VaultName
        $TableMember | Add-Member -type NoteProperty -name Location -Value $KeyVault.Location
        $TableMember | Add-Member -type NoteProperty -name RGN -Value $KeyVault.ResourceGroupName
        $TableArray.Add($TableMember) | Out-Null
    }
    $TableArray  = $TableArray | Sort-Object Name
    $WordTable = AddWordTable -CustomObject $TableArray  -Columns Name, Location, RGN -Headers "Name", "Location", "ResourceGroup"
}
else {
    $Selection.TypeText("No keyvaults found.")  
    $Selection.TypeParagraph()
}
FindWordDocumentEnd
### Update the TOC now when all data has been written to the document 
$toc.Update()

# Save the document
Write-Output "Creating file $Report."
$Document.SaveAs([ref]$Report,[ref]$SaveFormat::wdFormatDocument)
$word.Quit()

# Free up memory
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word 
$elapsedTime = $(get-date) - $StartScriptTime
$Minutes = $elapsedTime.Minutes
$Seconds = $elapsedTime.Seconds
Write-Output "Script done in $Minutes Minute(s) and $Seconds Second(s)."
Write-Output "Script end."