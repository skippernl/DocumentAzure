<#
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
.NOTES
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
Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Document.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Selection.EndKey($wdStory,$wdMove) | Out-Null
}


$StartScriptTime = get-date 
Write-Output "Script Started."
#Connect-AzAccount | Out-Null
if (!($?)) {
    Write-Output "Error logging in to Azure."
    Write-Output "Script stopped."
    exit
}
if ($TenantId -and $SubscriptionId) {
    Select-AzSubscription -TenantId  $TenantId -SubscriptionId $SubscriptionId | Out-Null
    if (!($?)) {
        Write-Output "Unable to find Tennant or Subscription."
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
#$MediumShading1 = $AllStyles[38].Namelocal
## Add some text to start with
$Selection.Style = $Title
$Selection.TypeText("Azure Documentation for $Customer")
$Selection.TypeParagraph()
$Selection.TypeParagraph()

### Add the TOC
$range = $Selection.Range
$toc = $Document.TablesOfContents.Add($range)
$Selection.TypeParagraph()


###
### VIRTUAL MACHINES
###


## Add some text
$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Virtual Machines")
$Selection.TypeParagraph()

## Get all VMs from Azure
#Connect-AzAccount
Write-Output "Getting All Azure Resources"
$ALLAzureResources = Get-AzResource

Write-Output "Getting VM's"
$VMs = Get-AzVM -Status | Sort-Object Name
$TableArray = [System.Collections.ArrayList]@()

## Values
Write-Output "Creating VM table"
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
    $TableArray.Add($TableMember) | Out-Null
}

FindWordDocumentEnd
$WordTable = AddWordTable -CustomObject $TableArray -Columns VMName, Computername, RGN, Size, NIC, Status, Zone -Headers  "VM Name", "Computer name", "Resource Group Name", "VM Size", "Network Interface", "Power Status", "Zone"
FindWordDocumentEnd
$Selection.TypeParagraph()
$Selection.Style = $Heading2
$Selection.TypeText("Virtual Machine Disks")
$Selection.TypeParagraph()
Write-Output "Getting Disk information"
$Disks = get-Azdisk | Sort-Object Name

## Add a table for Disks
$TableArray = [System.Collections.ArrayList]@()

## Values
Write-Output "Creating Disk table"
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
$Selection.TypeText("Network Interfaces")
$Selection.TypeParagraph()

Write-Output "Getting Network interfaces"
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
Write-Output "Getting Reservations"
#If there are no reservations - Operation returned an invalid status code 'Forbidden' is being displayed
$ALLReservationOrders = Get-AzReservationOrder | Sort-Object Name

Write-Output "Creating Reservation table"
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
    $Selection.TypeText("No Reservations found.")  
}
FindWordDocumentEnd
$Selection.TypeParagraph()

########
######## Create a table for NSG
########

$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Network Security Groups")
$Selection.TypeParagraph()
Write-Output "Getting NSGs"
$NSGs = Get-AzNetworkSecurityGroup | Sort-Object Name

## Write NICs to NIC table 

Write-Output "Creating NSG table"
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

Write-Output "Creating Rule table"
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
        $TableMember | Add-Member -type NoteProperty -name SourcePort -Value $NSGRule.SourcePortRange
        $TableMember | Add-Member -type NoteProperty -name DestPort -Value $NSGRule.DestinationPortRange
        $TableMember | Add-Member -type NoteProperty -name SourcePref -Value $NSGRule.SourceAddressPrefix
        $TableMember | Add-Member -type NoteProperty -name DestPref -Value $NSGRule.DestinationAddressPrefix
        $TableMember | Add-Member -type NoteProperty -name Access -Value $NSGRule.Access
        $TableMember | Add-Member -type NoteProperty -name Prio -Value $NSGRule.Priority.ToString()
        $TableMember | Add-Member -type NoteProperty -name Direction -Value $NSGRule.Direction
        $TableArray.Add($TableMember) | Out-Null
    }

    ForEach ($NSGRULE in $NSGRulesDefault) {
        $TableMember = New-Object System.Object;
        $TableMember | Add-Member -type NoteProperty -name RuleName -Value $NSGRule.Name
        $TableMember | Add-Member -type NoteProperty -name Protocol -Value $NSGRule.Protocol
        $TableMember | Add-Member -type NoteProperty -name SourcePort -Value $NSGRule.SourcePortRange
        $TableMember | Add-Member -type NoteProperty -name DestPort -Value $NSGRule.DestinationPortRange
        $TableMember | Add-Member -type NoteProperty -name SourcePref -Value $NSGRule.SourceAddressPrefix
        $TableMember | Add-Member -type NoteProperty -name DestPref -Value $NSGRule.DestinationAddressPrefix
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
$Selection.TypeText("VPN Information")
$Selection.TypeParagraph()
$Selection.Style = $Heading2
$Selection.TypeText("VPN Gateway Connections")
$Selection.TypeParagraph()
Write-Output "Getting VPN Gateway Connections"
$NetworkConnections = $ALLAzureResources | Where-Object {$_.ResourceType -eq "Microsoft.Network/connections" } | Sort-Object Name
Foreach ($NetworkConnection in $NetworkConnections) {
    $NSG = Get-AzVirtualNetworkGatewayConnection -ResourceName $NetworkConnection.ResourceName -ResourceGroupName $NetworkConnection.ResourceGroupName
    $NetworkGatewayConnections.Add($NSG) | Out-Null
}
$NetworkGatewayConnections = $NetworkGatewayConnections | Sort-Object Name
$TableArray = [System.Collections.ArrayList]@()

########
######## Create a table for VPN GatewayConnections
########
## Values
Write-Output "Creating VPN table"
Foreach ($NGC in $NetworkGatewayConnections) {
    $TableMember = New-Object System.Object;
    $ResourceGroupName = $NGC.ResourceGroupName
    $Parts = $NGC.VirtualNetworkGateway1.id.Split("/")
    $AzEndpoint = $Parts[8]
    $Parts = $NGC.LocalNetworkGateway2.id.Split("/")
    $LocalEndpoint = $Parts[8]
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

$Selection.Style = $Heading2
$Selection.TypeText("VPN Local Gateways")
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
Write-Output "Creating VPN LocalGateway table"
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
    $Selection.TypeText("No Local gateway found.")  
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
Write-Output "Creating Public IP table"
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
        $Endpoint = "Unused"
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
$Selection.TypeText("Azure Backups")
$Selection.TypeParagraph()
$Vaults = Get-AzRecoveryServicesVault | Sort-Object Name
########
######## Create a table the backupjobs found
########

## Values
#Get Only Restore points of the last week.
$startDate = (Get-Date).AddDays(-7)
$endDate = Get-Date
Write-Output "Creating Backup job table"
$TableArray = [System.Collections.ArrayList]@()
$BackupFailed = 0
$BackupJobFailed = $null
$CounterVault = 0
$MaxCounterVault=$Vaults.Count
Foreach ($Vault in $Vaults) {
    $ProcVault = $CounterVault/$MaxCounterVault*100
    $ProcVaultString = $ProcVault.ToString("0.00")
    Write-Progress -ID 0 -Activity "Parsing vault $($Vault.Name) ($ProcVaultString%)" -PercentComplete ($ProcVault)
    $BackupJobs = Get-AzRecoveryServicesBackupJob -VaultId $Vault.ID
    $namedContainerVMs = Get-AzRecoveryServicesBackupContainer  -ContainerType "AzureVM" -Status "Registered" -VaultId $Vault.ID
    $CounterVault++
    $CounterBackupJob = 0
    $MaxCounterBackupJob = $BackupJobs.Count
    foreach ($BackupJob in $BackupJobs) {
        $TableMember = New-Object System.Object;
        
        $ProcBackup = $CounterBackupJob/$MaxCounterBackupJob*100
        $ProcBackupString = $ProcBackup.ToString("0.00")
        Write-Progress -ID 1 -Activity "Parsing backup job $($BackupJob.WorkloadName.ToUpper()) ($ProcBackupString%)" -PercentComplete ($ProcBackup)
        #Write-Output "Getting restore points for Vault $($Vault.Name), Job $($BackupJob.WorkloadName.ToUpper())."
        #There can be multiple Restore Points due to the fact that there could be more Jobs (after changing resource group etc)
        $CounterBackupJob++
        $rp = @()
        switch ($BackupJob.BackupManagementType) {
            "AzureVM" {
                foreach ($namedContainer in $namedContainerVMs) {
                    #Friendly name can be in multiple namedcontainers
                    if ($BackupJob.workloadname.ToUpper() -eq $namedContainer.FriendlyName.ToUpper()) {
                        $BackupnamedContainer = $namedContainer
                        $backupitem = Get-AzRecoveryServicesBackupItem -Container $BackupnamedContainer  -WorkloadType $BackupJob.BackupManagementType -VaultId $Vault.ID
                        $rp += Get-AzRecoveryServicesBackupRecoveryPoint -Item $backupitem -StartDate $startdate.ToUniversalTime() -EndDate $enddate.ToUniversalTime() -VaultId $Vault.ID
                    }
                }
                $WorkloadName = $BackupJob.workloadname.ToUpper()
                
                if ($rp) {
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
                $WorkloadName = $SQLServer.ToUpper() + " " + $SQLDatabase.ToUpper()
                foreach ($backupitem in $bkpItems) {
                    $ServerNameArray=$backupitem.ServerName.Split(".")
                    if ($ServerNameArray[0] -eq $SQLServer){ 
                        $rp = Get-AzRecoveryServicesBackupRecoveryPoint -Item $backupitem -StartDate $startdate.ToUniversalTime() -EndDate $enddate.ToUniversalTime() -VaultId $Vault.ID
                        if ($rp) {
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
                    }
                }  
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
}
FindWordDocumentEnd
if ($TableArray){ 
    $TableArray = $TableArray | Sort-Object Name, Workload
    $WordTable = AddWordTable -CustomObject $TableArray -Columns Name, Workload, Status, StartTime, EndTime, RP -Headers "Vault", "Backup Item", "Status", "Start Time (UTC)", "End Time (UTC)", "Latest RestorePoint (UTC)"
    FindWordDocumentEnd
    $Selection.TypeParagraph()
    switch ($backupFailed) {
        1 { 
            $Selection.TypeText("One failed backup where found!")
            $Selection.TypeText("The job that failed is: $BackupJobFailed.")   
            $Selection.TypeParagraph()
        }
        2 {
            $Selection.TypeText("Two or more failed backups where found!")
            $Selection.TypeText("The jobs that failed are $BackupJobFailed.")   
            $Selection.TypeParagraph()
        }
        Default {}
    }
    if ($BackupFailed -eq 1) {

    }
    else {

    }
}
else {
    $Selection.TypeParagraph()
    $Selection.TypeText("No Backups found.")  
}
FindWordDocumentEnd
$Selection.TypeParagraph()

Write-Output "Creating Replication job table"
#Adding recovery information
#Only get information from last 24h
$startDate = (Get-Date).AddDays(-1)
$endDate = Get-Date
$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Azure Replication")
$Selection.TypeParagraph()
$TableArray = [System.Collections.ArrayList]@()
$BackupFailed = 0
$BackupJobFailed = $null
$CounterVault = 0
$MaxCounterVault=$Vaults.Count
Foreach ($Vault in $Vaults) {   
    $ProcVault = $CounterVault/$MaxCounterVault*100
    $ProcVaultString = $ProcVault.ToString("0.00")
    Write-Progress -ID 0 -Activity "Parsing vault $($Vault.Name) ($ProcVaultString%)" -PercentComplete ($ProcVault)
    $BackupJobs = Get-AzRecoveryServicesBackupJob -VaultId $Vault.ID
    $namedContainerVMs = Get-AzRecoveryServicesBackupContainer  -ContainerType "AzureVM" -Status "Registered" -VaultId $Vault.ID
    $CounterVault++
    Set-AzRecoveryServicesAsrVaultContext -Vault $vault | Out-Null
    $ReplicationJobs = Get-AzRecoveryServicesAsrJob -StartTime $startdate.ToUniversalTime() -EndTime $enddate.ToUniversalTime()

    #Write-Output "Getting restore points for Vault $($Vault.Name), Job $($BackupJob.WorkloadName.ToUpper())."
    #There can be multiple Restore Points due to the fact that there could be more Jobs (after changing resource group etc)
   
    if ($ReplicationJobs) {
        $CounterReplJob = 0
        $MaxCounterReplJob=$ReplicationJobs.Count
        foreach ($ReplicationJob in $ReplicationJobs) {
            $TableMember = New-Object System.Object;

            $ProcRepl = $CounterReplJob/$MaxCounterReplJob*100
            $ProcReplString = $ProcRepl.ToString("0.00")
            Write-Progress -ID 1 -Activity "Parsing replication job $($ReplicationJob.TargetObjectName) ($ProcReplString%)" -PercentComplete ($ProcRepl)
            $CounterReplJob++
            $TableMember | Add-Member -type NoteProperty -name Vault -Value $Vault.Name
            $TableMember | Add-Member -type NoteProperty -name Server -Value $ReplicationJob.TargetObjectName
            $TableMember | Add-Member -type NoteProperty -name JobType -Value $ReplicationJob.JobType
            $TableMember | Add-Member -type NoteProperty -name State -Value $ReplicationJob.State
            $TableMember | Add-Member -type NoteProperty -name StartTime -Value $ReplicationJob.StartTime
            $TableMember | Add-Member -type NoteProperty -name EndTime -Value $ReplicationJob.EndTime
            $TableArray.Add($TableMember) | Out-Null
        }
    }
}
FindWordDocumentEnd
if ($TableArray){ 
    $TableArray = $TableArray | Sort-Object Vault, Server, JobType
    $WordTable = AddWordTable -CustomObject $TableArray -Columns Vault, Server, JobType, State, StartTime, EndTime -Headers "Vault", "Server", "JobType", "Status", "Start Time (UTC)", "End Time (UTC)"
    FindWordDocumentEnd
    $Selection.TypeParagraph()
}
else {
    $Selection.TypeParagraph()
    $Selection.TypeText("No Replication Jobs found.")  
}
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