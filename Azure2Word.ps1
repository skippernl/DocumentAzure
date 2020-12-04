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

Uses modules AZ and Az.Reservations
Install-module -Name az
Install-Module -Name Az.Accounts -RequiredVersion 1.9.2
Install-Module -Name Az.Reservations
Last Modified: 2020/12/2
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
#Get-AzVirtualNetworkGatewayConnection
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

Write-Host "Script Started."
Connect-AzAccount | Out-Null
if ($TenantId -and $SubscriptionId) {
    Select-AzSubscription -TenantId  $TenantId -SubscriptionId $SubscriptionId | Out-Null
    if (!($?)) {
        Write-Host "Unable to find Tennant or Subscription."
        Write-Host "Script stopped."
        exit
    }
}
$NetworkGatewayConnections = [System.Collections.ArrayList]@()
$LocalVPNEndpoints = [System.Collections.ArrayList]@()
$LocalGatewayArray = [System.Collections.ArrayList]@()
$Report = "$ReportPath\$Customer-Azure.docx"
# Creating the Word Object, set Word to not be visual and add document
$Word = New-Object -ComObject Word.Application
$Word.Visible = $false
$Document = $Word.Documents.Add()
$Selection = $Word.Selection
Write-Host "Getting Word information."
$ALLStyles = $document.Styles | Select-Object NameLocal 
$Title = $AllStyles[360].Namelocal
$Heading1 = $AllStyles[149].Namelocal
$Heading2 = $AllStyles[150].Namelocal
$MediumShading1 = $AllStyles[38].Namelocal
## Add some text to start with
$Selection.Style = $Title
$Selection.TypeText("Azure Documentation for $Customer")
$Selection.TypeParagraph()

### Add the TOC
$range = $Selection.Range
$toc = $Document.TablesOfContents.Add($range)
$Selection.TypeParagraph()


###
### VIRTUAL MACHINES
###


## Add some text
$Selection.Style = $Heading1
$Selection.TypeText("Virtual Machines")
$Selection.TypeParagraph()

## Get all VMs from Azure
#Connect-AzAccount
Write-Host "Getting All Azure Resources"
$ALLAzureResources = Get-AzResource

Write-Host "Getting VM's"
$VMs = Get-AzVM | Sort-Object Name

## Add a table for VMs
$VMTable = $Selection.Tables.add($Word.Selection.Range, $VMs.Count + 1, 5,
[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

$VMTable.Style = $MediumShading1
$VMTable.Cell(1,1).Range.Text = "Name"
$VMTable.Cell(1,2).Range.Text = "Computer Name"
$VMTable.Cell(1,3).Range.Text = "VM Size"
$VMTable.Cell(1,4).Range.Text = "Resource Group Name"
$VMTable.Cell(1,5).Range.Text = "Network Interface"

## Values
$i=2
Write-Host "Creating VM table"
Foreach ($VM in $VMs) {

        $VMName = $VM.NetworkProfile.NetworkInterfaces.ID
        $Parts = $VMName.Split("/")
        $NICLabel = $PArts[8]

    $VMTable.cell(($i),1).range.Bold = 0
    $VMTable.cell(($i),1).range.text = $VM.Name
    $VMTable.cell(($i),2).range.Bold = 0
    if ($null -eq $VM.OSProfile.ComputerName) {
        #https://docs.microsoft.com/en-us/troubleshoot/azure/virtual-machines/computer-names-missing-blank
        $VirtualMachineName = "(" + $VM.Name + ")"
    }
    else {
        $VirtualMachineName = $VM.OSProfile.ComputerName
    }
    $VMTable.cell(($i),2).range.text = $VirtualMachineName
    $VMTable.cell(($i),3).range.Bold = 0
    $VMTable.cell(($i),3).range.text = $VM.HardwareProfile.VmSize
    $VMTable.cell(($i),4).range.Bold = 0
    $VMTable.cell(($i),4).range.text = $VM.ResourceGroupName
    $VMTable.cell(($i),5).range.Bold = 0
    $VMTable.cell(($i),5).range.text = $NICLabel

$i++
}


$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()
$Selection.Style = $Heading2
$Selection.TypeText("Virtual Machine Disks")
$Selection.TypeParagraph()
Write-Host "Getting Disk information"
$Disks = get-Azdisk | Sort-Object Name

## Add a table for Disks
$DiskTable = $Selection.Tables.add($Word.Selection.Range, $Disks.Count + 1, 6,
[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

$DiskTable.Style = $MediumShading1
$DiskTable.Cell(1,1).Range.Text = "DiskName"
$DiskTable.Cell(1,2).Range.Text = "Servername"
$DiskTable.Cell(1,3).Range.Text = "DiskIOPSReadWrite"
$DiskTable.Cell(1,4).Range.Text = "DiskMBpsReadWrite"
$DiskTable.Cell(1,5).Range.Text = "Resource Group Name"
$DiskTable.Cell(1,6).Range.Text = "DiskSizeGB"

## Values
$i=2
Write-Host "Creating Disk table"
Foreach ($Disk in $Disks) {

        $DiskID = $Disk.ManagedBy
        if ($null -eq $DiskID) {
            $Server = "Unknown"
        }
        else {
            $Parts = $DiskID.Split("/")
            $Server = $Parts[8]
        }

    $DiskTable.cell(($i),1).range.Bold = 0
    $DiskTable.cell(($i),1).range.text = $Disk.Name
    $DiskTable.cell(($i),2).range.Bold = 0
    $DiskTable.cell(($i),2).range.text = $Server
    $DiskTable.cell(($i),3).range.Bold = 0
    $DiskTable.cell(($i),3).range.text = $Disk.DiskIOPSReadWrite.ToString()
    $DiskTable.cell(($i),4).range.Bold = 0
    $DiskTable.cell(($i),4).range.text = $Disk.DiskMBpsReadWrite.ToString()
    $DiskTable.cell(($i),5).range.Bold = 0
    $DiskTable.cell(($i),5).range.text = $Disk.ResourceGroupName
    $DiskTable.cell(($i),6).range.Bold = 0
    $DiskTable.cell(($i),6).range.text = $Disk.DiskSizeGB.ToString()

$i++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

########
######## NETWORK INTERFACE
########

$Selection.Style = $Heading2
$Selection.TypeText("Network Interfaces")
$Selection.TypeParagraph()

Write-Host "Getting Network interfaces"
$NICs = Get-AzNetworkInterface | Sort-Object Name

$NICTable = $Selection.Tables.add($Word.Selection.Range, $NICs.Count + 1, 7,
[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

$NICTable.Style = $MediumShading1
$NICTable.Cell(1,1).Range.Text = "Virtual Machine"
$NICTable.Cell(1,2).Range.Text = "Network Card Name"
$NICTable.Cell(1,3).Range.Text = "Resource Group Name"
$NICTable.Cell(1,4).Range.Text = "VNET"
$NICTable.Cell(1,5).Range.Text = "Subnet"
$NICTable.Cell(1,6).Range.Text = "Private IP Address"
$NICTable.Cell(1,7).Range.Text = "Private IP Allocation Method"

## Write NICs to NIC table 
$i=2

Write-Host "Creating NIC table"
Foreach ($NIC in $NICs) {

## Get connected VM, if there is one connected to the network interface
If (!$NIC.VirtualMachine.id) 
    { $VMLabel = " "}
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

    $NICTable.cell(($i),1).range.Bold = 0
    $NICTable.cell(($i),1).range.text = $VMLabel
    $NICTable.cell(($i),2).range.Bold = 0
    $NICTable.cell(($i),2).range.text = $NIC.Name
    $NICTable.cell(($i),3).range.Bold = 0
    $NICTable.cell(($i),3).range.text = $NIC.ResourceGroupName
    $NICTable.cell(($i),4).range.Bold = 0
    $NICTable.cell(($i),4).range.text = $VNETNAME 
    $NICTable.cell(($i),5).range.Bold = 0
    $NICTable.cell(($i),5).range.text = $SUBNETNAME
    $NICTable.cell(($i),6).range.Bold = 0   
    $NICTable.cell(($i),6).range.text = $NIC.IPconfigurations.PrivateIpAddress
    $NICTable.cell(($i),7).range.Bold = 0
    $NICTable.cell(($i),7).range.text = $NIC.IPconfigurations.PrivateIpAllocationMethod


$i++
}

$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()


##Find reservations
$Selection.Style = $Heading2
$Selection.TypeText("Reservations")
$Selection.TypeParagraph()
Write-Host "Getting Reservations"
$ALLReservations = Get-AzReservationOrder | Sort-Object Name
$ReservationTable = $Selection.Tables.add($Word.Selection.Range, $ALLReservations.Count + 1, 4,
[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

$ReservationTable.Style = $MediumShading1
$ReservationTable.Cell(1,1).Range.Text = "DisplayName"
$ReservationTable.Cell(1,2).Range.Text = "Start"
$ReservationTable.Cell(1,3).Range.Text = "Term"
$ReservationTable.Cell(1,4).Range.Text = "End"

$i=2

Write-Host "Creating Reservation table"
Foreach ($Reservation in $ALLReservations) {
    $ReservationTable.cell(($i),1).range.Bold = 0
    $ReservationTable.cell(($i),1).range.text = $Reservation.DisplayName
    $ReservationTable.cell(($i),2).range.Bold = 0
    $StartTime = $Reservation.CreatedDateTime
    $ReservationTable.cell(($i),2).range.text = $StartTime.ToString()
    $ReservationTable.cell(($i),3).range.Bold = 0
    $Term = $Reservation.Term
    $ReservationTable.cell(($i),3).range.text = $Term
    $Duration = $Term.substring(1,1)
    $LastChar = $Term.substring(2,1)
    #Using switch to be flexible
    switch ($LastChar) {
        "Y" {
            $EndTime = $StartTime.AddYears($Duration)
        }
    }
    $ReservationTable.cell(($i),4).range.Bold = 0
    $ReservationTable.cell(($i),4).range.text = $EndTime.ToString()
    $i++
}

$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()
########
######## Create a table for NSG
########

$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Network Security Groups")
$Selection.TypeParagraph()
Write-Host "Getting NSGs"
$NSGs = Get-AzNetworkSecurityGroup | Sort-Object Name

$NSGTable = $Selection.Tables.add($Word.Selection.Range, $NSGs.Count + 1, 4,
[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)

$NSGTable.Style = $MediumShading1
$NSGTable.Cell(1,1).Range.Text = "NSG Name"
$NSGTable.Cell(1,2).Range.Text = "Resource Group Name"
$NSGTable.Cell(1,3).Range.Text = "Network Interfaces"
$NSGTable.Cell(1,4).Range.Text = "Subnets"

## Write NICs to NIC table 
$i=2

Write-Host "Creating NSG table"
Foreach ($NSG in $NSGs) {

## Get connected NIC, if there is one connected 
If (!$NSG.NetworkInterfaces.Id) 
    { $NICLabel = " "}
Else
    {
        $SubnetName = $NSG.NetworkInterfaces.Id
        $Parts = $SubnetName.Split("/")
        $NICLabel = $PArts[8]
    }



## Get connected SUBNET, if there is one connected 
If (!$NSG.Subnets.Id) 
    { $SubnetLabel = " "}
Else
    {
        $SUBNETName = $NSG.Subnets.Id
        $Parts = $SUBNETName.Split("/")
        $SUBNETLabel = $PArts[10]
      }


    $NSGTable.cell(($i),1).range.Bold = 0
    $NSGTable.cell(($i),1).range.text = $NSG.Name
    $NSGTable.cell(($i),2).range.Bold = 0
    $NSGTable.cell(($i),2).range.text = $NSG.ResourceGroupName
    $NSGTable.cell(($i),3).range.Bold = 0
    $NSGTable.cell(($i),3).range.text = $NICLabel
    $NSGTable.cell(($i),4).range.Bold = 0
    $NSGTable.cell(($i),4).range.text = $SUBNETLabel

$i++
}

$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

########
######## Create a table for each NSG
########


### Get all NSGs
#$NSGs = Get-AzNetworkSecurityGroup
Write-Host "Creating Rule table"
ForEach ($NSG in $NSGs) {

    ## Add Heading for each NSG
    $Selection.Style = $Heading2
    $Selection.TypeText($NSG.Name)
    $Selection.TypeParagraph()

		$NSGRulesCustom = Get-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $NSG | Sort-Object Name
		$NSGRulesDefault = Get-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $NSG -DefaultRules | Sort-Object Name
		$NSGRuleCount = $NSGRulesCustom.Count + $NSGRulesDefault.Count
        ### Add a table for each NSG, the NSg has custom rules
        $NSGRuleTable = $Selection.Tables.add($Word.Selection.Range, $NSGRuleCount + 1, 9,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
           )

        $NSGRuleTable.Style = $MediumShading1
        $NSGRuleTable.Cell(1,1).Range.Text = "Rule Name"
        $NSGRuleTable.Cell(1,2).Range.Text = "Protocol"
        $NSGRuleTable.Cell(1,3).Range.Text = "Source Port Range"
        $NSGRuleTable.Cell(1,4).Range.Text = "Destination Port Range"
        $NSGRuleTable.Cell(1,5).Range.Text = "Source Address Prefix"
        $NSGRuleTable.Cell(1,6).Range.Text = "Destination Address Prefix"
        $NSGRuleTable.Cell(1,7).Range.Text = "Access"
        $NSGRuleTable.Cell(1,8).Range.Text = "Priority"
        $NSGRuleTable.Cell(1,9).Range.Text = "Direction"


        ### Get all custom Security Rules in the NSG
        #$NSGRules = Get-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $NSG
        $i = 2

        ForEach ($NSGRULE in $NSGRulesCustom) {
                $NSGRuleTable.cell(($i),1).range.Bold = 0
                $NSGRuleTable.cell(($i),1).range.text = $NSGRule.Name
                $NSGRuleTable.cell(($i),2).range.Bold = 0
                $NSGRuleTable.cell(($i),2).range.text = $NSGRule.Protocol

                $NSGRuleTable.cell(($i),3).range.Bold = 0
                $NSGRuleTable.cell(($i),3).range.text = ConvertArrayToLine $NSGRule.SourcePortRange

                $NSGRuleTable.cell(($i),4).range.Bold = 0
                $NSGRuleTable.cell(($i),4).range.text = ConvertArrayToLine $NSGRule.DestinationPortRange
 
                $NSGRuleTable.cell(($i),5).range.Bold = 0
                $NSGRuleTable.cell(($i),5).range.text = ConvertArrayToLine $NSGRule.SourceAddressPrefix

                $NSGRuleTable.cell(($i),6).range.Bold = 0
                $NSGRuleTable.cell(($i),6).range.text = ConvertArrayToLine $NSGRule.DestinationAddressPrefix

                $NSGRuleTable.cell(($i),7).range.Bold = 0
                $NSGRuleTable.cell(($i),7).range.text = $NSGRule.Access

                $NSGRuleTable.cell(($i),8).range.Bold = 0
                $NSGRuleTable.cell(($i),8).range.text = [string]$NSGRule.Priority

                $NSGRuleTable.cell(($i),9).range.Bold = 0
                $NSGRuleTable.cell(($i),9).range.text = $NSGRule.Direction

                $i++
            }

             ### Get all default Security Rules in the NSG
        #$NSGRules = Get-AzNetworkSecurityRuleConfig -NetworkSecurityGroup $NSG -DefaultRules
        ForEach ($NSGRULE in $NSGRulesDefault) {
                $NSGRuleTable.cell(($i),1).range.Bold = 0
                $NSGRuleTable.cell(($i),1).range.text = $NSGRule.Name
                $NSGRuleTable.cell(($i),2).range.Bold = 0
                $NSGRuleTable.cell(($i),2).range.text = $NSGRule.Protocol

                $NSGRuleTable.cell(($i),3).range.Bold = 0
                $NSGRuleTable.cell(($i),3).range.text = ConvertArrayToLine $NSGRule.SourcePortRange

                $NSGRuleTable.cell(($i),4).range.Bold = 0
                $NSGRuleTable.cell(($i),4).range.text = ConvertArrayToLine $NSGRule.DestinationPortRange

                $NSGRuleTable.cell(($i),5).range.Bold = 0
                $NSGRuleTable.cell(($i),5).range.text = ConvertArrayToLine $NSGRule.SourceAddressPrefix

                $NSGRuleTable.cell(($i),6).range.Bold = 0
                $NSGRuleTable.cell(($i),6).range.text = ConvertArrayToLine $NSGRule.DestinationAddressPrefix
 
                $NSGRuleTable.cell(($i),7).range.Bold = 0
                $NSGRuleTable.cell(($i),7).range.text = $NSGRule.Access

                $NSGRuleTable.cell(($i),8).range.Bold = 0
                $NSGRuleTable.cell(($i),8).range.text = [string]$NSGRule.Priority

                $NSGRuleTable.cell(($i),9).range.Bold = 0
                $NSGRuleTable.cell(($i),9).range.text = $NSGRule.Direction

                $i++
            }

            ### Close the NSG table
            $Word.Selection.Start= $Document.Content.End
            $Selection.TypeParagraph()

        }



##Get al Azure VPN

$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("VPN Information")
$Selection.TypeParagraph()
$Selection.Style = $Heading2
$Selection.TypeText("VPN GatewayConnections")
$Selection.TypeParagraph()
Write-Host "Getting VPN GatewayConnections"
$NetworkConnections = $ALLAzureResources | Where-Object {$_.ResourceType -eq "Microsoft.Network/connections" } | Sort-Object Name
Foreach ($NetworkConnection in $NetworkConnections) {
    $NSG = Get-AzVirtualNetworkGatewayConnection -ResourceName $NetworkConnection.ResourceName -ResourceGroupName $NetworkConnection.ResourceGroupName
    $NetworkGatewayConnections.Add($NSG) | Out-Null
}
$NetworkGatewayConnections = $NetworkGatewayConnections | Sort-Object Name
########
######## Create a table for VPN GatewayConnections
########
$NGCTable = $Selection.Tables.add($Word.Selection.Range, $NetworkGatewayConnections.Count + 2, 7,
[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent)

$NGCTable.Style = $MediumShading1
$NGCTable.Cell(1,1).Range.Text = "VPN"
$NGCTable.Cell(1,2).Range.Text = "ResourceGroup"
$NGCTable.Cell(1,3).Range.Text = "AzureEndpoint"
$NGCTable.Cell(1,4).Range.Text = "LocalEndpoint"
$NGCTable.Cell(1,5).Range.Text = "Status"
$NGCTable.Cell(1,6).Range.Text = "EgressBytesTransferredGB"
$NGCTable.Cell(1,7).Range.Text = "IngressBytesTransferredGB"

## Values
$i=2
Write-Host "Creating VPN table"
Foreach ($NGC in $NetworkGatewayConnections) {


    $NGCTable.cell(($i),1).range.Bold = 0
    $NGCTable.cell(($i),1).range.text = $NGC.Name
    $NGCTable.cell(($i),2).range.Bold = 0
    $ResourceGroupName = $NGC.ResourceGroupName
    $NGCTable.cell(($i),2).range.text = $ResourceGroupName
    if (!($LocalVPNEndpoints.Contains($ResourceGroupName))) { $LocalVPNEndpoints.Add($ResourceGroupName) | Out-Null }
    $NGCTable.cell(($i),3).range.Bold = 0
    $Parts = $NGC.VirtualNetworkGateway1.id.Split("/")
    $Endpoint = $Parts[8]
    $NGCTable.cell(($i),3).range.text = $Endpoint
    $NGCTable.cell(($i),4).range.Bold = 0
    $Parts = $NGC.LocalNetworkGateway2.id.Split("/")
    $Endpoint = $Parts[8]
    $NGCTable.cell(($i),4).range.text = $Endpoint
    $NGCTable.cell(($i),5).range.Bold = 0
    $NGCTable.cell(($i),5).range.text = $NGC.ConnectionStatus
    $NGCTable.cell(($i),6).range.Bold = 0
    $DataGB = $NGC.EgressBytesTransferred/1048576
    $DataGB = [math]::Round($DataGB)
    $NGCTable.cell(($i),6).range.text = $DataGB.ToString()
    $NGCTable.cell(($i),7).range.Bold = 0
    $DataGB = $NGC.IngressBytesTransferred/1048576
    $DataGB = [math]::Round($DataGB)
    $NGCTable.cell(($i),7).range.text = $DataGB.ToString()
$i++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

$Selection.Style = $Heading2
$Selection.TypeText("VPN LocalGateways")
$Selection.TypeParagraph()
foreach ($LocalVPNEndpoint in $LocalVPNEndpoints) {
    $LocalGateway = Get-AzLocalNetworkGateway -ResourceGroupName $LocalVPNEndpoint
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
$LocalGatewayTable = $Selection.Tables.add($Word.Selection.Range, $LocalGatewayArray.Count + 1, 4,
[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent)

$LocalGatewayTable.Style = $MediumShading1
$LocalGatewayTable.Cell(1,1).Range.Text = "Name"
$LocalGatewayTable.Cell(1,2).Range.Text = "ResourceGroup"
$LocalGatewayTable.Cell(1,3).Range.Text = "GatewayIPAddress"
$LocalGatewayTable.Cell(1,4).Range.Text = "LocalNetworkAddressSpace"

## Values
$i=2
Write-Host "Creating VPN LocalGateway table"
Foreach ($LocalGateway in $LocalGatewayArray) {


    $LocalGatewayTable.cell(($i),1).range.Bold = 0
    $LocalGatewayTable.cell(($i),1).range.text = $LocalGateway.Name
    $LocalGatewayTable.cell(($i),2).range.Bold = 0
    $LocalGatewayTable.cell(($i),2).range.text = $LocalGateway.ResourceGroupName
    $LocalGatewayTable.cell(($i),3).range.Bold = 0
    $LocalGatewayTable.cell(($i),3).range.text = $LocalGateway.GatewayIpAddress
    $LocalGatewayTable.cell(($i),4).range.Bold = 0
    if ($LocalGateway.LocalNetworkAddressSpace.AddressPrefixes) {
        $LocalAddressSpace = ConvertArrayToLine $LocalGateway.LocalNetworkAddressSpace.AddressPrefixes
    }
    else {
        $LocalAddressSpace = "NOT defined"
    }
    $LocalGatewayTable.cell(($i),4).range.text = $LocalAddressSpace
    $i++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

$Selection.InsertNewPage()
$Selection.Style = $Heading1
$Selection.TypeText("Public IP's")
$Selection.TypeParagraph()
$AllPublicIPs = Get-AzPublicIpAddress | Sort-Object Name
########
######## Create a table for Public IP addresses
########
$AllPublicIPTable = $Selection.Tables.add($Word.Selection.Range, $AllPublicIPs.Count + 1, 4,
[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent)

$AllPublicIPTable.Style = $MediumShading1
$AllPublicIPTable.Cell(1,1).Range.Text = "Name"
$AllPublicIPTable.Cell(1,2).Range.Text = "ResourceGroup"
$AllPublicIPTable.Cell(1,3).Range.Text = "IPAddress"
$AllPublicIPTable.Cell(1,4).Range.Text = "PublicIpAllocationMethod"
$AllPublicIPTable.Cell(1,4).Range.Text = "Usedby"
## Values
$i=2
Write-Host "Creating Public IP table"
Foreach ($PublicIP in $AllPublicIPs) {


    $AllPublicIPTable.cell(($i),1).range.Bold = 0
    $AllPublicIPTable.cell(($i),1).range.text = $PublicIP.Name
    $AllPublicIPTable.cell(($i),2).range.Bold = 0
    $AllPublicIPTable.cell(($i),2).range.text = $PublicIP.ResourceGroupName
    $AllPublicIPTable.cell(($i),3).range.Bold = 0
    $AllPublicIPTable.cell(($i),3).range.text = $PublicIP.IpAddress
    $AllPublicIPTable.cell(($i),4).range.Bold = 0
    if ($PublicIP.IpConfiguration.id) {
        $Parts = $PublicIP.IpConfiguration.id.Split("/")
        $Endpoint = $Parts[8]
    }
    else {
        $Endpoint = "Unused"
    }
    $AllPublicIPTable.cell(($i),4).range.text = $Endpoint
    $i++
}
$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

### Update the TOC now when all data has been written to the document 
$toc.Update()

# Save the document
Write-Host "Creating file $Report."
$Document.SaveAs([ref]$Report,[ref]$SaveFormat::wdFormatDocument)
$word.Quit()

# Free up memory
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word 
Write-Host "Script end"
