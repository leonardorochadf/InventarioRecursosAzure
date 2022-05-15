﻿param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle)

If ($Task -eq 'Processing')
{
    <######### Insert the resource extraction here ########>

        $BASTION = $Resources | Where-Object {$_.TYPE -eq 'microsoft.network/bastionhosts'}

    <######### Insert the resource Process here ########>

    if($BASTION)
        {
            $tmp = @()

            foreach ($1 in $BASTION) {
                $ResUCount = 1
                $sub1 = $SUB | Where-Object { $_.Id -eq $1.subscriptionId }
                $data = $1.PROPERTIES
                $BastVNET = $data.ipConfigurations.properties.subnet.id.split("/")[8]
                $BastPIP = $data.ipConfigurations.properties.publicIPAddress.id.split("/")[8]
                $Tags = if(![string]::IsNullOrEmpty($1.tags.psobject.properties)){$1.tags.psobject.properties}else{'0'}
                    foreach ($Tag in $Tags) {
                        $obj = @{
                            'ID'              = $1.id;
                            'Subscription'    = $sub1.Name;
                            'Resource Group'  = $1.RESOURCEGROUP;
                            'Name'            = $1.NAME;
                            'Location'        = $1.LOCATION;
                            'SKU'             = $1.sku.name;
                            'DNS Name'        = $data.dnsName;
                            'Virtual Network' = $BastVNET;
                            'Public IP'       = $BastPIP;
                            'Scale Units'     = $data.scaleUnits;
                            'Tag Name'        = [string]$Tag.Name;
                            'Tag Value'       = [string]$Tag.Value
                        }
                        $tmp += $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }               
            }
            $tmp
        }
}

<######## Resource Excel Reporting Begins Here ########>

Else
{
    <######## $SmaResources.(RESOURCE FILE NAME) ##########>

    if($SmaResources.BASTION)
    {

        $TableName = ('BASTIONTable_'+($SmaResources.BASTION.id | Select-Object -Unique).count)
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'

        $Exc = New-Object System.Collections.Generic.List[System.Object]
        $Exc.Add('Subscription')
        $Exc.Add('Resource Group')
        $Exc.Add('Name')
        $Exc.Add('Location')
        $Exc.Add('SKU')
        $Exc.Add('DNS Name')
        $Exc.Add('Virtual Network')
        $Exc.Add('Public IP')
        $Exc.Add('Scale Units')
        if($InTag)
            {
                $Exc.Add('Tag Name')
                $Exc.Add('Tag Value') 
            }

        $ExcelVar = $SmaResources.BASTION  

        $ExcelVar | 
        ForEach-Object { [PSCustomObject]$_ } | Select-Object -Unique $Exc | 
        Export-Excel -Path $File -WorksheetName 'Bastion Hosts' -AutoSize -MaxAutoSizeRows 100 -TableName $TableName -TableStyle $tableStyle -Style $Style

        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}