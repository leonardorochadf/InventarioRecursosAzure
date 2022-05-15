﻿param($SCPath, $Sub, $Intag, $Resources, $Task ,$File, $SmaResources, $TableStyle, $Unsupported, $AzCost)
 
If ($Task -eq 'Processing')
{
 
    <######### Insert the resource extraction here ########>

        $wrkspace = $Resources | Where-Object {$_.TYPE -eq 'microsoft.operationalinsights/workspaces'}

    <######### Insert the resource Process here ########>

    if($wrkspace)
        {
            $tmp = @()

            foreach ($1 in $wrkspace) {
                $ResUCount = 1
                $sub1 = $SUB | Where-Object { $_.id -eq $1.subscriptionId }
                $data = $1.PROPERTIES
                $Tags = if(![string]::IsNullOrEmpty($1.tags.psobject.properties)){$1.tags.psobject.properties}else{'0'}
                    foreach ($Tag in $Tags) {
                        $obj = @{
                            'ID'               = $1.id;
                            'Subscription'     = $sub1.Name;
                            'Resource Group'   = $1.RESOURCEGROUP;
                            'Name'             = $1.NAME;
                            'Location'         = $1.LOCATION;
                            'Currency'         = $Cost.Currency;
                            'Daily Cost'       = '{0:C}' -f $Cost.Cost;
                            'SKU'              = $data.sku.name;
                            'Retention Days'   = $data.retentionInDays;
                            'Daily Quota (GB)' = [decimal]$data.workspaceCapping.dailyQuotaGb;
                            'Resource U'       = $ResUCount;
                            'Tag Name'         = [string]$Tag.Name;
                            'Tag Value'        = [string]$Tag.Value
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

    if($SmaResources.WrkSpace)
    {

        $TableName = ('WorkSpaceTable_'+($SmaResources.WrkSpace.id | Select-Object -Unique).count)
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0.0'

        $condtxt = @()

        $Exc = New-Object System.Collections.Generic.List[System.Object]
        $Exc.Add('Subscription')
        $Exc.Add('Resource Group')
        $Exc.Add('Name')
        $Exc.Add('Location')
        $Exc.Add('SKU')
        $Exc.Add('Retention Days')
        $Exc.Add('Daily Quota (GB)')
        if($InTag)
            {
                $Exc.Add('Tag Name')
                $Exc.Add('Tag Value') 
            }

        $ExcelVar = $SmaResources.WrkSpace 

        $ExcelVar | 
        ForEach-Object { [PSCustomObject]$_ } | Select-Object -Unique $Exc | 
        Export-Excel -Path $File -WorksheetName 'Workspaces' -AutoSize -MaxAutoSizeRows 100 -ConditionalText $condtxt -TableName $TableName -TableStyle $tableStyle -Style $Style


        <######## Insert Column comments and documentations here following this model #########>


        #$excel = Open-ExcelPackage -Path $File -KillExcel


        #Close-ExcelPackage $excel 

    }
}