﻿param($SCPath, $Sub, $Intag, $Resources, $Task , $File, $SmaResources, $TableStyle) 
If ($Task -eq 'Processing') {

    $ROUTETABLE = $Resources | Where-Object { $_.TYPE -eq 'microsoft.network/routetables' }

    if($ROUTETABLE)
        {
            $tmp = @()

            foreach ($1 in $ROUTETABLE) {
                $ResUCount = 1
                $sub1 = $SUB | Where-Object { $_.Id -eq $1.subscriptionId }
                $data = $1.PROPERTIES
                $Tags = if(![string]::IsNullOrEmpty($1.tags.psobject.properties)){$1.tags.psobject.properties}else{'0'}
                    foreach ($TagKey in $Tags) { 
                        $obj = @{
                            'ID'                            = $1.id;
                            'Subscription'                  = $sub1.Name;
                            'Resource Group'                = $1.RESOURCEGROUP;
                            'Name'                          = $1.NAME;
                            'Location'                      = $1.LOCATION;
                            'Disable BGP Route Propagation' = $data.disableBgpRoutePropagation;
                            'Routes'                        = [string]$data.routes.name;
                            'Routes Prefixes'               = [string]$data.routes.properties.addressPrefix;
                            'Routes BGP Override'           = [string]$data.routes.properties.hasBgpOverride;
                            'Routes Next Hop IP'            = [string]$data.routes.properties.nextHopIpAddress;
                            'Routes Next Hop Type'          = [string]$data.routes.properties.nextHopType;
                            'Resource U'                    = $ResUCount;
                            'Tag Name'                      = [string]$Tag.Name;
                            'Tag Value'                     = [string]$Tag.Value
                        }
                        $tmp += $obj
                        if ($ResUCount -eq 1) { $ResUCount = 0 } 
                    }               
            }
            $tmp
        }
}
Else {
    if ($SmaResources.ROUTETABLE) {

        $TableName = ('RouteTbTable_'+($SmaResources.ROUTETABLE.id | Select-Object -Unique).count)
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 0

        $Exc = New-Object System.Collections.Generic.List[System.Object]
        $Exc.Add('Subscription')
        $Exc.Add('Resource Group')
        $Exc.Add('Name')
        $Exc.Add('Location')
        $Exc.Add('Disable BGP Route Propagation')
        $Exc.Add('Routes')
        $Exc.Add('Routes Prefixes')
        $Exc.Add('Routes BGP Override')
        $Exc.Add('Routes Next Hop IP')
        $Exc.Add('Routes Next Hop Type')
        if($InTag)
            {
                $Exc.Add('Tag Name')
                $Exc.Add('Tag Value') 
            }

        $ExcelVar = $SmaResources.ROUTETABLE 

        $ExcelVar | 
        ForEach-Object { [PSCustomObject]$_ } | Select-Object -Unique $Exc | 
        Export-Excel -Path $File -WorksheetName 'Route Tables' -AutoSize -MaxAutoSizeRows 100 -TableName $TableName -TableStyle $tableStyle -Style $Style
    
    }
}