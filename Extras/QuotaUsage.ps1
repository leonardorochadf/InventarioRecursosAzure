param($File, $AzQuota, $TableStyle)

$tmp = @()
$Total = ($AzQuota.Data).count
foreach($Quota in $AzQuota)
{
    foreach($Data in $Quota.Data)
        {            
            $FreevCPU = ''
            if($Data.Name.LocalizedValue -like '*vCPUs'){$FreevCPU = $Data.limit - $Data.CurrentValue}
            $obj = @{
                'Subscription' = $Quota.Subscription;
                'Region' = $Quota.Location;
                'Current Usage' = $Data.currentValue;
                'Limit' = $Data.limit;
                'Quota' = $Data.Name.LocalizedValue;
                'vCPUs Available' = $FreevCPU;
                'Total' = $Total
            }
            $tmp += $obj
        }
}

        $ExcelVar = $tmp

            $TableName = ('QuotaTable_'+$ExcelVar[0].Total)
            $ExcelVar | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object -Unique 'Subscription',         
            'Region',
            'Current Usage',
            'Limit',
            'Quota',
            'vCPUs Available' | 
            Export-Excel -Path $File -WorksheetName 'Quota Usage' -AutoSize -MaxAutoSizeRows 100 -TableName $TableName -TableStyle $tableStyle -Numberformat '0' -MoveToEnd



                