param($Subscriptions,$resources, $Task ,$File, $Sub, $TableStyle)

If ($Task -eq 'Processing')
{
    $ResTable = $resources | Where-Object { $_.type -ne 'microsoft.advisor/recommendations' }
    $resTable2 = $ResTable | Select-Object id, Type, location, resourcegroup, subscriptionid
    $ResTable3 = $ResTable2 | Group-Object -Property type, location, resourcegroup, subscriptionid 

    $tmp = @()

            foreach ($ResourcesSUB in $ResTable3) {
                $ResourceDetails = $ResourcesSUB.name -split ","
                $SubName = $Subscriptions | Where-Object { $_.Subscription.Id -eq ($ResourceDetails[3] -replace (" ", "")) }

                $obj = @{
                    'Subscription'   = $SubName.Subscription.Name;
                    'Resource Group' = $ResourceDetails[2];
                    'Location'       = $ResourceDetails[1];
                    'Resource Type'  = $ResourceDetails[0];
                    'Resources'      = $ResourcesSUB.Count
                }
                $tmp += $obj
            }
    $tmp
}
else 
{
    $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat '0'
    
    $Sub | 
        ForEach-Object { [PSCustomObject]$_ } | 
        Select-Object 'Subscription',
        'Resource Group',
        'Location',
        'Resource Type',
        'Resources' | Export-Excel -Path $File -WorksheetName 'Subscriptions' -AutoSize -MaxAutoSizeRows 100 -TableName 'Subscriptions' -TableStyle $tableStyle -Style $Style -Numberformat '0' -MoveToEnd 


}