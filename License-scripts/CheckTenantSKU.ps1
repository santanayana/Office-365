# Check tenant plan 
$tenantSKU = Get-MsolAccountSku
# Check all avaliable SKU's in current tenant service plan
Get-MsolAccountSku | Select -ExpandProperty ServiceStatus 
#or
#Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPACK"} | ForEach-Object {$_.ServiceStatus}
