Import-Module Sharegate

#You will have to modify the file path to match the path on your drive
$csvFile = "C:\temp\CopyContent.csv"

$table = Import-Csv $csvFile -Delimiter ";"

foreach ($row in $table)

{
$srcsite = Connect-Site -Url $row.SourceSite
$dstsite = Connect-Site -Url $row.DestinationSite

$srclist = Get-List -Site $srcsite -Name "Shared Documents" 
$srclist2 = Get-List -Site $srcsite -Name "Private Documents "
$dstlist = Get-List -Site $dstsite -Name "Documents"

Copy-Content -SourceList $srclist -DestinationList $dstlist 
Copy-Content -SourceList $srclist2 -DestinationList $dstlist
}
