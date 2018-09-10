[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
#Configure the location for the output file
$Output="C:\Output.csv";
"Site URL"+","+"Owner Login"+","+"Owner Email"+","+"Root Site Last Modified"+","+"Quota Limit (MB)"+","+"Total Storage Used (MB)"+","+"Site Quota Percentage Used" | Out-File -Encoding Default -FilePath $Output;
#Specify the root site collection within the Web app
$Siteurl="Your site URL";
$Rootweb=New-Object Microsoft.Sharepoint.Spsite($Siteurl);
$Webapp=$Rootweb.Webapplication;
#Loops through each site collection within the Web app, if the owner has an e-mail address this is written to the output file
Foreach ($Site in $Webapp.Sites)
{if ($Site.Quota.Storagemaximumlevel -gt 0) {[int]$MaxStorage=$Site.Quota.StorageMaximumLevel /1MB} else {$MaxStorage="0"}; 
if ($Site.Usage.Storage -gt 0) {[int]$StorageUsed=$Site.Usage.Storage /1MB};
if ($Storageused-gt 0 -and $Maxstorage-gt 0){[int]$SiteQuotaUsed=$Storageused/$Maxstorage* 100} else {$SiteQuotaUsed="0"}; 
$Web=$Site.Rootweb; $Site.Url + "," + $Site.Owner.Name + "," + $Site.Owner.Email + "," +$Web.LastItemModifiedDate.ToShortDateString() + "," +$MaxStorage+","+$StorageUsed + "," + $SiteQuotaUsed | Out-File -Encoding Default -Append -FilePath $Output;$Site.Dispose()};