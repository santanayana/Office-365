$url = "https://acteurope.sharepoint.com/sites/demo5"
$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$web = $clientContext.Web
$clientContext.Load($web)
$clientContext.ExecuteQuery()
$web.RegionalSettings.LocaleId = 2057
$web.RegionalSettings.Time24 = $true
$web.Update()
$clientContext.ExecuteQuery()