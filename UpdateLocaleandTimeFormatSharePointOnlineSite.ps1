[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
$url = "https://acteurope.sharepoint.com/sites/demo5"
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext("$url")

$username = 'macsta@statoilfuelretail.com'

$password = Read-Host -Prompt "Password for $username" -AsSecureString

$ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)


$web = $ctx.Web;
$RegionalSettings = $ctx.Web.RegionalSettings
$ctx.Load($web);
$ctx.Load($RegionalSettings);
$ctx.ExecuteQuery();

Write-Host $RegionalSettings.LocaleId "Locale before update"

$web.RegionalSettings.LocaleId = 2057;
Write-Host $RegionalSettings.Time24 "TIme format before update"
$web.RegionalSettings.Time24 =$true
$web.RegionalSettings.FirstDayOfWeek = 1
$web.Update();
$ctx.ExecuteQuery();
Write-Host $web.RegionalSettings.LocaleId 