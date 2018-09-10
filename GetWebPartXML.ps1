$url = "https://acteurope.sharepoint.com/sites/CKE-Partners-Management-SFR-CSO-BC-01/HR%20PDS/HRInstructions"
Connect-PnPOnline -Url $url -UseAdfs -Credentials (Get-Credential) 
$wpLocation = "/SitePages/Home.aspx" 
$allWebParts = Get-PnPWebPart -ServerRelativePageUrl $wpLocation
$wp = Get-PnPWebPart -ServerRelativePageUrl $wpLocation -Identity $allWebParts[1].Id 
$wpXml = Get-PnPWebPartXml -ServerRelativePageUrl $wpLocation -Identity $wp.Id > C:\temp\Siteusers.xml 