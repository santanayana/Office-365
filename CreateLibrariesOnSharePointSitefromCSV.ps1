#Create libraries from Csv file and break inheritance
$openCsv = Import-Csv -path 'C:\temp\Customer list.csv' -Delimiter ";"
$url = "https://acteurope.sharepoint.com/teams/CKE-EU-FC_AND_AC_IRL_communication/CircleK_Ireland"
$web = Get-PnPWeb $url

if ($cred -eq $null)
{
    $cred = Get-Credential
    Connect-PnPOnline $url -Credentials $cred
}

$CSOM_context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
$CSOM_credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.UserName, $cred.Password)
$CSOM_context.Credentials = $CSOM_credentials

$lists = $CSOM_context.Web.Lists
$CSOM_context.Load($lists)
$CSOM_context.ExecuteQuery()

$openCsv | ForEach-Object { 
    Write-Host  $_.JCUSTOMER -NoNewline
    
    try {
        
        New-PnPList -Title $_.JCUSTOMER -Template DocumentLibrary -OnQuickLaunch
        #$spolist = Get-PnPList "$_.JCUSTOMER"
        Write-Host -ForegroundColor Green "  $_.JCUSTOMER  Created "  `n  
        #$spolist.BreakRoleInheritance($true, $true)
        #$spolist.Update()
        #$spolist.Context.Load($spolist)
        #$spolist.Context.ExecuteQuery()
        
        
    }
    catch  {
        Write-Host -ForegroundColor Red   "   [FAILURE] - " $_.Exception.Message `n 
    }   
}
$ignoreList = "Composed Looks", "Access Requests", "Site Pages", "Site Assets", "Style Library", "_catalogs/hubsite", "Translation Packages" , "Content Organizer Rules", "Drop Off Library", "Master Page Gallery" ,"MicroFeed"

$listsnew  = $CSOM_context.Web.Lists
$CSOM_context.Load($listsnew)
$CSOM_context.ExecuteQuery()
$listsnew  | Where-Object {  $_.BaseTemplate -eq 101 -and $_.Title -inotin $ignoreList}  | ForEach-Object {

    Write-Host $_.Title -NoNewline

    try {
        $spoList = Get-PnPList $_.Title
        $spoList.BreakRoleInheritance($true, $true)
        $spoList.Update()
        $spoList.Context.Load($spoList)
        $spoList.Context.ExecuteQuery()

        Write-Host -ForegroundColor Green "   [permission inheritance broken] "  `n  
    }
    catch  {
        Write-Host -ForegroundColor Red   "   [FAILURE] - " $_.Exception.Message `n  
    }
}