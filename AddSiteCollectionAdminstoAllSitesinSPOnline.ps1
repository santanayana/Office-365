$sites = get-sposite -limit All
foreach($site in $sites){
    Write-Host "$site.Title" + " " + "$site.Url"  -ForegroundColor Yellow
    Set-SPOUser -LoginName "sa-sfr-spadmin3@statoilfuelretail.com" -IsSiteCollectionAdmin $true -Site $site.Url
    Set-SPOUser -LoginName "sa-sfr-spadmin4@statoilfuelretail.com" -IsSiteCollectionAdmin $true -Site $site.Url 
    Set-SPOUser -LoginName "sa-sfr-spadmin5@statoilfuelretail.com" -IsSiteCollectionAdmin $true -Site $site.Url
}