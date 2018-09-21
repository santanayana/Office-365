#
# New_ProjectSite.ps1
#
[CmdletBinding()]
Param(

	[Parameter(Mandatory=$True)]
	[string]$Title,
    
    [Parameter(Mandatory=$True)]
    [string]$Url,	

    [Parameter(Mandatory=$True)]
    [string]$Folder,	

    [Parameter(Mandatory=$True)]
    [string]$LogFile,	

	[Parameter(Mandatory=$True)]
	[string]$Description,
    
    [Parameter(Mandatory=$True)]
    [string]$SiteOwner1,

    [Parameter(Mandatory=$False)]
    [string]$SiteOwner2
)

Push-Location $folder
$Credential =  Import-Clixml "C:\Users\macsta\Documents\PowerShell\statoil2.xml"
$settings = .\Get-Settings.ps1
$VerbosePreference = "continue"
$serverRelativeUrl = $url.Substring($url.IndexOf('/',8))    # Skip 8 characters for "https://"
$serverRelativeHomePageUrl = $serverRelativeUrl + '/SitePages/Home.aspx'

############################################################################
# Add-ListViewWebPart
############################################################################
function Add-ListViewWebPart ($serverRelativePageUrl, $webPartName, $listName, $isLibrary, $column, $row) {

    Write-Verbose "Adding web part $webPartName"

    $list = Get-PnPList -Identity $listName
	$listId = $list.Id.ToString().ToUpper()
	if ($isLibrary) {
	    $listUrl = $serverRelativeUrl + "/" + $listName
	} else {
	    $listUrl = $serverRelativeUrl + "/lists/" + $listName
	}
    $views = Get-PnPView -List $list | Where-Object {$_.DefaultView}
    $viewId = ($views[0]).Id.ToString().ToUpper()

	$webPartXml = ((Get-Content "./WebParts/$webPartName") `
                  | Foreach-Object {$_ -replace "%ListId%", $listId} `
                  | Foreach-Object {$_ -replace "%ViewId%", $viewId} `
                  | Foreach-Object {$_ -replace "%ListUrl%", "$listUrl"} `
                  | Foreach-Object {$_ -replace "%PageUrl%", $serverRelativePageUrl} `
                  )

    $webPartXml > temp.xml

	Add-PnPWebPartToWikiPage -Path ./temp.xml -ServerRelativePageUrl $serverRelativePageUrl -Row $row -Column $column
        
    Remove-Item -Path .\temp.xml
}

############################################################################
# Add-ColumnToListView
############################################################################
function Add-ColumnToListView ($listName, $fieldName) {

    $ctx = Get-PnPContext
    $list = Get-PnPList -Identity $listName
    $ctx.Load($list.Views)
    $ctx.ExecuteQuery()
    $view = $list.Views | Where-Object {$_.Title -eq ""}
    $view.ViewFields.Add($fieldName)
    $view.Update()
    $ctx.ExecuteQuery()

}

############################################################################
# Add-ScriptedWebPart
############################################################################
function Add-ScriptedWebPart ($serverRelativePageUrl, $webPartName, $scriptSite, $column, $row) {

    Write-Verbose "Adding web part $webPartName"

	$webPartXml = ((Get-Content "./WebParts/$webPartName") `
                  | Foreach-Object {$_ -replace "%ScriptSite%", $scriptSite} `
                  )

    $webPartXml > temp.xml

	Add-PnPWebPartToWikiPage -Path ./temp.xml -ServerRelativePageUrl $serverRelativePageUrl -Row $row -Column $column
        
    Remove-Item -Path .\temp.xml
}

############################################################################
# New-List
############################################################################
function New-List ($listTemplateName, $listName) {
    $web = Get-PnPWeb
    $ctx = Get-PnPContext
    $listTemplate = $null

    # Work around problem where some list templates are missing on a newly created site collection
    Do {
        Start-Sleep -Seconds 1
        $ctx.Load($web.ListTemplates)
        $ctx.ExecuteQuery()
        $c = $web.ListTemplates.Count
        Write-Verbose "Found $c list templates"
    } Until ($web.ListTemplates.Count -gt 30)

    $listTemplate = $web.ListTemplates | Where-Object {$_.Name -eq $listTemplateName}
    New-PnPList -Title $listName -Template $listTemplate.ListTemplateTypeKind
    Write-Verbose "Created $listName"
}

############################################################################
# Site Provisioning (Inline code)
############################################################################

Write-Verbose "1   - Creating Site Collection"
Write-Verbose "1.1 - Connecting to main site collection"

Connect-PnPOnline $settings.RootSite -Credentials $Credential

Write-Verbose "1.2 - Creating new site collection $Url"

New-PnPTenantSite -Title $title `
				  -Url $url `
				  -Description $Description `
				  -Owner $Credential.UserName `
				  -Lcid $settings.Lcid `
				  -Template $settings.Template `
				  -TimeZone $settings.Timezone `
				  -ResourceQuota $settings.ResourceQuota `
				  -ResourceQuotaWarningLevel $settings.ResourceQuotaWarningLevel `
				  -StorageQuota $settings.StorageQuota `
				  -StorageQuotaWarningLevel $settings.StorageQuotaWarningLevel `
				  -Wait

Write-Verbose "2   - Provisioning Site Contents"
Write-Verbose "2.1 - Connect to new site collection $url"

Connect-PnPOnline $url -Credentials $Credential


if ((Get-PnPWeb).Url.ToLower() -ne $url.ToLower()) {

    Write-Verbose "Unable to connect to new site collection $url"
    Write-Error "Unable to connect to new site collection $url"

} else {

    Write-Verbose "2.2 - Set Security Settings"
    if (($SiteOwner1 -ne $null) -and ($siteOwner1 -ne "")) {
        Write-Verbose "Setting secondary site collection administrator: $siteOwner1"
        $secondaryAdminUser = New-PnPUser -LoginName $siteOwner1
        $secondaryAdminUser.IsSiteAdmin = $true
        $secondaryAdminUser.Update()
    } else {
        Write-Verbose "Site Owner 1 not specified"
		Write-Error "Site Owner 1 not specified"
    }
    if (($SiteOwner2 -ne $null) -and ($siteOwner2 -ne "")) {
        Write-Verbose "Setting secondary site collection administrator: $siteOwner2"
        $secondaryAdminUser = New-PnPUser -LoginName $siteOwner2
        $secondaryAdminUser.IsSiteAdmin = $true
        $secondaryAdminUser.Update()
    } else {
        Write-Verbose "Site Owner 2 not specified"
    }

############################################################################
# Change Regional Settings to UK Locale and 24h time format
############################################################################
	Write-Verbose "2.2.1 - Change RegionalSettings on site:  Get-PnPWeb"
	$ctx = Get-PnPContext
	$web = Get-PnPWeb
	$RegionalSettings=$web.RegionalSettings
	$ctx.Load($Web)
	$ctx.Load($RegionalSettings)
	Invoke-PnPQuery

	Write-Host " RegionalSettings before change: " $web.RegionalSettings.LocaleId 
	$web.RegionalSettings.LocaleId = 2057;
	Write-Host "RegionalSettings after change: " $web.RegionalSettings.LocaleId
	Write-Host "Current time format is 24?" $RegionalSettings.Time24 
	$web.RegionalSettings.Time24 =$true
	Write-Host "New time format is 24?" $RegionalSettings.Time24 
	$web.RegionalSettings.FirstDayOfWeek = 1
	$web.Update();
	Invoke-PnPQuery
	
	Write-Verbose "2.2.2 - Trying to Enable desired features"
	Write-Host "Enabling features on SharePoint site "
    
    
############################################################################
# Enabling all required features on newly created site collection
############################################################################

	$ctx = Get-PnPContext
	Enable-PnPfeature -Identity 063c26fa-3ccc-4180-8a84-b6f98e991df3 -scope site -force
	Enable-PnPfeature -Identity 8581a8a7-cf16-4770-ac54-260265ddb0b2 -scope site -force
	Enable-PnPfeature -Identity b21b090c-c796-4b0f-ac0f-7ef1659c20ae -scope site -force
	Enable-PnPfeature -Identity 2fcd5f8a-26b7-4a6a-9755-918566dba90a -scope site -force
	Enable-PnPfeature -Identity 7c637b23-06c4-472d-9a9a-7c175762c5c4 -scope site -force
	Enable-PnPfeature -Identity 8a4b8de2-6fd8-41e9-923c-c7c3c00f8295 -scope site -force
	Enable-PnPfeature -Identity 7094bd89-2cfe-490a-8c7e-fbace37b4a34 -scope site -force
	Enable-PnPfeature -Identity b435069a-e096-46e0-ae30-899daca4b304 -scope site -force
	Enable-PnPfeature -Identity f6924d36-2fa8-4f0b-b16d-06b7250180fa -scope site -force
	Enable-PnPfeature -Identity 2fcd5f8a-26b7-4a6a-9755-918566dba90a -scope site -force
	Enable-PnPfeature -Identity fde5d850-671e-4143-950a-87b473922dc7 -scope site -force
	Enable-PnPfeature -Identity 6e1e5426-2ebd-4871-8027-c5ca86371ead -scope site -force
	Enable-PnPfeature -Identity 5d0a60c3-1c10-420a-a072-ce0d99913580 -scope site -force
	Enable-PnPfeature -Identity 0af5989a-3aea-4519-8ab0-85d91abe39ff -scope site -force
	Enable-PnPfeature -Identity 365356ee-6c88-4cf1-92b8-fa94a8b8c118 -scope site -force
	Enable-PnPfeature -Identity c6561405-ea03-40a9-a57f-f25472942a22 -scope site -force
	Enable-PnPfeature -Identity c4773de6-ba70-4583-b751-2a7b1dc67e3a -scope site -force
	Enable-PnPfeature -Identity b5934f65-a844-4e67-82e5-92f66aafe912 -scope site -force
	Enable-PnPfeature -Identity e0a9f213-54f5-4a5a-81d5-f5f3dbe48977 -scope site -force
	Enable-PnPfeature -Identity da2e115b-07e4-49d9-bb2c-35e93bb9fca9 -scope site -force
    Enable-PnPfeature -Identity b50e3104-6812-424f-a011-cc90e6327318 -scope site -force
    Enable-PnPfeature -Identity 6c09612b-46af-4b2f-8dfc-59185c962a29 -Scope Site -Force
    Enable-PnPfeature -Identity 02464c6a-9d07-4f30-ba04-e9035cf54392 -Scope Site -Force
    Enable-PnPfeature -Identity c845ed8d-9ce5-448c-bd3e-ea71350ce45b -Scope Site -Force
    Enable-PnPfeature -Identity a44d2aa3-affc-4d58-8db4-f4a3af053188 -Scope Site -Force
																
	Enable-PnPfeature -Identity a7a2793e-67cd-4dc1-9fd0-43f61581207a -scope web
	Enable-PnPfeature -Identity 7201d6a4-a5d3-49a1-8c19-19c4bac6e668 -scope web
	Enable-PnPfeature -Identity 87294c72-f260-42f3-a41b-981a2ffce37a -scope web
	Enable-PnPfeature -Identity d95c97f3-e528-4da2-ae9f-32b3535fbb59 -scope web
	Enable-PnPfeature -Identity d250636f-0a26-4019-8425-a5232d592c01 -scope web
	Enable-PnPfeature -Identity f151bb39-7c3b-414f-bb36-6bf18872052f -scope web
	Enable-PnPfeature -Identity b6917cb1-93a0-4b97-a84d-7cf49975d4ec -scope web
	Enable-PnPfeature -Identity 00bfea71-4ea5-48d4-a4ad-7ea5c011abe5 -scope web
	Enable-PnPfeature -Identity 00bfea71-d8fe-4fec-8dad-01c19a6e4053 -scope web
	Enable-PnPfeature -Identity 57311b7a-9afd-4ff0-866e-9393ad6647b1 -scope web
	
	Write-Host "Features on SharePoint site enabled"
	
	
	Write-Verbose "2.3 - Add template and breadcrumbsolution"
	
	Install-PnPSolution -PackageId 25251922-d03e-4a78-8de4-25beb5d317e9 -SourceFilePath $folder\Migrationtemplate.wsp
    Install-PnPSolution -PackageId 571e2000-6838-4e1b-8eb8-2558256b0f9a -SourceFilePath $folder\SharePointBreadCrumb.wsp
    Install-PnPSolution -PackageId 12100c1c-ab50-488b-8934-f19e776ad888 -SourceFilePath $folder\MeetingWorkTempl.wsp
    Install-PnPSolution  PackageId 5d773141-2e17-47d8-b067-2e05f4405944 -SourceFilePath $folder\Migrationtemplatev3.wsp
	Write-Host "Uploading for template and bredcrumb nawigation solution - done"
	
    Write-Verbose "2.3 - Creating lists"

    # Announcements List
    New-List -ListTemplateName "Announcements" -ListName "Announcements"
    $announcementsList = Get-PnPList -Identity "Announcements"
    Add-PnPField -List $announcementsList `
                 -DisplayName "Urgent" `
                 -InternalName "Urgent" `
                 -Type Boolean `
                 -AddToDefaultView
    
    # Calendar
    New-List -ListTemplateName "Calendar" -ListName "Calendar"
	# Documents Library
	New-List -listTemplateName "Documents" -ListName "Documents Library"
	$documentlibrarylist = Get-PnPList -Identity "Documents Library"
	Add-PnPField -List $documentlibrarylist `
	-DisplayName "" `
	-InternalName "" `
	-Type 
	

    Write-Verbose "2.4 - Removing built-in web parts"

    $webParts = @("Get started with your site", "Site Feed", "Documents")

    foreach ($webPart in $webParts) {
        Write-Verbose "Removing web part $webPart"
        Remove-PnPWebPart -Title $webPart -ServerRelativePageUrl $serverRelativeHomePageUrl
    }

    # To create additional web part templates:
    #
    # 1. Add web part to a test or template master site
    # 2. Extract the XML with Get-PnPWebPartXml
    # 3. Edit the XML and replace the list GUID with %ListId%, view GUID with %ViewId%, and page URL with %PageUrl%

    Write-Verbose "2.5 - Adding new web parts"
    #Add-ListViewWebPart -serverRelativePageUrl $serverRelativeHomePageUrl -webPartName "AnnouncementsWPOrginal.xml" -listName "Announcements" -isLibrary $False -column 1 -row 2
    #Add-ColumnToListView -listName "Announcements" -fieldName "Urgent"
    #Add-ListViewWebPart -serverRelativePageUrl $serverRelativeHomePageUrl -webPartName "CalendarWP.xml" -listName "Calendar" -isLibrary $False -column 1 -row 2
	#Add-ListViewWebPart -serverRelativePageUrl $serverRelativeHomePageUrl -webPartName "LinksWP.xml" -listName "Links" -isLibrary $False -column 1 -row 2
	Add-ScriptedWebPart -serverRelativePageUrl $serverRelativeHomePageUrl -webPartName "MetadataWP.xml" -scriptSite $settings.ScriptSiteUrl -column 2 -row 2
	#Add-ScriptedWebPart -serverRelativePageUrl $serverRelativeHomePageUrl -webPartName "WeatherWP.xml" -scriptSite $settings.ScriptSiteUrl -column 2 -row 2
	Add-ScriptedWebPart -serverRelativePageUrl $serverRelativeHomePageUrl -webPartName "Siteusers.xml" -scriptSite -listName "SiteUsers" -isLibrary $False -column 2 -row 2
    #Add-ListViewWebPart -serverRelativePageUrl $serverRelativeHomePageUrl -webPartName "DocumentsWPOrginal.xml" -listName "Documents Library" -isLibrary $True -column 1 -row 1
	#Add-ColumnToListView -listName "Documents%20Library" -fieldName "Created"
	#Add-ColumnToListView -listName "Documents%20Library" -fieldName "Created By"
	#Add-ColumnToListView -listName "Documents%20Library" -fieldName "File Size"
  
    # Apply site branding - Very light, just a theme and logo

    Write-Verbose "2.6 - Adding branding"
    Set-PnPTheme -ColorPaletteUrl "$serverRelativeUrl/_catalogs/theme/15/palette013.spcolor" -FontSchemeUrl "$serverRelativeUrl/_catalogs/theme/15/fontscheme002.spfont"
    Write-Verbose "2.6.1 - Set theme"
    $ctx = Get-PnPContext
    $web = Get-PnPWeb
    $web.SiteLogoUrl = $settings.ScriptSiteUrl + "/SiteAssets/Circle-K_Europe_logo.gif"
    $web.Update()
    $ctx.ExecuteQuery()
    Write-Verbose "2.6.2 - Set logo"


    # Provision site metadata form

    Write-Verbose "2.7 - Adding site mtadata form"
    .\Add-MetadataForm.ps1 -Url $url -Credentials $Credential

}

Pop-Location