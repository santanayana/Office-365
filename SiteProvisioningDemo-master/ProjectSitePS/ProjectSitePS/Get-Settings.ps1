#
# Settings for Project Site Provisioning scripts
#
[PSCustomObject]@{

	ScriptSiteUrl = "https://acteurope.sharepoint.com";
    RootSite = "https://acteurope.sharepoint.com";
    Lcid = 1033;
	ResourceQuotaWarningLevel = 0;
	ResourceQuota = 0;
	StorageQuotaWarningLevel = 300;
	StorageQuota = 500;
	Template = "STS#0";
	# Use Get-SPOTimeZoneId to get a list of time zones
	Timezone = 4;   # Central Europe Standard Time
}