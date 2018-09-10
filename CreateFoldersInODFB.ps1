<#
.Synopsis
   This script will create 3 folders into each users' ODFB provided in a CSV file.
.DESCRIPTION
   This script will create 3 folders into each users' ODFB provided in a CSV file and called Private Documents, Shared Documents and Migrated from FDrive, but more can be added

   +++ PRE-REQUISITES +++   
        >> The .CSV file needs to contain a column with header titled "UserPrincipalName" 
        >> SharePoint Online SDK Components and SPO Management Shell installed on the machine running the script
        >> The account provided in the variable $AdminAccount needs to be added as Site Collection Admin to each ODFB PRIOR to running this script (in order to have permission to create the folders)

.EXAMPLE
   .\CreateFoldersInODFB.ps1 -AdminAcct <admin@domain.com> -SPOAcct <SCA_on_ODFB> -TenantName <Name_of_the_O365_Tenant> -CsvFileLocation <CsvFile_Location> 
.EXAMPLE
   .\CreateFoldersInODFB.ps1 (if no parameters are entered, you will be prompted for them)

=================================
Author: Veronique Lengelle 
Date: 03 Jan 2017 
Modified Date 24-04-2018 Maciej Stasiak
Version: 1.1
=================================
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "This is the Admin account neccessary to connect to SPO to run the script", Position = 1)] 
    [string]$AdminAcct,   
    [Parameter(Mandatory = $true, HelpMessage = "This is the Admin account added as a Secondary Site Collection owner on each ODFB", Position = 2)] 
    [string]$SPOAcct,    
    [Parameter(Mandatory = $true, HelpMessage = "Provide O365 tenant name", Position = 3)] 
    [string]$TenantName,
    [Parameter(Mandatory = $true, HelpMessage = "This is the location of the CSV file containing all the users that need new folders", Position = 4)] 
    [string]$CsvFileLocation
)
Start-Transcript
#Script started at
$startTime = "{0:G}" -f (Get-date)
Write-Host "*** Script started on $startTime ***" -f White -b DarkYellow

#Loading assemblies
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles")

#Declare Variables
$Password = Read-Host -Prompt "Please enter your O365 Admin password" -AsSecureString
$Users = Import-Csv -Path $CsvFileLocation


ForEach ($User in $Users) {
    $creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminAcct, $Password)
			
    #Split UPN in 2 parts
    $SplitUPN = ($User.UserPrincipalName).IndexOf("@")
    $Left = ($User.UserPrincipalName).Substring(0, $SplitUPN)
    $Right = ($User.UserPrincipalName).Substring($SplitUPN + 1)

    #Get the username without the @domain.com part
    $shortUserName = ($User.UserPrincipalName) -replace "@" + $Right
			
    #Modify the UPN to replace dot by underscore to match personal URL
    $ShortUPNUnderscore = $shortUserName.Replace(".", "_")

    Write-Host "** Creating folders for"$user.UserPrincipalName -f Yellow

    #Transform domain with underscore to match personal URL
    $DomainUnderscore = $Right.Replace(".", "_")

    #Use the $shortUsername to build the full path
    $spoOD4BUrl = ("https://$TenantName-my.sharepoint.com/personal/" + $ShortUPNUnderscore + "_" + $DomainUnderscore)
    Write-Host ("URL is: " + $spoOD4BUrl) -f Gray

    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($spoOD4BUrl)
    $ctx.RequestTimeout = 16384000
    $ctx.Credentials = $creds
    $ctx.ExecuteQuery()
	
    $web = $ctx.Web
    $ctx.Load($web)
	
    #Target the Document library in user's ODFB
    $spoDocLibName = "Documents"
    $spoList = $web.Lists.GetByTitle($spoDocLibName)
    $ctx.Load($spoList.RootFolder)
	
    #Create Private Documents
    $spoFolder = $spoList.RootFolder
    $Folder1 = "Private Documents"
    $newFolder = $spoFolder.Folders.Add($Folder1)
    $web.Context.Load($newFolder)
    $web.Context.ExecuteQuery()

    #Create Shared Documents 
    $Folder2 = "Shared Documents"
    $newFolder = $spoFolder.Folders.Add($Folder2)
    $web.Context.Load($newFolder)
    $web.Context.ExecuteQuery()

    #Create folder Migrated from FDrive
    $Folder3 = "Migrated from FDrive"
    $newFolder = $spoFolder.Folders.Add($Folder3)
    $web.Context.Load($newFolder)
    $web.Context.ExecuteQuery()
}

Write-Host "'.:All reminded Folders created:.' $Folder1 , $Folder2 , $Folder3 " -f Green

#Script finished at
$endTime = "{0:G}" -f (Get-date)
Write-Host "*** Script finished on $endTime ***" -f White -b DarkYellow
Write-Host "Time elapsed: $(New-Timespan $startTime $endTime)" -f White -b DarkRed

Stop-Transcript