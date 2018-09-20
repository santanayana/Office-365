#Create the Office 365 Group
New-UnifiedGroup –DisplayName "miles new value proposition" –Alias miles_new_value_proposition –EmailAddresses "miles_new_value_proposition@statoilfuelretail.com" -owner MDI@statoilfuelretail.com -RequireSenderAuthenticationEnabled $False -Verbose
#This is optional, but may be a good practice initially since Office 365 Groups may clutter your Global Addressbook
Set-UnifiedGroup –Identity "miles new value proposition" –HiddenFromAddressListsEnabled $true
#Create the Team, provide the GUID object ID to specify the Group
$group = New-Team -Group (Get-UnifiedGroup NICDemo96).ExternalDirectoryObjectId -Verbose
 
#Check your Teams, will only list teams you are a member of
Get-Team
 
#Add Channels to the Team
New-TeamChannel -GroupId $group.GroupId -DisplayName "1 Documentation" -Verbose
New-TeamChannel -GroupId $group.GroupId -DisplayName "2 Development" -Verbose
New-TeamChannel -GroupId $group.GroupId -DisplayName "3 Communication" -Verbose
New-TeamChannel -GroupId $group.GroupId -DisplayName "4 Change Management" -Verbose
Set-TeamFunSettings -GroupId $group.GroupId -AllowCustomMemes true -Verbose
 
#add owners and members, easier to do with Teams cmdlet
$Owners = "MDI@statoilfuelretail.com","ML169@statoilfuelretail.com"
$Users = "MDI@statoilfuelretail.com" ,"ML169@statoilfuelretail.com"
ForEach ($Owner in $Owners){Add-TeamUser -GroupId $group.GroupId -User $Owner -Role Owner}
ForEach ($User in $Users){Add-TeamUser -GroupId $group.GroupId -User $User -Role Member -Verbose}
 
#Check that members are added, know that it could take up to 24 hours until they are actually added to Microsoft Teams
Get-TeamUser -GroupId $group.GroupId
Get-UnifiedGroupLinks "miles new value proposition" -LinkType owner
Get-UnifiedGroupLinks "miles new value proposition" -LinkType member