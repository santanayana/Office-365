#Create the Office 365 Group
New-UnifiedGroup –DisplayName NICDemo96 –Alias NICDemo96 –EmailAddresses "NICDemo96@M365x963508.onmicrosoft.com" -owner GA-sha256@M365x963508.onmicrosoft.com -RequireSenderAuthenticationEnabled $False -Verbose
#This is optional, but may be a good practice initially since Office 365 Groups may clutter your Global Addressbook
Set-UnifiedGroup –Identity NICDemo96 –HiddenFromAddressListsEnabled $true
#Create the Team, provide the GUID object ID to specify the Group
$group = New-Team -Group (Get-UnifiedGroup NICDemo96).ExternalDirectoryObjectId -Verbose
 
#Check your Teams, will only list teams you are a member of
Get-Team
 
#Add Channels to the Team
New-TeamChannel -GroupId $group.GroupId -DisplayName "1 Adoption" -Verbose
New-TeamChannel -GroupId $group.GroupId -DisplayName "2 Deployment" -Verbose
New-TeamChannel -GroupId $group.GroupId -DisplayName "3 Operations" -Verbose
New-TeamChannel -GroupId $group.GroupId -DisplayName "4 Change Management" -Verbose
Set-TeamFunSettings -GroupId $group.GroupId -AllowCustomMemes true -Verbose
 
#add owners and members, easier to do with Teams cmdlet
$Owners = "PradeepG@M365x963508.onmicrosoft.com","PattiF@M365x963508.onmicrosoft.com","LidiaH@M365x963508.onmicrosoft.com","MiriamG@M365x963508.onmicrosoft.com"
$Users = "IrvinS@M365x963508.onmicrosoft.com","JohannaL@M365x963508.onmicrosoft.com","DebraB@M365x963508.onmicrosoft.com"
ForEach ($Owner in $Owners){Add-TeamUser -GroupId $group.GroupId -User $Owner -Role Owner}
ForEach ($User in $Users){Add-TeamUser -GroupId $group.GroupId -User $User -Role Member -Verbose}
 
#Check that members are added, know that it could take up to 24 hours until they are actually added to Microsoft Teams
Get-TeamUser -GroupId $group.GroupId
Get-UnifiedGroupLinks NICDemo96 -LinkType owner
Get-UnifiedGroupLinks NICDemo96 -LinkType member