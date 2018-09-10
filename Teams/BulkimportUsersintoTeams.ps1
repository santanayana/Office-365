$csv = Import-CSV C:\temp\linkedinusers.csv 
Connect-MicrosoftTeams 
$Teams = Get-Team
$Team = $Teams | ? {$_.DisplayName -like "*learning*"}
foreach($row in $csv)
{
    Write-Host ($row.UPN)
    Add-TeamUser -GroupId ($Team.GroupId) -Role Member -User ($row.UPN)
}
Disconnect-MicrosoftTeams