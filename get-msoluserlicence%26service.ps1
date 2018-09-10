<#
.DESCRITION
    This script will create a comma separated file with a line per user and the following columns: 
    Display Name, Domain, UPN, Is Licensed?, all the SKUs in tenant, all the services, 
    Errors, ImmutableId and BlockCredential.
    
    It is based on previoius work from Alan Byrne and a techet article:
    - Export a Licence reconciliation report from Office 365 using Powershell https://gallery.technet.microsoft.com/scriptcenter/Export-a-Licence-b200ca2a
    - View account license and service details with Office 365 PowerShell https://technet.microsoft.com/en-us/library/dn771771.aspx

    Since Microsoft is using other names for licenses then the names which are commonly used (i.e. EnterPrisePack instead of E3 licenses) a translation table is used
    However it only containce the licenses could find. If the script errors, please complete the table at line 32.

.PARAMETER
    After starting the script it will ask for the credentials to connect to Office 365
    
.OUTPUT
    Office_365_License_service_report.csv will be writen in the current directory

.NOTES
    Version:          1.0
    Author:           Marcus Tarquinio
    Creation Date:    31 January 2016
    Purpose/Change:   Initial script development
#>

# Create a look up table for Human friendly license names

$Sku = @{
	"DESKLESSPACK" = "Office 365 (Plan K1)"
	"DESKLESSWOFFPACK" = "Office 365 (Plan K2)"
	"LITEPACK" = "Office 365 (Plan P1)"
	"EXCHANGESTANDARD" = "Office 365 Exchange Online Only"
	"STANDARDPACK" = "Enterprise Plan E1"
	"STANDARDWOFFPACK" = "Office 365 (Plan E2)"
	"ENTERPRISEPACK" = "Enterprise Plan E3"
	"ENTERPRISEPACKLRG" = "Enterprise Plan E3"
	"ENTERPRISEWITHSCAL" = "Enterprise Plan E4"
	"STANDARDPACK_STUDENT" = "Office 365 (Plan A1) for Students"
	"STANDARDWOFFPACKPACK_STUDENT" = "Office 365 (Plan A2) for Students"
	"ENTERPRISEPACK_STUDENT" = "Office 365 (Plan A3) for Students"
	"ENTERPRISEWITHSCAL_STUDENT" = "Office 365 (Plan A4) for Students"
	"STANDARDPACK_FACULTY" = "Office 365 (Plan A1) for Faculty"
	"STANDARDWOFFPACKPACK_FACULTY" = "Office 365 (Plan A2) for Faculty"
	"ENTERPRISEPACK_FACULTY" = "Office 365 (Plan A3) for Faculty"
	"ENTERPRISEWITHSCAL_FACULTY" = "Office 365 (Plan A4) for Faculty"
	"ENTERPRISEPACK_B_PILOT" = "Office 365 (Enterprise Preview)"
	"STANDARD_B_PILOT" = "Office 365 (Small Business Preview)"
	"VISIOCLIENT" = "Visio Pro Online"
	"POWER_BI_ADDON" = "Office 365 Power BI Addon"
    "POWER_BI_INDIVIDUAL_USE" = "Power BI Individual User"
    "POWER_BI_STANDALONE" = "Power BI Stand Alone"
    "POWER_BI_STANDARD" = "Power-BI standard"
	"PROJECTESSENTIALS" = "Project Lite"
    "PROJECTCLIENT" = "Project Professional"
	"PROJECTONLINE_PLAN_1" = "Project Online"
	"PROJECTONLINE_PLAN_2" = "Project Online and PRO"
	"ECAL_SERVICES" = "ECAL"
    "EMS" = "Enterprise Mobility Suite"
    "RIGHTSMANAGEMENT_ADHOC" = "Windows Azure Rights Management"
    "MCOMEETADV" = "PSTN conferencing"
    "SHAREPOINTSTORAGE" = "SharePoint storage"
    "PLANNERSTANDALONE" = "Planner Standalone"
    "CRMIUR" = "CMRIUR"
    "BI_AZURE_P1" = "Power BI Reporting and Analytics"
    "INTUNE_A" = "Windows Intune Plan A"
	}
		
$credential = Get-Credential

Connect-MsolService -Credential $credential 

write-host "Connecting to Office 365..."

# The Output will be written to this file in the current working directory
$LogFile = "Office_365_License_service_report.csv"

# Get a list of all licences that exist within the tenant
$licensetype = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 1}

# Build the Header for the CSV file
$headerstring = "Display Name, Domain, UPN, Is Licensed?"
$allservices = ""
$numlicenses = 0
# i am sure there is a more elegant way to declare an array :-(
$numservices = 0,0,0,0,0,0,0,0,0,0,0,0,0
$nameservices = " "," "," "," "," "," "," "," "," "," "," "," "," "

# Loop through all licence types found in the tenant to add licensetypes
write-host "Geting the licenses and writing the header..."
foreach ($license in $licensetype) 
{	
    $headerstring = $headerstring + "," + $Sku.Item($license.SkuPartNumber)
        
    # Get a list of all the services in the tenant and add them to the hearder
    $numservices[$numlicenses] = (Get-MsolAccountSku | where {$_.AccountSkuId -eq $license.AccountSkuId}).ServiceStatus.serviceplan.servicename.count
    for($i=0;$i -lt $numservices[$numlicenses]; ++$i)
    {$headerstring = $headerstring + "," + (Get-MsolAccountSku | where {$_.AccountSkuId -eq $license.AccountSkuId}).ServiceStatus[$i].serviceplan.servicename}
    $numlicenses = $numlicenses + 1
}

# Add other attributres
$headerstring = $headerstring + ",Errors, ImmutableId, BlockCredential"

Out-File -FilePath $LogFile -InputObject $headerstring -Encoding UTF8 -append

# Get a list of all the users in the tenant
write-host "Getting all users in the Office 365 tenant..."
$users = Get-MsolUser -all

# Loop through all users found in the tenant
foreach ($user in $users) 
{
# We use last name, comma first name as display name so I removed the comma (if you dont use the next line instead)
    $linestring = $user.displayname -Replace ",",""
#   $linestring = $user.displayname

	write-host ("Processing " + $linestring)
    $linestring = $linestring + "," + $user.UserPrincipalName.Split("@")[1]  + "," + $user.userprincipalname  + "," + $user.isLicensed

    if ($user.isLicensed) 
    {
        # Loop through all licence types found in the tenant
        for($j=0;$j -lt $numlicenses; ++$j)
        {	
            $userhaslicense = "No"
            $aux = 0
            # Loop through all licences assigned to this user
	        foreach ($row in $user.licenses)
	        {
                if ($row.AccountSkuId.ToString() -eq $licensetype.AccountSkuId[$j])
                {
                    $userhaslicense = "Yes"
                    $index = $aux
                }
                $aux = $aux + 1
            }

            $linestring = $linestring + "," + $userhaslicense
            
            if ($userhaslicense -eq "Yes")
            {
            # Now let's get the services enabled for each user
                for($i=0;$i -lt $numservices[$j]; ++$i)
                {
                    $linestring = $linestring + "," 
                    $linestring = $linestring + $user.licenses[$index].ServiceStatus[$i].ProvisioningStatus
                }
            }
            Else
            {
            # and spaces when the user dont have that SkuId
                for($i=0;$i -lt $numservices[$j]; ++$i)
                {$linestring = $linestring + ","}
            }
        }
    }
    else
    {
        for($j=0;$j -lt $numlicenses; ++$j)
        {
            $linestring = $linestring + ",No"
            for($i=0;$i -lt $numservices[$j]; ++$i)
            {$linestring = $linestring + "," }
        }
    }

    $linestring = $linestring + "," + $user.Errors + "," + $user.ImmutableId + "," + $user.BlockCredential
    Out-File -FilePath $LogFile -InputObject $linestring -Encoding UTF8 -append
}
write-host ("Script Completed. Results available in " + $LogFile)
