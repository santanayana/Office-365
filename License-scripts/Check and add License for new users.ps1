#Connecto to 0365 - required modules must be installed AZURE AD and Office Online POwershell Module
$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection


Import-PSSession $Session -AllowClobber
Connect-MsolService -Credential $UserCredential

#Import existing users

$Users= Import-Csv "C:\Script\O365\New_Employees.csv" -Delimiter ";"

#Add licences        
foreach ($MSOLuser in $users) {

$user =  Get-MsolUser -UserPrincipalName $MSOLuser.UserPrincipalName 
    
    if (!$user.IsLicensed)  { 

        $LicenseOption = New-MsolLicenseOptions -AccountSkuId "mynetcompany:DEVELOPERPACK" -DisabledPlans MCOSTANDARD, RMS_S_ENTERPRISE, EXCHANGE_S_ENTERPRISE  
        $location = Read-Host -Prompt 'Input user location PL,DE,NO,DK,GE'
                 
                 foreach ($MSOLuser in $users) {
                    
                    
                    Set-MsolUser -UserPrincipalName $user.UserPrincipalName -UsageLocation $location
                    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses mynetcompany:DEVELOPERPACK -LicenseOptions $LicenseOption
                        
                        }
                    }

    else { 
    write-host $MSOLuser.UserPrincipalName "has license or not exisit"
    }
    }