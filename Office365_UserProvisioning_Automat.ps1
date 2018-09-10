#Required -Modules ActiveDirectory,MSOnline,MSOnlineExtended
#Required -Version 3.0



#region CONFIGURATION

   ###############################################################################
 ###################################################################################
###                                                                               ###
##   General Variables                                                             ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$DEBUG = $false
$PILOT = $true
$NO_MAILBOX_CSV = "$PSScriptRoot\NO_MAILBOX.csv"
$EnableMFA = $false
$GeneralError = 0 -as [System.UInt16]
$Now = Get-Date
$MaxLogSize = 5MB
$MaxLogArchived = 30
$Error.Clear()
$Users = @()
$UsageLocation = "NO"
$RemoteDomain = "acteurope.mail.onmicrosoft.com"
#$ApiUrl = "http://idmohs-int-pub.statoilfuelretail.com:10229/CircleKApi/OIDMOffice365Service/"
$ApiUrl = "https://sso.statoilfuelretail.com/CircleKApi/OIDMOffice365Service/"
$AlertTimeSpanHrs = 24
$AlertAttempts = 3
$GiveUpTimeSpanHrs = 72
$GiveUpAttempts = 6
$IdamErrMsg = ""

## Mail Config
$SmtpServer = "relay.statoilfuelretail.com"
$From = "User Provisioning <sa-sfr-o365-01@circlekeurope.com>"
$ToAdmin = @("maciej.stasiak@circlekeurope.com")
#$ToIdam = @("idmsy@circlekeurope.com","maciej.stasiak@circlekeurope.com")               #"idmsy@circlekeurope.com"
$ToIdam = @("idmsy@circlekeurope.com")
$ToLicence = @("maciej.stasiak@circlekeurope.com")   #"Gita.Vimba@circlekeurope.com"




   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Log File                                                                      ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$LogEntry = "".PadRight(80,"=") + "`r`n[$($Now.ToString('yyyy-MM-dd HH:mm:ss'))]                                           START PROCESSING`r`n"
$LogFile = "$PSScriptRoot\$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)).LOG"
#"`$LogFile       : $LogFile"

if(Test-Path $LogFile)
{

    if((Get-ChildItem $LogFile).Length -ge $MaxLogSize)
    {
            
        $CurFolder = [System.IO.Path]::GetDirectoryName($LogFile)
        $ArchFolder = "$CurFolder\LogArch"
        $NewName = "$([System.IO.Path]::GetFileNameWithoutExtension($LogFile)).$($Now.ToString("yyyyMMddHHmmss")).LOG"

        #"`$CurFolder     : $CurFolder"
        #"`$ArchFolder    : $ArchFolder"
        #"`$NewName       : $NewName"

        # Rename the oversized log file
        try
        {
            Rename-Item -LiteralPath $LogFile -NewName "$NewName"
            $LogEntry += "`r`n`r`nACTION:`r`nOversized log file renamed to '$NewName'"
        }
        catch
        {
            $LogEntry += "`r`n`r`nERROR:`r`nCannot rename the oversized LOG file.`r`n$($Error[0].Exception.Message)"
            $GeneralError = $GeneralError -bor 1
        }


        # Move the renamed oversized log file to archive folder
        if (($GeneralError -band 1) -eq 0)
        {

            # Check if the archive folder exists
            try
            {
                if(!(Test-Path $ArchFolder))
                {
                    New-Item -Path $ArchFolder -ItemType Directory -ErrorAction Stop | Out-Null
                    $LogEntry += "`r`n`r`nACTION:`r`nArchive folder created as '$ArchFolder'"
                }
            }
            catch
            {
                $LogEntry += "`r`n`r`nERROR:`r`nCannot create a folder for log archive '$ArchFolder'.`r`n$($Error[0].Exception.Message)"
                $GeneralError = $GeneralError -bor 1
            }


            # Move the renamed, oversized log file to the archive folder and create new log file
            if (($GeneralError -band 1) -eq 0)
            {
                try
                {
                    Move-Item -Path "$CurFolder\$NewName" -Destination $ArchFolder
                    $LogEntry += "`r`n`r`nACTION:`r`nLog file '$NewName' moved to archive '$ArchFolder'."
                }
                catch
                {
                    $LogEntry += "`r`n`r`nERROR:`r`nCannot move the oversized log to archive folder '$ArchFolder'.`r`n$($Error[0].Exception.Message)"
                    $GeneralError = $GeneralError -bor 1
                }
            }
        }


        #
        # Remove the old archive files, if their amount exceeds
        # the number configured in variable $MaxLogArchived
        #

        $FileMask = "$([System.IO.Path]::GetFileNameWithoutExtension($LogFile)).??????????????.LOG"
        $Logs = Get-ChildItem "$ArchFolder\$FileMask" | sort Name

        
        ##DEBUG
        #0..15 | % { Copy-Item "D:\o365\IDAM\_fake_API.txt" "D:\o365\IDAM\LogArch\_trash_document.$($_.ToString("000")).txt" }
        #$MaxLogArchived = 52
        #$MaxLogArchived--
        #$Logs = Get-ChildItem "$ArchFolder\*" | sort Name

        try
        {
            if ($Logs.Count -gt $MaxLogArchived)
            {
                $Logs = $Logs | select -First ($Logs.Count - $MaxLogArchived)
                $Logs | Remove-Item
                $LogEntry += "`r`n`r`nACTION:`r`nRemoved old archived log files:`r`n"
                $Logs | % { $LogEntry += "  • $($_.Name)`r`n" }
            }
            
        }
        catch
        {
            $LogEntry += "`r`n`r`nERROR:`r`nCannot cleanup log archive folder '$ArchFolder'.`r`n$($Error[0].Exception.Message)"
            $GeneralError = $GeneralError -bor 512
        }
        finally
        {
            Remove-Variable FileMask,Logs
        }

    }
}



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Commands file for debugging mode                                              ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

if ($DEBUG) { "$("`r`n`r`n".PadRight(60,'='))  $($Now.ToString("yyyy-MM-dd HH:mm:ss"))  $("`r`n`r`n".PadLeft(60,'='))" | Out-File "$PSScriptRoot\DEBUG.Commands.txt" -Append }



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Credentials and security context                                              ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

if ($env:COMPUTERNAME -ne "CRKEXCRTV001P" -or $env:USERNAME -ne "sa-sfr-o365-01")
{
    $LogEntry += "`r`n`r`nATTENTION!!!`r`nUnapproved security principals:`r`n`r`n"
    $LogEntry += ""
    $LogEntry += "               Required             Actual`r`n"
    $LogEntry += "               ------------------   ------------------`r`n"
    $LogEntry += "    Computer   CRKEXCRTV001P        $env:COMPUTERNAME`r`n"
    $LogEntry += "    User       sa-sfr-o365-01       $env:USERNAME`r`n"

    Send-MailMessage -SmtpServer $SmtpServer -From $From -To $ToAdmin -Subject "Unauthorized process launch!!!" -Body $LogEntry
    $LogEntry | Out-File $LogFile -Append
    Exit
}
<#
## Code to run to get the credentials
## NOTE: This encrypted password only works when run from the same computer and by the same user!
(Get-Credential -UserName "user.prov@acteurope.onmicrosoft.com" -Message "Password to access on-prem endpoint:").Password | ConvertFrom-SecureString
#>
$OnPremCred = New-Object -TypeName PSCredential -ArgumentList "SFR\sa-sfr-o365-01",( '01000000d08c9ddf0115d1118c7a00c04fc297eb010000002732b457a126a745add3f42e775069eb0000000002000000000003660000c0000000100000001ffd624d7b5f522e0e076a84b86d59260000000004800000a0000000100000002f943dc01465c8cdb0dfffc080491a6018000000be90637b275a4759ebd71349b0b92eb975e06d668005455c14000000dba15bf14d98046ff0883b6b264f58014dfc0d76' | ConvertTo-SecureString )
$OPCmd = "`$OPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://CRKEXCRTV001P.statoilfuelretail.com/PowerShell/ -Authentication Default -Credential `$OnPremCred"

$OnlineCred = New-Object -TypeName PSCredential -ArgumentList "user.prov@acteurope.onmicrosoft.com",( '01000000d08c9ddf0115d1118c7a00c04fc297eb010000002732b457a126a745add3f42e775069eb0000000002000000000003660000c0000000100000006fdc90e5d9fa6984288f707e989a742e0000000004800000a00000001000000059756d6c95a3d22e4a552f189b99407b18000000f8a6b186c612a1b767fb8039381ecaff5a53ca5a24d7dd31140000008e99e803c56736236ac803f8a6e3c4f8d0b6cdb2' | ConvertTo-SecureString )
#$OLCmd = "`$OLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential `$OnlineCred -Authentication Basic -AllowRedirection -SessionOption (New-PSSessionOption -ProxyAccessType NoProxyServer)"



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Config File                                                                   ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$ConfigFile = "$PSScriptRoot\$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)).Config.JSON"
#"`$ConfigFile    : $ConfigFile"
$Config = $null

if (Test-Path $ConfigFile)
{
    try
    {
        # Import config file
        $Config = Get-Content $ConfigFile -Raw -ErrorAction Stop | ConvertFrom-Json
        $LogEntry += "`r`n`r`nACTION:`r`nImported configuration file '$ConfigFile'."


        # Check 'LastDateTo'. 'LastDateFrom' is not really relevant
        if ("LastDateTo" -in ($Config | Get-Member -MemberType NoteProperty).Name)
        {
            $Config.LastDateTo = $Config.LastDateTo -as [System.DateTime]
            if ($Config.LastDateTo -eq $null -or $Config.LastDateTo.GetType().FullName -ne "System.DateTime")
            {
                $Config.LastDateTo = $Now.AddDays(-30)
                $LogEntry += "`r`n`r`nWARNING:`r`nDue to wrong format of 'LastDateTo' property in config file, default value was assumed."
            }
        }
        else
        {
            $Config | Add-Member -NotePropertyName LastDateTo -NotePropertyValue $Now.AddDays(-30)
            $LogEntry += "`r`n`r`nWARNING:`r`nDue to missing 'LastDateTo' property in config file, the property was added with default value."
        }

        # Slide the time window
        $Config.LastDateFrom = $Config.LastDateTo
        $Config.LastDateTo = $Now


        # Save to file
        #$Config | select LastDateFrom,LastDateTo | ConvertTo-Json | Out-File $ConfigFile
        $Config | select @{N="LastDateFrom";E={if($_.LastDateFrom -eq $null){$null}else{$_.LastDateFrom.ToString("yyyy-MM-dd HH:mm:ss")}}},@{N="LastDateTo";E={if($_.LastDateTo -eq $null){$null}else{$_.LastDateTo.ToString("yyyy-MM-dd HH:mm:ss")}}} | ConvertTo-Json | Out-File $ConfigFile

    }
    catch [System.ArgumentException]
    {
        $LogEntry += "`r`n`r`nERROR:`r`nWrong JSON format exception when importing file '$ConfigFile'"
        $Config = $null
        $GeneralError = $GeneralError -bor 2
    }
    catch [System.Exception]
    {
        $LogEntry += "`r`n`r`nERROR:`r`n$($Error[0].Exception.Message)"
        $Config = $null
        $GeneralError = $GeneralError -bor 2
    }
}


# If config file missing or contains wrong data, create one with default values
if ($Config -eq $null)
{
    $Config = New-Object -TypeName PSObject
    $Config | Add-Member -NotePropertyName LastDateFrom -NotePropertyValue $Now.AddDays(-30).ToString("yyyy-MM-dd HH:mm:ss")
    $Config | Add-Member -NotePropertyName LastDateTo -NotePropertyValue $Now.ToString("yyyy-MM-dd HH:mm:ss")
    $Config | ConvertTo-Json | Out-File $ConfigFile
}



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Office 365 Licenses Definition                                                ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$Lic = '[
  {
    "Name" : "E3",
    "AccountSkuId" : "acteurope:ENTERPRISEPACK",
    "IdamRole" : "Mailbox_FullAccess",
    "Available" : null,
    "Needed" : null,
    "AlertLevel" : 50
  },
  {
    "Name" : "E1",
    "AccountSkuId" : "acteurope:STANDARDPACK",
    "IdamRole" : "Mailbox_Limited",
    "Available" : null,
    "Needed" : null,
    "AlertLevel" : 50
  },
  {
    "Name" : "F1",
    "AccountSkuId" : "acteurope:DESKLESSPACK",
    "IdamRole" : "Mailbox_Webmail",
    "Available" : null,
    "Needed" : null,
    "AlertLevel" : 50
  },
  {
    "Name" : null,
    "AccountSkuId" : "#NONE#",
    "IdamRole" : "NO_MAILBOX",
    "Available" : null,
    "Needed" : null,
    "AlertLevel" : null
  }
]' | ConvertFrom-Json



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   General error definition                                                      ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$GeneralErrorDef = '
[
  { "No" :   1, "Description" : "Log file archiving" },
  { "No" :   2, "Description" : "Config file reading" },
  { "No" :   4, "Description" : "Importing JSON from previous session" },
  { "No" :   8, "Description" : "Connecting to IDAM API" },
  { "No" :  16, "Description" : "Data returned from IDAM API" },
  { "No" :  32, "Description" : "MSOL connection" },
  { "No" :  64, "Description" : "Exchange on-premises connection" },
  { "No" : 128, "Description" : "Cannot get the info about licenses" },
  { "No" : 256, "Description" : "Shortage of licenses" },
  { "No" : 512, "Description" : "Cannot cleanup log archive folder" }
]' | ConvertFrom-Json



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   User error definition                                                         ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$UserErrorDef = '
[
  { "No" :   1, "Description" : "Parsing from file" },
  { "No" :   2, "Description" : "Parsing from API" },
  { "No" :   4, "Description" : "IDAM role to o365 license conversion" },
  { "No" :   8, "Description" : "Station list BAD_FORMAT" },
  { "No" :  16, "Description" : "Reading AD data" },
  { "No" :  32, "Description" : "AD account disabled while license required" },
  { "No" :  64, "Description" : "User/station assignment failure" },
  { "No" : 128, "Description" : "Licensing or MFA enabling error" },
  { "No" : 256, "Description" : "Recipient verification/modification" }
]' | ConvertFrom-Json



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Processing steps definition                                                   ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$ProcessingStepsDef = '
[
  { "No" :  0, "Description" : "Import from file/API" },
  { "No" :  1, "Description" : "Converting IDAM role to Office 365 License" },
  { "No" :  2, "Description" : "Converting Distribution Group to list of stations" },
  { "No" :  3, "Description" : "Checking Active Directory data" },
  { "No" :  4, "Description" : "Stations assignment" },
  { "No" :  5, "Description" : "Checking current Office 365 license" },
  { "No" :  6, "Description" : "Removing unnecessary Office 365 license" },
  { "No" :  7, "Description" : "Assigning Office 365 license" },
  { "No" :  8, "Description" : "Verifying recipient type" }
]' | ConvertFrom-Json




   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Define 'User' object type                                                     ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

function New-UserObject
{
    $u = New-Object PSObject
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName ShortID             # In AD: SamAccountName
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName Alias               # Email alias
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName Enabled             # Account status in AD
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName UserPrincipalName   # Needed ID in MS Online
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName DistinguishedName   # Common ID for AD, MSOL and on-prem Exchange
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName OU                  # Organization Unit derived from Canonical Name
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName EffectiveFrom       # Date/time of the change taking effect
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName IdamRole            # IDAM role
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName AsIsLicense         # Currenet Office 365 license
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName ToBeLicense         # Desired Office 365 license
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName Recipient           # Recipient type details
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName MFA                 # Multi-factor authentication state
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName DistGroup           # Station assignment according to API
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName AsIsStations        # Current assignment to stations
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName ToBeStations        # Desired assignment to stations
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName Attempts            # How many times was the account processed?
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName LastStep            # What was the last processing step?
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName FirstAttempt        # When account processing first started (in this session)?
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName LastAttempt         # When was the last processing attempt?
    $u | Add-Member -NotePropertyValue @()    -NotePropertyName Message             # Message from last processing step
    $u | Add-Member -NotePropertyValue $null  -NotePropertyName Source              # 0 - file from previous attempts; 1 - IDAM API
    $u | Add-Member -NotePropertyValue 0      -NotePropertyName Error               # Binary flags determining errors
    $u | Add-Member -NotePropertyValue 0      -NotePropertyName ErrorReported       # Binary flags: 1 - Admin; 2 - IDAM; 4 - BAD_FORMAT
    
    return $u
}

#endregion



#region DATA IMPORT


   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Get data from previous sessions                                               ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

# Read the file
$UsersFile = "$PSScriptRoot\$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)).Users.JSON"
$UTmp = $null
try
{
    if (Test-Path $UsersFile)
    {
        $LogEntry += "`r`n`r`nACTION:`r`nReading users from file '$($UsersFile)':"
        
        # The built-in cmdlet 'ConvertFrom-Json' has a size limitation for input length of 2MB
        #$UTmp = Get-Content $UsersFile -Raw | ConvertFrom-Json

        # Use different approach that allows increasing the input string length to 2GB
        $jsonserial = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer
        $jsonserial.MaxJsonLength = [System.Int32]::MaxValue
        $UTmp = $jsonserial.DeserializeObject((Get-Content $UsersFile -Raw))
        Remove-Variable jsonserial
    }
}
catch
{
    $LogEntry += "`r`n`r`nERROR`r`nError while reading users from file '$UsersFile'.`r`n$($Error[0].Exception.Message)"
    $GeneralError = $GeneralError -bor 4
}


# Iterate through file contents
foreach ($i in $UTmp)
{
    $u = New-UserObject

    try
    {
        $LogEntry += "`r`n  - $($i.UserPrincipalName.PadRight(42))"

        $u.Source = 0 -as [System.Byte]
        $u.Error = 0 -as [System.UInt16]
        $u.Message = $i.Message
        $u.LastAttempt = $Now
        $u.LastStep = 0 -as [System.UInt16]
        $u.UserPrincipalName = $i.UserPrincipalName
        $u.IdamRole = $i.IdamRole
        $u.DistGroup = $i.DistGroup
        $u.Attempts = $i.Attempts -as [System.UInt16]
        $u.EffectiveFrom = $i.EffectiveFrom -as [System.DateTime]
        $u.FirstAttempt = $i.FirstAttempt -as [System.DateTime]
        $u.ErrorReported = $i.ErrorReported -as [System.Byte]

        $LogEntry += "OK"
    }
    catch
    {
        $u.Error = $u.Error -bor 1
        $LogEntry += "Parsing error: $($Error[0].Exception.Message)"
        $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  ERROR - Parsing from file: $($Error[0].Exception.Message)"
    }
    finally
    {
        $Users += $u
        Remove-Variable i,u
    }
}
Remove-Variable UTmp




   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Get data from IDAM API                                                        ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$DateFrom = ($Config.LastDateFrom -as [System.DateTime]).ToString("yyyy-MM-dd'%20'HH'%3A'mm'%3A'ss")
$DateTo = ($Config.LastDateTo -as [System.DateTime]).ToString("yyyy-MM-dd'%20'HH'%3A'mm'%3A'ss")


##### DEBUG #####
#$DateFrom = "2017-04-06 00:00:00" -as [System.DateTime]
#$DateTo = "2017-04-06 13:00:00" -as [System.DateTime]
#[System.Uri]::EscapeUriString($DateFrom.ToString("yyyy-MM-dd HH:mm:ss"))
#$ApiUrl = "http://idmohs-int-pub.statoilfuelretail.com:10229/CircleKApi/OIDMOffice365Service/"   ## Correct
#$ApiUrl = "http://idmohs-int-pub.statoilfuelretail.com:10229/CircleKApi1/OIDMOffice365Service/"  ## Error



$Uri = "$ApiUrl$DateFrom/$DateTo/"
$Uri = Add-Member -InputObject $Uri -MemberType ScriptProperty -Value { [System.Uri]::UnescapeDataString($this) } -Name Unescape -PassThru

#Write-Host ([System.Uri]::UnescapeDataString($Uri)) -ForegroundColor Cyan   ## --- DEBUG ---


# Connect to API
try
{
    $R = Invoke-RestMethod -Uri $Uri
    
    ### DEBUG ###
    #$R = Get-Content "$PSScriptRoot\_fake_API.txt" -Raw | ConvertFrom-Json  ### DEBUG
    
    $LogEntry += "`r`n`r`nACTION:`r`nData from IDAM API retrieved without server/connection errors.`r`n$($Uri.Unescape)"
}
catch
{
    $GeneralError = $GeneralError -bor 8
    $LogEntry += "`r`n`r`nERROR:`r`nCannot retrieve data from IDAM API using the URI:`r`n$($Uri.Unescape)`r`n$($Error[0].Exception.Message)"
    $IdamErrMsg += "`r`nUser provisioning process for Office 365 cannot retrieve data from IDAM API using the URI:`r`n`r`n$($Uri.Unescape)`r`n`r`nERROR MESSAGE:`r`n$($Error[0].Exception.Message)`r`n"
}


# Read from API
if (($GeneralError -band 8) -eq 0 -and $R.status -ne "OK")
{
    $GeneralError = $GeneralError -bor 16
    $LogEntry += "`r`n`r`nERROR:`r`nConnection to API was successful for URI '$($Uri.Unescape)', but other errors occurred:`r`nStatus : $($R.status)`r`r$($R.statusInfo)"
    $IdamErrMsg += "`r`nUser provisioning process for Office 365 successfully retrieved data from IDAM API using the URI:`r`n`r`n$($Uri.Unescape)`r`n`r`n"
    $IdamErrMsg += "The data is however marked as error - see the JSON below.$("`r`n`r`n".PadRight(78,"="))`r`n$($R | ConvertTo-Json)$("`r`n".PadRight(78,"="))"
}
else
{
    $LogEntry += "`r`n`r`nACTION:`r`nImporting users ($($R.results.Count)) from API:"
    foreach ($i in $R.results)
    {
        $u = New-UserObject

        try
        {
            $LogEntry += "`r`n  - $($i.userId.PadRight(42))"

            $u.Source = 1 -as [System.Byte]
            $u.Error = 0 -as [System.UInt16]
            $u.ErrorReported = 0 -as [System.Byte]
            $u.Message = @()
            $u.Attempts = 0 -as [System.UInt16]
            $u.LastStep = 0 -as [System.UInt16]
            $u.FirstAttempt = $Now
            $u.LastAttempt = $Now
            $u.UserPrincipalName = $i.userId.Trim()
            $u.IdamRole = $i.role.Trim()
            $u.DistGroup = $i.distGroup
            $u.AsIsStations = @()
            $u.ToBeStations = @()
            $u.AsIsLicense = $null
            $u.ToBeLicense = $null
            $u.EffectiveFrom = $i.provStartDate -as [System.DateTime]

            $LogEntry += "OK"

        }
        catch
        {
            $u.Error = $u.Error -bor 2
            $LogEntry += "User parsing error: $($Error[0].Exception.Message)"
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  ERROR - Parsing from API: $($Error[0].Exception.Message)"
        }
        finally
        {
            $Users += $u
            Remove-Variable i,u
        }
    }
}



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Remove duplicates                                                             ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$dbl = $Users | Group-Object -Property UserPrincipalName
$Users = @()

# Add singletons to new table
$Users += ($dbl | ? Count -EQ 1).Group | ? { $_ -ne $null }

# Handle duplicates
$dbl = $dbl | ? Count -GT 1
foreach ($i in $dbl)
{
    $i = $i.Group
    $top = $i | sort Source,LastAttempt -Descending | select -First 1
    $top.Attempts = ($i | Measure-Object -Property Attempts -Sum).Sum
    $top.FirstAttempt = ($i | Measure-Object -Property FirstAttempt -Minimum).Minimum
    $top.Message = @() + (($i | sort Source,LastAttempt).Message)
    
    $Users += $top
    Remove-Variable i,top
}
Remove-Variable dbl



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Parse input data (licenses and stations)                                      ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$Log1 = "`r`n`r`nConverting IDAM roles to Office 365 licenses"
$Log2 = "`r`n`r`nConverting distribution groups to stations list"
foreach ($u in $Users)
{
    $u.Attempts++
    $Log1 += "`r`n  - $($u.UserPrincipalName.PadRight(42))"
    $Log2 += "`r`n  - $($u.UserPrincipalName.PadRight(42))"
    
    # Convert IDAM role to Office 365 license
    $u.LastStep = 1
    if ($u.IdamRole -in ($Lic.IdamRole))
    {
        $u.ToBeLicense = ($Lic | ? { $_.IdamRole -EQ "$($u.IdamRole)"} ).AccountSkuId
        $Log1 += "OK"
        $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  IDAM role '$($u.IdamRole)' correctly parsed"
    }
    else
    {
        $u.Error = $u.Error -bor 4
        $Log1 += "Unknown IDAM role '$($u.IdamRole)'"
        $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Unknown IDAM role: '$($u.IdamRole)'"
    }


    # Parse station assignment
    $u.LastStep = 2
    if ([System.String]::IsNullOrEmpty($u.DistGroup))
    {
        $u.ToBeStations = @()
        $Log2 += "(none)"
        $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  No TO-BE stations assignment"
    }
    elseif ($u.DistGroup -like "*BAD_FORMAT*")
    {
        $u.Error = $u.Error -bor 8
        $u.ToBeStations = @()
        $Log2 += "BAD_FORMAT"
        $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  BAD_FORMAT in stations assignment"
    }
    else
    {
        $u.ToBeStations = @() + ($u.DistGroup -split ",").Trim()
        $Log2 += "OK"
        $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Stations assignment parsed"

    }
    Remove-Variable u
}
$LogEntry += "`r`n`r`nPARSING:$Log1$Log2"
Remove-Variable Log1,Log2


#endregion



#region USER PROCESSING



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Retrieve Active Directory data                                                ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$LogEntry += "`r`n`r`nACTION:`r`nChecking Active Directory data:"
foreach ($u in $Users)
{
    $u.LastStep = 3
    $LogEntry += "`r`n  - $($u.UserPrincipalName.PadRight(42))"
    $ad = Get-ADUser -LDAPFilter "(UserPrincipalName=$($u.UserPrincipalName))" -Properties mailNickname,CanonicalName,MemberOf
    if ($ad -eq $null)
    {
        $u.Error = $u.Error -bor 16
        $LogEntry += "User NOT found in Active Directory"
        $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  User NOT found in Active Directory"
    }
    else
    {
        $u.Enabled = $ad.Enabled
        
        if (!$u.Enabled -and $u.ToBeLicense -ne "#NONE#" -and $u.ToBeLicense -ne $null)
        {
            # ERROR if account disabled while license required
            $u.Error = $u.Error -bor 32
            $LogEntry += "AD account disabled while license required"
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  AD account disabled while license required"
        }
        else
        {
            $LogEntry += "OK"
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  User found in Active Directory"
        }


        
        $u.ShortID = $ad.SamAccountName
        $u.DistinguishedName = $ad.DistinguishedName
        $u.OU = ($ad.CanonicalName | Select-String ".*(?=/.*)").Matches[0].Value
        if ($ad.mailNickname -eq $null) { $u.Alias = $ad.SamAccountName } else { $u.Alias = $ad.mailNickname }
        $s = (($ad.MemberOf | Get-ADGroup | ? SamAccountName -Match "\d{5}_User").SamAccountName | Select-String "\d{5}(?=_User)").Matches.Value
        if ($s -eq $null) { $u.AsIsStations = @() } else { $u.AsIsStations = @() + $s }
                
    }
    Remove-Variable ad
}





   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Handle assignment to station mailboxes                                        ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$LogEntry += "`r`n`r`nSTATIONS:`r`nStations assignment"
# Error converting IDAM role to Office 365 license and disabled account error should be ignored (65535 - 32 - 4 = 65499)
foreach ($u in $Users | ? { ($_.Error -band 65499) -eq 0 -and ( $_.AsIsStations.Count -gt 0 -or $_.ToBeStations.Count -gt 0 ) })
{
    $u.LastStep = 4
    $LogEntry += "`r`n  - $($u.UserPrincipalName)"
    
    $Add = @() + $u.ToBeStations | ? { $_ -notin $u.AsIsStations }
    $Del = @() + $u.AsIsStations | ? { $_ -notin $u.ToBeStations }

    if ($Del.Count -eq 0 -and $Add.Count -eq 0)
    {
        $LogEntry += "`r`n    No changes in stations assignment."
        $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  No changes in stations assignment"
    }
    else
    {
        if ($Del.Count -gt 0)
        {
            $LogEntry += "`r`n    Removing station assignment : $($Del -join ",")"
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Removing station assignment : $($Del -join ",")"
                
            $Del | ForEach-Object {
                $grp = "$($_)_User"
                try
                {
                    if ($DEBUG)
                    {
                        "Remove-ADGroupMember -Identity $grp -Members $($u.DistinguishedName) -Confirm:`$false" | Out-File "$PSScriptRoot\DEBUG.Commands.txt" -Append
                    }
                    else
                    {
                        Remove-ADGroupMember -Identity $grp -Members $u.DistinguishedName -Confirm:$false
                    }
                }
                catch
                {
                    $u.Error = $u.Error -bor 64
                    $LogEntry += "`r`n    ERROR removing from group '$grp'"
                    $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  ERROR removing from group '$grp'; $($Error[0].Exception.Message)"
                }
                Remove-Variable grp
            }
        }

        if ($Add.Count -gt 0)
        {
            $LogEntry += "`r`n    Adding station assignment   : $($Add -join ",")"
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Adding station assignment   : $($Add -join ",")"
                
            $Add | ForEach-Object {
                $grp = "$($_)_User"
                try
                {
                    if($DEBUG)
                    {
                        "Add-ADGroupMember -Identity $grp -Members $($u.DistinguishedName) -Confirm:`$false" | Out-File "$PSScriptRoot\DEBUG.Commands.txt" -Append
                    }
                    else
                    {
                        Add-ADGroupMember -Identity $grp -Members $u.DistinguishedName -Confirm:$false
                    }
                }
                catch
                {
                    $u.Error = $u.Error -bor 64
                    $LogEntry += "`r`n    ERROR adding to group '$grp'"
                    $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  ERROR adding to group '$grp'; $($Error[0].Exception.Message)"
                }
                Remove-Variable grp
            }
        }
    }
    Remove-Variable Add,Del,u
}




   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Connecting to MS Online services                                              ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

Get-MsolDomain -ErrorAction SilentlyContinue | Out-Null
#Connect-MsolService -Credential $OnlineCred -ErrorAction Stop
if(!$?)
{
    $LogEntry += "`r`n`r`nACTION:`r`nConnecting to MSOL services                   "
    Connect-MsolService -Credential $OnlineCred -ErrorAction Stop
    if ($?)
    {
        $LogEntry += "OK"
    }
    else
    {
        $GeneralError = $GeneralError -bor 32
        $LogEntry += "ERROR - Cannot connect to MS Online services`r`n$("""".PadLeft(46))$($Error[0].Exception.Message)"
    }
}




   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Connecting to on-premises Exchange Server                                     ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

if ($OPSession -eq $null -or $OPSession.State -ne "Opened")
{
    $LogEntry += "`r`n`r`nACTION:`r`nConnecting to on-prem Exchange Server         "
    try
    {
        # Output: $OPSession
        Invoke-Expression $OPCmd
        $LogEntry += "OK"
    }
    catch
    {
        $GeneralError = $GeneralError -bor 64
        $LogEntry += "ERROR - Cannot connect to on-premises Exchange Server`r`n$("""".PadLeft(46))$($Error[0].Exception.Message)"
    }
}



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Only manage the licenses if successfully connected to the services            ##
##                   (32 + 64 = 96)                                                ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

if (($GeneralError -band 96) -eq 0)
{

       ###############################################################################
     ###################################################################################
    ###                                                                               ###
    ##   Check Office 365 licenses availability pool                                   ##
    ###                                                                               ###
     ###################################################################################
       ###############################################################################

    try
    {
        $LogEntry += "`r`n`r`nLICENSES:`r`nChecking Office 365 Licenses Availability"
        
        $OLLic = Get-MsolAccountSku
        foreach ($i in $Lic | ? Name -NE $null)
        {
            $TmpLic = $OLLic | ? { $_.AccountSkuId -eq ($i.AccountSkuId) }
            if ($TmpLic -ne $null)
            {
                $i.Available = $TmpLic.ActiveUnits - $TmpLic.ConsumedUnits
            }
            else
            {
                $GeneralError = $GeneralError -bor 128
                $LogEntry += "`r`nERROR - Cannot get SKU info for license '$($i.Name)'"
            }

            Remove-Variable i,TmpLic
        }

        Remove-Variable OLLic
    }
    catch
    {
        $LogEntry += "`r`nERROR - Cannot get the information about the licenses from MSOnline"
    }





       ###############################################################################
     ###################################################################################
    ###                                                                               ###
    ##   Check actual Office 365 licenses assignment                                   ##
    ###                                                                               ###
     ###################################################################################
       ###############################################################################

    $LogEntry += "`r`n`r`nLICENSES:`r`nChecking current users license assignment:"
    # Stations assignment and disabled account errors not relevant here (error bits: 65535 - 8 - 32 - 64 = 65431)
    foreach ($u in $Users | ? { ($_.Error -band 65431) -eq 0 } )  
    {
        try
        {
            $u.LastStep = 5
            $LogEntry += "`r`n  - $($u.UserPrincipalName.PadRight(42))"

            #MFA
            $MSOLUser = Get-MsolUser -UserPrincipalName $u.UserPrincipalName
            $u.AsIsLicense = $MSOLUser.Licenses.AccountSkuId | ? { $_ -in $Lic.AccountSkuId }
            $u.MFA = $MSOLUser.StrongAuthenticationRequirements.State
            if ($u.MFA -eq $null) { $u.MFA = "Disabled" }
            Remove-Variable MSOLUser

            $LogEntry += "$($u.AsIsLicense)"
        }
        catch
        {
            $u.Error = $u.Error -bor 128
            $LogEntry += "ERROR - Cannot determine the license or MFA`r`n$("""".PadRight(46))($Error[0].Exception.Message)"
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  ERROR: Cannot determine the current user license or MFA state"
        }
        Remove-Variable u
    }

    


       ###############################################################################
     ###################################################################################
    ###                                                                               ###
    ##   Calculate the number of needed licenses                                       ##
    ###                                                                               ###
     ###################################################################################
       ###############################################################################

    foreach ($i in $Lic | ? Name -ne $null)
    {
        $i.Needed = ($Users | ? { ($_.Error -band 32) -ne 32 -and $_.Enabled -and $_.ToBeLicense -EQ $i.AccountSkuId } | Measure-Object).Count - ($Users | ? { ($_.Error -band 32) -ne 32 -and $_.ToBeLicense -ne $null -and $_.AsIsLicense -EQ $i.AccountSkuId } | Measure-Object).Count
        Remove-Variable i

    }
    $LicShortage = @() + ($Lic | ? { $_.Name -ne $null -and $_.Available -lt $_.Needed })
    if ($LicShortage.Count -gt 0)
    {
        $GeneralError = $GeneralError -bor 256
        $LogEntry += "`r`n`r`nWARNING`r`nThere are too few licenses available:" + ( $LicShortage | Format-List Name,Available,Needed,@{N="Shortage";E={$_.Available - $_.Needed}} | Out-String )
    }
    Remove-Variable LicShortage



       ###############################################################################
     ###################################################################################
    ###                                                                               ###
    ##   Remove unnecessary license assignments                                        ##
    ###                                                                               ###
     ###################################################################################
       ###############################################################################

    $LogEntry += "`r`n`r`nLICENSES:`r`nRemoving Office 365 license assignment"
    # Stations assignment and disabled account errors not relevant here (error bits: 65535 - 64 - 8 = 65463)
    foreach ($u in $Users | ? { ($_.Error -band 65463) -eq 0 -and $_.ToBeLicense -ne $null -and $_.AsIsLicense -ne $null -and $_.ToBeLicense -ne $_.AsIsLicense } )
    {
        $LogEntry += "`r`n  - $($u.UserPrincipalName.PadRight(42))"
        try
        {
            $u.LastStep = 6
            $LogEntry += "Removing license '$($u.AsIsLicense)'"
            if ($DEBUG)
            {
                "Set-MsolUserLicense -UserPrincipalName $($u.UserPrincipalName) -RemoveLicenses $($u.AsIsLicense)" | Out-File "$PSScriptRoot\DEBUG.Commands.txt" -Append
            }
            else
            {
                if ($PILOT -and $u.ToBeLicense -eq "#NONE#")
                {
                    "$($Now.ToString("yyyy-MM-dd HH:mm:ss")),$($u.UserPrincipalName),$($u.IdamRole),$($u.Enabled),$($u.AsIsLicense)" | Out-File $NO_MAILBOX_CSV -Append
                }
                else
                {
                    Set-MsolUserLicense -UserPrincipalName $u.UserPrincipalName -RemoveLicenses $u.AsIsLicense -ErrorAction Stop
                }
            }
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Removing license '$($u.AsIsLicense)'"
        }
        catch
        {
            $u.Error = $u.Error -bor 128
            $LogEntry += "ERROR - Cannot remove license '$($u.AsIsLicense)'`r`n$("""".PadLeft(46))$($Error[0].Exception.Message)"
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  ERROR - Cannot remove license '$($u.AsIsLicense)'"
        }
        Remove-Variable u
    }




       ###############################################################################
     ###################################################################################
    ###                                                                               ###
    ##   Assign licenses                                                               ##
    ###                                                                               ###
     ###################################################################################
       ###############################################################################

    $LogEntry += "`r`n`r`nLICENSES:`r`nAssigning Office 365 licenses"
    # Stations assignment errors not relevant here (error bits: 65535 - 64 - 8 = 65463)
    foreach ($u in $Users | ? { ($_.Error -band 65463) -eq 0 -and $_.ToBeLicense -ne $null -and $_.ToBeLicense -ne "#NONE#" -and $_.ToBeLicense -ne $_.AsIsLicense } )
    {
        $LogEntry += "`r`n  - $($u.UserPrincipalName.PadRight(42))"
        try
        {
            $u.LastStep = 7
            $LogEntry += "Assigning license '$($u.ToBeLicense)'"
            if ($DEBUG)
            {
                "Set-MsolUserLicense -UserPrincipalName $($u.UserPrincipalName) -AddLicenses $($u.ToBeLicense)" | Out-File "$PSScriptRoot\DEBUG.Commands.txt" -Append
                "Set-MsolUser -UserPrincipalName $($u.UserPrincipalName) -StrongAuthenticationRequirements `$mfa" | Out-File "$PSScriptRoot\DEBUG.Commands.txt" -Append
            }
            else
            {
                # Assign Office 365 license
                Set-MsolUser -UserPrincipalName $u.UserPrincipalName -UsageLocation $UsageLocation -ErrorAction Stop
                Set-MsolUserLicense -UserPrincipalName $u.UserPrincipalName -AddLicenses $u.ToBeLicense -ErrorAction Stop

                # Enable MFA (if so configured)
                if ($EnableMFA -and $u.MFA -eq "Disabled")
                {
                    $mf = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
                    $mf.RelyingParty = "*"
                    $mf.State = "Enabled"
                    $mf.RememberDevicesNotIssuedBefore = $Now
                    $mfa = @($mf)

                    Set-MsolUser -UserPrincipalName $u.UserPrincipalName -StrongAuthenticationRequirements $mfa

                    Remove-Variable mf,mfa
                }
            }
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Assigning license '$($u.ToBeLicense)'"
        }
        catch
        {
            $u.Error = $u.Error -bor 128
            $LogEntry += "  ERROR - Cannot assign license '$($u.ToBeLicense)' and/or enable MFA`r`n$("""".PadLeft(46))$($Error[0].Exception.Message)"
            $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  ERROR - Cannot assign license '$($u.ToBeLicense)' and/or enable MFA"
        }
        Remove-Variable u
    }




       ###############################################################################
     ###################################################################################
    ###                                                                               ###
    ##   Verifying recipient type                                                      ##
    ###                                                                               ###
     ###################################################################################
       ###############################################################################

    $LogEntry += "`r`n`r`nRECIPIENTS:`r`nVerifying recipient types"
    # Stations assignment errors not relevant here (error bits: 65535 - 64 - 8 = 65463)
    foreach ($u in $Users | ? { ($_.Error -band 65463) -eq 0 -and $_.ToBeLicense -ne $null })
    {
        $usr = $null
        try
        {
            $u.LastStep = 8
            $LogEntry += "`r`n  - $($u.UserPrincipalName.PadRight(42))"
            $usr = Invoke-Command $OPSession { Get-User -Identity $Using:u.DistinguishedName }
            $u.Recipient = "$($usr.RecipientType),$($usr.RecipientTypeDetails)"
            $LogEntry += $u.Recipient.PadRight(28)

            if ($u.Recipient -in "MailUser,RemoteUserMailbox","User,User","User,DisabledUser")
            {
                $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Recipient type: $($usr.RecipientType),$($usr.RecipientTypeDetails)"

                #f ($u.ToBeLicense -eq "#NONE#" -and $usr.RecipientTypeDetails -ne "User")
                if ($u.ToBeLicense -eq "#NONE#" -and $usr.RecipientTypeDetails -eq "RemoteUserMailbox")
                {
                    $LogEntry += "- Disabling remote mailbox"
                    $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Disabling remote mailbox"
                    if ($DEBUG)
                    {
                        "Disable-RemoteMailbox -Identity ""$($u.DistinguishedName)"" -Confirm:`$false" | Out-File "$PSScriptRoot\DEBUG.Commands.txt" -Append
                    }
                    else
                    {
                        if (!$PILOT)
                        {
                            Invoke-Command $OPSession { Disable-RemoteMailbox -Identity $Using:u.DistinguishedName -Confirm:$Using:false } -ErrorAction Stop
                        }
                    }
                }
                elseif ($u.ToBeLicense -ne "#NONE#" -and $usr.RecipientTypeDetails -ne "RemoteUserMailbox")
                {
                    $LogEntry += "- Enabling remote mailbox"
                    $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Enabling remote mailbox"
                    $RoutingAddress = "$($u.ShortID)@$($RemoteDomain)"
                    if ($DEBUG)
                    {
                        "Enable-RemoteMailbox -Identity ""$($u.DistinguishedName)"" -RemoteRoutingAddress $RoutingAddress -Confirm:`$false" | Out-File "$PSScriptRoot\DEBUG.Commands.txt" -Append
                    }
                    else
                    {
                        Invoke-Command $OPSession { Enable-RemoteMailbox -Identity $Using:u.DistinguishedName -RemoteRoutingAddress $Using:RoutingAddress -Confirm:$Using:false } -ErrorAction Stop
                    }
                    Remove-Variable RoutingAddress
                }
                else
                {
                    $LogEntry += "- OK"
                }

            }
            else
            {
                $u.Error = $u.Error -bor 256
                $LogEntry += "- WARNING: Unexpected recipient type"
                $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  Unexpected recipient type: $($usr.RecipientType),$($usr.RecipientTypeDetails)"
            }
        }
        catch
        {
                $u.Error = $u.Error -bor 256
                $LogEntry += "`r`n    ERROR - cannot verify/modify recipient type`r`n    $($Error[0].Exception.Message)"
                $u.Message += "[$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]  ERROR - Cannot verify/modify recipient type"
        }
        Remove-Variable u,usr
    }



}


#endregion






#region NOTIFICATIONS & LOGGING


   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Email alert to administrator - general errors                                 ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

if ($GeneralError -ne 0)
{
    $Msg = "Hello,`r`n`r`nThe following general errors occurred during user processing at [$($Now.ToString("yyyy-MM-dd HH:mm:ss"))]:`r`n"
    foreach ($e in $GeneralErrorDef)
    {
        if (($GeneralError -band $e.No) -eq $e.No) { $Msg += "`r`n  - $($e.Description)" }
        Remove-Variable e
    }
    $Msg += "`r`n`r`nBest regards,`nUser Provisioning"
    Send-MailMessage -SmtpServer $SmtpServer -From $From -Subject "User Provisioning General Error" -To $ToAdmin -Body $Msg
    Remove-Variable Msg
}



   ###############################################################################
 ###################################################################################
###                                                                               ###
##  Email alert to IDAM - general error                                            ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

if (($GeneralError -band 24) -ne 0)
{
    $IdamErrMsg = "Hello,`r`n`r`n" + $IdamErrMsg + "`r`n`r`nBest regards,`nUser Provisioning"
    Send-MailMessage -SmtpServer $SmtpServer -From $From -Subject "IDAM API - Connection Errors" -To $ToIdam -Body $IdamErrMsg
}



   ###############################################################################
 ###################################################################################
###                                                                               ###
##  Email alert to administrator - user processing errors                          ##
##  Skip the following errors:                                                     ##
##    •   8 - Station list BAD_FORMAT                                              ##
##    •  32 - AD account disabled while license required                           ##
##    •  64 - User/station assignment failure                                      ##
###                                                                               ###
 ###################################################################################
   ###############################################################################


$Msg = ""; $i = 0
$ErrMask = (511 - 8 - 32 - 64) -as [System.UInt16]   ##  (mask to skip some errors)
foreach ($u in ($Users | ? { ($_.Error -band $ErrMask) -ne 0 -and ($_.ErrorReported -band 1) -eq 0 -and ($Now - $_.FirstAttempt).TotalHours -gt $AlertTimeSpanHrs -and $_.Attempts -gt $AlertAttempts } ))
{
    $i++
    $Msg += "`r`n".PadRight(78,"=")
    $Msg += $u | fl UserPrincipalName,
                    OU,
                    Enabled,
                    @{N="EffectiveFrom";E={if($_.EffectiveFrom -ne $null){$_.EffectiveFrom.ToString("yyyy-MM-dd HH:mm:ss")}}},
                    IdamRole,
                    AsIsLicense,
                    ToBeLicense,
                    Recipient,
                    DistGroup,
                    AsIsStations,
                    ToBeStations,
                    Attempts,
                    @{N="FirstAttempt";E={if($_.FirstAttempt -ne $null){$_.FirstAttempt.ToString("yyyy-MM-dd HH:mm:ss")}}},
                    @{N="LastAttempt";E={if($_.LastAttempt -ne $null){$_.LastAttempt.ToString("yyyy-MM-dd HH:mm:ss")}}},
                    @{N="LastStep";E={ ($ProcessingStepsDef | ? No -eq  $u.LastStep).Description }} | Out-String
        
    $Msg += "-------------------------------------------------`r`nErrors that ocurred during processing:`r`n"
        
    foreach ($e in $UserErrorDef)
    {
        if (($u.Error -band $e.No) -eq $e.No) { $Msg += "`r`n  - $($e.Description)" }
        Remove-Variable e
    }

    $Msg += "`r`n`r`n-------------------------------------------------`r`nMessages recorded during all processing attempts:`r`n"
    $u.Message | % { $Msg += "`r`n  $_" }

    $Msg += "`r`n`r`n"

    $u.ErrorReported = $u.ErrorReported -bor 1

    Remove-Variable u
}
Remove-Variable ErrMask

if ($i -gt 0)
{
    $Msg = "Hello,`r`n`r`nThere are ($i) user account(s) that have been failing provisioning process:`r`n" + $Msg + "$("`r`n`r`n".PadRight(78,"="))`r`n`r`nBest regards,`nUser Provisioning"
    Send-MailMessage -SmtpServer $SmtpServer -From $From -Subject "User Provisioning Processing Errors" -To $ToAdmin -Body $Msg
}

Remove-Variable Msg,i





   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Email alert to IDAM - user error (except for BAD_FORMAT)                      ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$Msg = ""; $i = 0
$IdamErr = (2 + 4 + 32) -as [System.UInt16]
foreach ($u in ($Users | ? { ($_.Error -band $IdamErr) -ne 0 -and ($_.ErrorReported -band 2) -eq 0 -and ($Now - $_.FirstAttempt).TotalHours -gt $AlertTimeSpanHrs -and $_.Attempts -gt $AlertAttempts } ))
{
    $i++
    $Msg += "`r`n`r`n`r`n".PadRight(78,"=")
    $Msg += $u | fl @{N="userId";E={$_.UserPrincipalName}},
                    @{N="provStartDate";E={if($_.EffectiveFrom -ne $null){$_.EffectiveFrom.ToString("yyyy-MM-dd HH:mm:ss")}}},
                    @{N="role";E={$_.IdamRole}},
                    @{N="distGroup";E={$_.DistGroup}},
                    @{N="AD Enabled";E={$_.Enabled}},
                    @{N="Org. Unit";E={$_.OU}} |
                    Out-String
        
    $Msg += "Errors that ocurred during processing:`r`n"

    foreach ($e in $UserErrorDef)
    {
        # Only IDAM-related errors
        if (($u.Error -band $IdamErr -band $e.No) -eq $e.No) { $Msg += "`r`n  - $($e.Description)" }
        Remove-Variable e
    }

    $u.ErrorReported = $u.ErrorReported -bor 2

    Remove-Variable u
}

if ($i -gt 0)
{
    $Msg = "Hello,`r`n`r`nThere are ($i) user account(s) that throw errors during provisioning process:`r`n" + $Msg + "$("`r`n`r`n".PadRight(78,"="))`r`n`r`nBest regards,`nUser Provisioning"
    Send-MailMessage -SmtpServer $SmtpServer -From $From -Subject "Office 365 Service - User Provisioning Errors" -To $ToIdam -Body $Msg

}
Remove-Variable Msg,i,IdamErr







   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Email alert to IDAM - BAD_FORMAT in station assignment                        ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

$Msg = ""; $i = 0
foreach ($u in ($Users | ? { ($_.Error -band 8) -ne 0 -and ($_.ErrorReported -band 4) -eq 0 -and ($Now - $_.FirstAttempt).TotalHours -gt $AlertTimeSpanHrs -and $_.Attempts -gt $AlertAttempts } ))
{
    $i++
    $Msg += "`r`n".PadRight(78,"=")
    $Msg += $u | fl @{N="userId";E={$_.UserPrincipalName}},
                    @{N="provStartDate";E={if($_.EffectiveFrom -ne $null){$_.EffectiveFrom.ToString("yyyy-MM-dd HH:mm:ss")}}},
                    @{N="role";E={$_.IdamRole}},
                    @{N="distGroup";E={$_.DistGroup}},
                    @{N="AD Enabled";E={$_.Enabled}},
                    @{N="Org. Unit";E={$_.OU}} |
                    Out-String

    $u.ErrorReported = $u.ErrorReported -bor 4

    Remove-Variable u
}

if ($i -gt 0)
{
    $Msg = "Hello,`r`n`r`nThere are ($i) user account(s) with BAD_FORMAT station assignment:`r`n`r`n`r`n" + $Msg + "$("`r`n`r`n".PadRight(78,"="))`r`n`r`nBest regards,`nUser Provisioning"
    Send-MailMessage -SmtpServer $SmtpServer -From $From -Subject "Office 365 Service - BAD_FORMAT in user/station assignment" -To $ToIdam -Body $Msg
}
Remove-Variable Msg,i






   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Email alert - licenses shortage                                               ##
###                                                                               ###
 ###################################################################################
   ###############################################################################
$LicErr = @() + ($Lic | ? { $_.Name -ne $null -and ($_.Available - $_.Needed) -le $_.AlertLevel })
if ($LicErr.Count -gt 0)
{
    $Msg = "Hello,`r`n`r`nThe number of available Office 365 licenses is low:`r`n`r`n"
    $Msg += $LicErr | ft -AutoSize Name,AccountSkuId,@{N="Available";E={if($_.Available -le $_.Needed){0}else{$_.Available - $_.Needed}}} | Out-String
    $Msg += "Please consider purchasing additional licenses.`r`n`r`nBest regards,`nUser Provisioning"

    Send-MailMessage -SmtpServer $SmtpServer -From $From -Subject "Shortage of Office 365 licenses!!!" -Priority High -To $ToLicence -Body $Msg

    Remove-Variable Msg
}
Remove-Variable LicErr





   ###############################################################################
 ###################################################################################
###                                                                               ###
##  Save the users for next attempt                                                ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

# Select users
if (($GeneralError -band 96) -ne 0)
{
    # If errors occurred when connectiong to MS Online services or
    # on-premises Exchange, select all users for next processing run
    $UsersToSave = $Users
}
else
{
    # If no connection errors occurred,
    # select only users that:
    #  • failed individually
    #  AND
    #  • have not exceeded the threshold for giving up
    #    OR
    #  • errors have not been reported yet
    $UsersToSave = @() + ($Users | ? { $_.Error -ne 0 -and (($Now - $_.FirstAttempt).TotalHours -lt $GiveUpTimeSpanHrs -or $_.Attempts -lt $GiveUpAttempts -or $_.ErrorReported -eq 0) } )
}

# Select only fields relevant for saving to JSON file
$UsersToSave = @() + $UsersToSave | select UserPrincipalName,DistGroup,IdamRole,
    @{N="EffectiveFrom";E={if($_.EffectiveFrom -eq $null){$null}else{$_.EffectiveFrom.ToString("yyyy-MM-dd HH:mm:ss")}}},
    @{N="FirstAttempt";E={if($_.FirstAttempt -eq $null){$null}else{$_.FirstAttempt.ToString("yyyy-MM-dd HH:mm:ss")}}},
    @{N="LastAttempt";E={if($_.LastAttempt -eq $null){$null}else{$_.LastAttempt.ToString("yyyy-MM-dd HH:mm:ss")}}},
    Recipient,Attempts,LastStep,Message,Error,ErrorReported

# Save to JSON file
try
{
    if ($UsersToSave.Count -eq 0)
    {
        $LogEntry += "`r`n`r`nALL PROCESSED:`r`nAll users processed."
        if (Test-Path $UsersFile) { Remove-Item $UsersFile }
    }
    else
    {
        $UsersToSave | ConvertTo-Json | Out-File $UsersFile
        $LogEntry += "`r`n`r`nPROCESSING COMPLETE`r`nSome users/accounts ($($UsersToSave.Count)) will require another attempt."
    }
}
catch
{
    $LogEntry += "`r`n`r`nERROR:`r`n$($Error[0].Exception.Message)"
}
finally
{
    Remove-Variable UsersToSave
}





   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Save to the log                                                               ##
###                                                                               ###
 ###################################################################################
   ###############################################################################

try
{
    "$LogEntry`r`n`r`n" | Out-File -LiteralPath $LogFile -Append
}
catch
{
    $LogEntry += "`r`n`r`nERROR:`r`nThe script cannot write to log file '$LogFile'.`r`n$($Error[0].Exception.Message)"
    $Msg = "Hello,`r`n`r`nThe user provisioning system is not able to write to the logfile '$LogFile'.`r`n`r`n"
    $Msg += "Error message:`r`n$($Error[0].Exception.Message)`r`n`r`nBest regards,`nUser Provisioning"
    Send-MailMessage -From $From -To $ToAdmin -SmtpServer $SmtpServer -Subject "Cannot write to log file" -Body $Msg
    Remove-Variable Msg
}



   ###############################################################################
 ###################################################################################
###                                                                               ###
##   Close remote sessions (if not run from PS ISE)                                ##
###                                                                               ###
 ###################################################################################
   ###############################################################################
if ($MyInvocation.CommandOrigin -eq "Runspace") { Get-PSSession | Remove-PSSession }


#endregion





## Data output (while debugging)

if ($MyInvocation.CommandOrigin -ne "Runspace")
{
    #$LogEntry
    #$Users | ft -AutoSize ShortID,Error,FirstAttempt,LastStep,Attempts,Message
    
    $Users | ft -AutoSize ShortID,UserPrincipalName,Error,LastStep,Attempts,@{N="FirstAttempt";E={$_.FirstAttempt.ToString("yyyy-MM-dd HH:mm:ss")}},@{N="LastAttempt";E={$_.LastAttempt.ToString("yyyy-MM-dd HH:mm:ss")}},@{N="TotalHours";E={("{0:N}" -f ($Now - $_.FirstAttempt).TotalHours).PadLeft(10)}},DistGroup,AsIsStations,ToBeStations
    $Users | ft -AutoSize ShortID,UserPrincipalName,Enabled,Error,ErrorReported,IdamRole,AsIsLicense,ToBeLicense,Recipient
    $Lic | ft -AutoSize

    $Users | select UserPrincipalName,DistGroup,IdamRole,
        @{N="EffectiveFrom";E={if($_.EffectiveFrom -eq $null){$null}else{$_.EffectiveFrom.ToString("yyyy-MM-dd HH:mm:ss")}}},
        @{N="FirstAttempt";E={if($_.FirstAttempt -eq $null){$null}else{$_.FirstAttempt.ToString("yyyy-MM-dd HH:mm:ss")}}},
        @{N="LastAttempt";E={if($_.LastAttempt -eq $null){$null}else{$_.LastAttempt.ToString("yyyy-MM-dd HH:mm:ss")}}},
        Recipient,Attempts,LastStep,Message,Error,ErrorReported | ConvertTo-Json | Out-File "$PSScriptRoot\ALL_USERS.JSON"
}







