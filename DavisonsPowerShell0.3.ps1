############################################################################################################################
#
# Fairsail Active Directory Synchronisation.
# Copyright (c) 2026 Sage Ltd
#
# This script is used on a server with Active Directory access to:
#	1. Authenticate with a Salesforce/Sage login
#   2. Fetch all Sage Employees
#   3. Compare the Flair Employees to the same users in AD and update the Sage Employees on the fields specified in $ADfieldToSageField, the users also need to have an employment status specified in $updateUsersWithTheseEmploymentStatuses
#   4. Create new users if they exist in Sage but not in AD if they have the required employment statuses specified in $createUsersWithTheseEmploymentStatuses
#   5. Disable users if they have their employment status is specified in $disableUsersWithTheseEmploymentStatuses
# Prerequisites:
#	1. The server can reach Active Directory and has the Powershell Active Directory Module
#	2. The user has permissions to access Active Directory and read and write to user accounts.
#	3. The parameters for Sage client, secret are set.
#	4. Remote Access is setup in your Sage/Salesforce org and the client id and secret are set.
#	5. The paths (Organisational Unit and Domain Controllers) for users in AD are set.
#   6. The $uniqueValueADandSage array must have a unique field in both AD and Sage.
#
#     
#############################################################################################################################

# Parameters
param(

    [string]$salesforceUrl = "https://davisonslaw.my.salesforce.com",
    [string]$client = "3MVG9WktJx3rjNu2vZrWMhn5uva5Rtmdrh94JcWuZzd_vIzxfaduYN32UZ62tf8S4YYoVYE4BEfBTc2ueqaL2",
    [string]$secret = "C9D43889FA069F1C75EFCC7D3CE66444BAD7A2FCEF57939E69A432EBAEC89F90",

    [string]$transcriptPath = $(if ($IsWindows) { "C:\scripts\HRSync\logs\transcript" } else { "$HOME/Downloads/HRSync/logs/transcript" }),
    [string]$logPath = $(if ($IsWindows) { "C:\scripts\HRSync\logs" } else { "$HOME/Downloads/HRSync/logs" }),
    
    [bool]$diag = $true, #$true/$false set to false when using a test user.
    [array]$paths = @("OU=Users,OU=Davisons,DC=davisons,DC=law","OU=Accounts,OU=Data,DC=davisons,DC=law"), #OUs to sync, comma delimited
    [string]$newUserPath = "OU=_NewJoiner,OU=Users,DC=davisons,DC=law", #OU to create new records in / only used if allowCreation is True
    
    # Feature Execution Switches (Client Phase Rollout Controllers)
    [bool]$allowCreation = $true,       # Set to $true to process and create new starters in AD
    [bool]$allowUpdate = $false,         # Set to $false to block updates to existing matching AD accounts
    [bool]$allowDeactivation = $false,   # Set to $false to block changing leaving employees to disabled
    
    [array]$createUsersWithTheseEmploymentStatuses = @(
		"Pre Joiner"
	),
    [array]$updateUsersWithTheseEmploymentStatuses = @(
		"Active",
		"Active Employee",
		"Pre-Joiner"
	),
	[bool]$disableUsersWhenLeft = $true, # Main toggle left intact for structural fallback
	[array]$disableUsersWithTheseEmploymentStatuses = @(
		"Leaver",
		"Left"
	),

    [array]$uniqueValueADandSage = @("employeeId","fHCM2__Unique_Id__c"), #First AD value, second Sage value
    [string]$testUser = $null, #Test user identifier
    [array]$SagefieldToADField = @{ #Map AD values to Sage fields, these are the only fields which will be synced to AD
        "Office"="fHCM2__Current_Employment__r.fHCM2__Work_Location__r.fHCM2__Address_City__c"
        "surname" = "fHCM2__Surname__c"
        "company" = "fHCM2__Current_Employment__r.fHCM2__Business_Name__c" 
    },

    [bool]$SyncEmailFromADToSage = $false, #Syncs the mail of the AD users to Work_Email_Address_Holding_Field__c
    $syncManagerFromSage = $true,

    #Mail Settings
    $shouldMailOnError = $true,
    $smtpFrom          = "hris-sync@davisons.law", 
    $smtpTo            = "support@davisons.law",  
    $smtpServer        = "smtp.davisons.law",    
    $smtpPort          = "587",
    $smtpUseSSL        = $true
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor
[Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

$workLocationAPIField = "fHCM2__Current_Employment__r.fHCM2__Work_Location__r.Name"

# Map the Sage Business Name to the correct Active Directory Domain Suffix
$domainMappingHash = @{
    "Davisons Law"             = "@davisons.law"
    "DP Law"                   = "@dplaw.law"
    "Default"                  = "@davisons.law" 
}

# String replacement mapping dictionary for specialized characters
$characterReplaceHash = @{
    "å" = "a"; "ä" = "a"; "ö" = "o"; "é" = "e"; "è" = "e"; "ü" = "u"; "ï" = "i"
}

########## DO NOT CHANGE ANYTHING BELOW THIS LINE ##########

Function WriteToLog($message){
    if($diag){
        $message = "[DIAGNOSTIC ON] " + (Get-Date -format "yyyyMMdd HH:mm") + " - " + $message
    }
    else{
        $message = (Get-Date -format "yyyyMMdd HH:ss") + " - " + $message
    }
    
    write-host $message
    $message | Out-File $logName -Force -Append
}

if ($IsWindows) {
    Import-Module ActiveDirectory
} else {
    Write-Host "Running on non-Windows OS. Skipping ActiveDirectory module import." -ForegroundColor Yellow
    
    # Mocking all necessary AD functions to prevent hangs
    function Get-ADUser { 
        [CmdletBinding()]
        param([Parameter(ValueFromPipelineByPropertyName=$true)]$Identity, $Filter) 
        return $null # Returns null, telling the script the user doesn't exist yet
    }
    function Set-ADUser { 
        param($Identity) 
        Write-Host "Mock Set-ADUser executed for $Identity" 
    }
    function New-ADUser { 
        param([Parameter(ValueFromPipelineByPropertyName=$true)]$SamAccountName, $name) 
        Write-Host "Mock New-ADUser created: $name" 
    }
    function Disable-ADAccount { 
        param($Identity) 
        Write-Host "Mock Disable-ADAccount executed for $Identity" 
    }
}

#set the script root
if ($psise) {
    $scriptRoot = Split-Path $psise.CurrentFile.FullPath
}
else {
    $scriptRoot = $global:PSScriptRoot
}

#Change values to LDAP names so they work with null values
$clearReplacementList = @{
    "state"="st"
    "city"="l"
    "Country"="C"
    "Office"="physicalDeliveryOfficeName"
}

[array]$allSageValues = @()
if($SagefieldToADField.Keys.count -gt 0){
    $allSageValues += $SagefieldToADField.values
}
$allSageValues += $uniqueValueADandSage[1]

$queryValues = @(
    "id"
    "fHCM2__First_Name__c"
    "fHCM2__Middle_Name__c"
    "fHCM2__Surname__c"
    "fHCM2__Preferred_Name__c"
    "fHCM2__Email__c"
    "fHCM2__Unique_Id__c"
    "fHCM2__Job_Title__c"
    "fHCM2__Division__c"
    "fHCM2__Employment_Status__c"
    "fHCM2__Manager__r.fHCM2__Unique_Id__c"
    "fHCM2__Current_Employment__r.Name"
    "fHCM2__Country__c"
    "fHCM2__Is_Manager__c"
    "fHCM2__Current_Employment__r.fHCM2__Business_Name__c"
)

foreach($val in $queryValues){
    if(!($allSageValues.Contains($val))){
        $allSageValues += $val
    }
}

$allSageValues += $workLocationAPIField
$allSageValues = $allSageValues | Sort-Object -Unique

$query = "SELECT " + ($allSageValues -join ",") + " FROM fHCM2__Team_Member__c WHERE fHCM2__Unique_Id__c != null"
if(!(Test-Path $transcriptPath)){
    new-item -ItemType Directory -Path $transcriptPath | Out-Null
}

if(!(Test-Path $logPath)){
    new-item -ItemType Directory -Path $logPath | Out-Null
}

$transcriptname = $transcriptPath + "\Transcript " + (Get-Date).ToString().Replace(":","-").Replace("/","-") + ".log"
$logName = $logPath + "\Output " + (Get-Date).ToString().Replace(":","-").Replace("/","-") + ".log"
Start-Transcript -Path $transcriptname

Function SendErrorMail($specificError = $null){
    if($shouldMailOnError){
        $body = $specificError
        if(!$body){
            if($error.count -gt 0){
                $body = $error | Out-String
            }
        }
        if($body){
            $subject = "$($error.count) Errors from SageScript - " + (Get-Date).ToString()
            
            try {
                $mail = New-Object System.Net.Mail.MailMessage
                $mail.From = $smtpFrom
                $mail.To.Add($smtpTo)
                $mail.Subject = $subject
                $mail.Body = $body

                $smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
                $smtp.EnableSsl = $smtpUseSSL
                
                # If mail infrastructure requires authentication, uncomment below:
                # $smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUsername, $smtpPassword)

                $smtp.Send($mail)
                $smtp.Dispose()
                $mail.Dispose()
            }
            catch {
                WriteToLog "Failed to send error notification email: $($_.Exception.Message)"
            }
        }
    }
}

Function GetAvailablesAMAccountName($givenName, $surname){
    foreach($key in $characterReplaceHash.Keys){
        $givenName = $givenName.Replace($key,($characterReplaceHash.$key))
        $surname = $surname.Replace($key,($characterReplaceHash.$key))
    }
    
    # Format Rule: First Initial + Surname Prefix (e.g., slee)
    $firstInitial = $givenName[0]
    $name = ($firstInitial + $surname)[0..19] -join ""
    $counter = 1
    
    while($counter -le 20){
        try{
            if(!(get-aduser -filter {sAMAccountName -eq $name} -ErrorAction Stop)){
                return $name.ToLower()
            }
        }
        catch{
            return $null
        }
        $name = ($firstInitial + $surname + $counter)[0..19] -join ""
        $counter++
    }
    WriteToLog "There are no available SAM with the parameters specified"
    return $null
}

Function GetAvailableUPN($givenName, $surname, $targetDomain){
    foreach($key in $characterReplaceHash.Keys){
        $givenName = $givenName.Replace($key,($characterReplaceHash.$key))
        $surname = $surname.Replace($key,($characterReplaceHash.$key))
    }
    
    if (!$targetDomain) { $targetDomain = $domainMappingHash["Default"] }

    # Format Rule: First Initial + Surname Suffix (e.g., slee@davisons.law)
    $firstInitial = $givenName[0]
    $name = ($firstInitial + $surname) + $targetDomain

    $counter = 1
    while($counter -le 20){
        try{
            if(!(get-aduser -filter {userprincipalname -eq $name} -ErrorAction Stop)){
                return $name.ToLower()
            }
        }
        catch{
            return $null
        }
        $name = ($firstInitial + $surname) + $counter + $targetDomain
        $counter++
    }
    WriteToLog "There are no available UPN with the parameters specified"
    return $null
}

Function TurnADuserIntoHash($aduser){
    $userhash = @{}
    $adUser | Get-Member -MemberType property | select -ExpandProperty name | ForEach-Object {
        $userhash[$_] = $adUser.$_
    }
    return $userhash
}

Function GetManagerSageID($manager){
    if($allUsersSage.($manager)){
        return $allUsersSage.($manager).id
    }
    return $null
}

[array] $usersCreated = @()
[array] $modifiedUsers = @()
[array] $disabledUsers = @()

if ($diag) {
    WriteToLog "Diagnostic mode ON. AD Users will not be affected."
}
else {
    WriteToLog "Diagnostic mode OFF. AD Users may be affected."
}

Set-Location $PSScriptRoot
WriteToLog "Script started"

try {
    WriteToLog "Authenticating with Salesforce"
    $sfLoginUrl = $salesforceUrl.TrimEnd('/') + "/services/oauth2/token"

    $session = Invoke-RestMethod -Uri $sfLoginUrl `
        -Method Post `
        -Headers @{ "Content-Type" = "application/x-www-form-urlencoded" } `
        -Body "grant_type=client_credentials&client_id=$client&client_secret=$secret" `
        -ErrorAction Stop
}
catch {
    try {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $authError = $reader.ReadToEnd()
        WriteToLog ("Error when trying to authenticate with Sage`n" + $authError)
    }
    catch {
        WriteToLog ("Error when trying to authenticate with Sage`n" + ($_ | Out-String))
    }

    WriteToLog "Script finished"
    break
}

$records = @()
$encodedQuery = [System.Web.HttpUtility]::UrlEncode($query)
$url = "$($session.instance_url)/services/data/v58.0/query?q=$encodedQuery"

do {
    try{
        $response = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "$($session.token_type) $($session.access_token)" } -ErrorAction Stop
        $records += $response.records
        $url = if ($response.nextRecordsUrl) { "$($session.instance_url)$($response.nextRecordsUrl)" } else { $null }
    }
    catch{
        WriteToLog ("Error when trying to fetch users from Sage`n" + ($_ | Out-String))
        WriteToLog "Script finished"
        return
    }
} while ($url -and $response.done -ne $true)

WriteToLog ("A total of " + $records.count + " records fetched from Sage")

# Process and log Employment Status Breakdown metrics
$statusCounts = @{}
$totalCount = 0
foreach ($record in $records) {
    $statusValue = $record.fHCM2__Employment_Status__c
    if ([string]::IsNullOrEmpty($statusValue)) { $statusValue = "[Blank/Null Status]" }
    $statusCounts[$statusValue]++
    $totalCount++
}

# Build the string for the log file
$breakdownString = "Employment Status Breakdown:`n"
foreach ($statusKey in $statusCounts.Keys | Sort-Object) {
    # Using ${statusKey} ensures the colon is treated as literal text
    $breakdownString += " - ${statusKey}: $($statusCounts[$statusKey])`n"
}
$breakdownString += "Total Records: $totalCount"
WriteToLog $breakdownString


$allUsersSage = @{}
$allUsersSageIdToUnique = @{}
$allUsersSageForManager = @{}

foreach($sageField in $SagefieldToADField.Values){
    if($sageField.Contains(".")){
        $records | Add-Member -MemberType NoteProperty -Name $sageField -Value ""
    }
}

foreach($record in $records){
    if($record.($uniqueValueADandSage[1])){
            if($record.fHCM2__Surname__c -ne $null -and $record.fHCM2__Surname__c -ne ""){
                $record.fHCM2__Surname__c = $record.fHCM2__Surname__c
            }
            if($record.Preferred_Name_or_First_Name__c -ne $null -and $record.Preferred_Name_or_First_Name__c -ne ""){
                $record.fHCM2__First_Name__c = $record.Preferred_Name_or_First_Name__c
            }

            foreach($sageField in $SagefieldToADField.Values){
                if($sageField.Contains(".")){
                    $splitValue = $sageField.Split(".")
                    $valueToSet = $record.$($splitValue[0])
                    for($i = 1; $i -lt $splitValue.Count; $i++){
                        $valueToSet = $valueToSet.($splitValue[$i])
                    }
                    $record.$sageField = $valueToSet
                }
            }

            # Extract the nested business entity field to clean up validation mapping rules
            $record | Add-Member -MemberType NoteProperty -Name "BusinessName" -Value $record.fHCM2__Current_Employment__r.fHCM2__Business_Name__c -Force

            $allUsersSage[$record.($uniqueValueADandSage[1])] = $record
            $allUsersSageIdToUnique[$record.id] = $record.($uniqueValueADandSage[1])
            $allUsersSageForManager[$record.($uniqueValueADandSage[1])] = $record
    }
}

if($testUser){
    WriteToLog "Sage users filtered to one user - $($testuser)"
    if($allUsersSage.ContainsKey($testUser)){
        $allUsersSage = @{$testuser=$allUsersSage[$testuser]}
    }
    else{
        WriteToLog "The test user either does not exist or is an inactive account in sage"
        WriteToLog "Script finished"
        return;
    }
}
else{
    WriteToLog "Sage users filtered to $($allUsersSage.count) users (removing users with an empty identifier)"
}

$allUsersSageIDToSAM = @{}
foreach($record in $records){
    if($record.($uniqueValueADandSage[1]) -eq $null){continue}
    $allUsersSageIDToSAM[$record.($uniqueValueADandSage[1])] = $record.id
}

$remainingUsersSage = $allUsersSage.Clone()

# --- Fetch or Mock AD Users ---
$allUsersAD = @()

if (!$IsWindows) {
    WriteToLog "Non-Windows OS detected. Generating a mock AD user record for visual preview..."
    
    $mockKey = if ($testUser) { $testUser } else { ($allUsersSage.Keys | Select-Object -First 1) }
    
    if ($mockKey -and $allUsersSage.ContainsKey($mockKey)) {
        $sampleSageUser = $allUsersSage[$mockKey]
        
        # Bypass conditional state criteria to trigger the mockup matrix output loop
        $sampleSageUser.fHCM2__Employment_Status__c = "Active"
        
        $mockADUser = [PSCustomObject]@{
            SamAccountName        = "preview.user"
            Name                  = "$($sampleSageUser.fHCM2__First_Name__c) $($sampleSageUser.fHCM2__Surname__c)"
            DistinguishedName     = "CN=Preview User,OU=Users,DC=yourdomain,DC=com"
            Enabled               = $true
            EmployeeID            = $mockKey
            Office                = "Old Office Location (AD)"
            surname               = "Old Surname (AD)"
            displayName           = "Old DisplayName (AD)"
            manager               = "CN=Old Manager,OU=Users,DC=yourdomain,DC=com"
            extensionAttribute2   = "False"
            extensionAttribute3   = "Old Division"
            cn                    = "Old CN"
        }
        $allUsersAD += $mockADUser
        WriteToLog "Mock AD User created successfully for Key ID: $mockKey (Forced Status: Active)"
    } else {
        WriteToLog "Warning: No Sage records available to create a mock preview user from."
    }
}
else {
    try {
        if($paths.count -eq 0){
            if($testUser){
                $allUsersAD = @(Get-ADUser -Filter "$($uniqueValueADandSage[0]) -eq '$testUser'" -Properties * -ErrorAction Stop)
            }
            else{
                $allUsersAD = Get-ADUser -Filter * -Properties * -ErrorAction Stop | where $uniqueValueADandSage[0] -ne $null
            }
        }
        else{
            if($testUser){
                foreach($path in $paths){
                    $allUsersAD += @(Get-ADUser -Filter "$($uniqueValueADandSage[0]) -eq '$testUser'" -Properties * -SearchBase $path -ErrorAction Stop)
                }
            }
            else{
                 foreach($path in $paths){
                    $allUsersAD += Get-ADUser -Filter * -SearchBase $path -Properties * -ErrorAction Stop | where $uniqueValueADandSage[0] -ne $null
                }
            }
        }
        WriteToLog ("A total of " + $allUsersAD.count + " records fetched from AD")
    }
    catch {
        $message = "Error when trying to fetch users from AD`n" + ($_ | Out-String)
        WriteToLog $message
        WriteToLog "Script finished"
        SendErrorMail $message
        break;
    }
}

$allUsersADHash = @{}
foreach($record in $allUsersAD){
    $record.EmployeeID = $record.EmployeeID -replace '^0+', ''
    if($record.($uniqueValueADandSage[0])){
        $allUsersADHash[($record.($uniqueValueADandSage[0])).Trim()] = $record
    }
}

$allUsersADClone = $allUsersAD
if($testUser){
    $allUsersAD = @($allUsersAD | where $uniqueValueADandSage[0] -eq $testuser)
    WriteToLog ("AD users filtered to one user - " + $testuser)
}

#Sage > AD
WriteToLog "###Updating users in AD###"
if ($allowUpdate) {
    foreach($adUser in $allUsersAD){
        if($aduser.($uniqueValueADandSage[0])){
            if($allUsersSage.ContainsKey($aduser.($uniqueValueADandSage[0]))){
                $sageUser = $allUsersSage.($aduser.($uniqueValueADandSage[0]))
                if(!($updateUsersWithTheseEmploymentStatuses.Contains($sageUser.fHCM2__Employment_Status__c))){
                    continue
                }
                $changedFields = @{}
                $newManager = $null
                $changedFieldsFriendly = @()

                $changedFieldsCN = $null
                $changedFieldsIsManager = $null
                $changedFieldsDivision = $null

                foreach($key in $SagefieldToADField.Keys){
                    if($key -eq "accountexpirationdate"){
                        if($adUser.accountexpirationdate -and $sageUser.($SagefieldToADField.accountexpirationdate)){
                            if([datetime]$adUser.accountexpirationdate -ne ([datetime]$sageUser.($SagefieldToADField.accountexpirationdate)).AddDays(1)){
                                $changedFields.Add("accountexpirationdate",[string]([datetime]$sageUser.($SagefieldToADField.accountexpirationdate)).AddDays(1).ToString("yyyy-MM-dd"))
                                $changedFieldsFriendly += [pscustomobject]@{
                                    "Property"=$key
                                    "Old Value"=([string]$aduser.($key)).Trim()
                                    "New Value"=([string]([datetime]$sageUser.($SagefieldToADField.$key)).AddDays(1)).Trim()
                                }
                            }
                        }
                        elseif($adUser.accountexpirationdate){
                            $changedFields.Add("accountexpirationdate",$null)
                            $changedFieldsFriendly += [pscustomobject]@{
                                "Property"=$key
                                "Old Value"=([string]$aduser.($key)).Trim()
                                "New Value"=([string]$sageUser.($SagefieldToADField.$key)).Trim()
                            }
                        }
                        elseif($sageUser.($SagefieldToADField.accountexpirationdate)){
                            $changedFields.Add("accountexpirationdate",[string]([datetime]$sageUser.($SagefieldToADField.accountexpirationdate)).AddDays(1).ToString("yyyy-MM-dd"))
                            $changedFieldsFriendly += [pscustomobject]@{
                                "Property"=$key
                                "Old Value"=([string]$aduser.($key)).Trim()
                                "New Value"=([string]([datetime]$sageUser.($SagefieldToADField.$key)).AddDays(1)).Trim()
                            }
                        }
                    }
                    else{
                        if(([string]$aduser.($key)).Trim() -ne ([string]$sageUser.($SagefieldToADField.$key)).Trim()){
                            $changedFields.Add($key,([string]$sageUser.($SagefieldToADField.$key)).Trim())
                            $changedFieldsFriendly += [pscustomobject]@{
                                "Property"=$key
                                "Old Value"=([string]$aduser.($key)).Trim()
                                "New Value"=([string]$sageUser.($SagefieldToADField.$key)).Trim()
                            }
                        }
                    }
                }

                if($aduser.extensionAttribute2 -ne $sageUser.fHCM2__Is_Manager__c){
                            $changedFieldsIsManager = $true
                            $changedFieldsIsManagerNew = $sageUser.fHCM2__Is_Manager__c
                            $changedFields.Add("extensionAttribute2",$changedFieldsIsManagerNew)
                            $changedFieldsFriendly += [pscustomobject]@{
                                "Property"="extensionAttribute2"
                                "Old Value"=$adUser.extensionAttribute2
                                "New Value"=$changedFieldsIsManagerNew
                            }
                        }
                if($aduser.extensionAttribute3 -ne $sageUser.fHCM2__Division__c){
                            $changedFieldsDivision = $true
                            $changedFieldsDivisionNew = $sageUser.fHCM2__Division__c
                            $changedFields.Add("extensionAttribute3",$changedFieldsDivisionNew)
                            $changedFieldsFriendly += [pscustomobject]@{
                                "Property"="extensionAttribute3"
                                "Old Value"=$adUser.extensionAttribute3
                                "New Value"=$changedFieldsDivisionNew
                            }
                        }

                if($SagefieldToADField.ContainsKey("GivenName")){
                    if($sageUser.fHCM2__First_Name__c -ne $null -and $sageUser.fHCM2__Surname__c -ne $null){
                        $fullname = $sageUser.fHCM2__First_Name__c + " " + $sageUser.fHCM2__Surname__c

                        if($aduser.displayName -ne $fullname){
                            $changedFields.Add("displayName",$fullname)
                            $changedFieldsFriendly += [pscustomobject]@{
                                "Property"="displayName"
                                "Old Value"=$adUser.displayName
                                "New Value"=$fullname
                            }
                        }
                        if($aduser.cn -ne $fullname){
                            $changedFieldsCN = $true
                            $changedFieldsFriendly += [pscustomobject]@{
                                "Property"="cn"
                                "Old Value"=$adUser.cn
                                "New Value"=$fullname
                            }
                        }
                    }
                }

                if(($changedFields.count -gt 0)){
                    $clear = @()
                    foreach($key in $changedFields.Keys){
                        if($changedFields.$key -eq ""){
                            $clear += $key
                        }
                    }
                    if($clear.Count -gt 0){
                        $changedFields.Add("Clear",@())
                        foreach($key in $clear){
                            $changedFields.Remove($key)
                            if($clearReplacementList.Contains($key)){
                                $changedFields.Clear += $clearReplacementList.$key
                            }
                            else{
                                $changedFields.Clear += $key
                            }
                        }
                    }
                    $modifiedUsers += $adUser.SamAccountName
                    if(!$diag){
                        try{
                            $changedFields.Remove('extensionAttribute2')
                            $changedFields.Remove('extensionAttribute3')

                            Set-ADUser $adUser @changedFields -ErrorAction Stop

                            if($changedFieldsIsManager -eq $true){
                                Set-ADUser $adUser -Replace @{ extensionAttribute2 = "$changedFieldsIsManagerNew" } -ErrorAction Stop
                                }
                            if($changedFieldsDivision -eq $true){
                                Set-ADUser $adUser -Replace @{ extensionAttribute3 = "$changedFieldsDivisionNew" } -ErrorAction Stop
                                }
                            if($changedFieldsCN -eq $true){
                                Rename-ADObject -Identity $adUser.DistinguishedName -NewName $fullname
                                }

                            $message =  "Updated AD user $($aduser.name)`n" + ($changedFieldsFriendly | Out-String)
                            WriteToLog $message
                        }
                        catch{
                            $message = "Error when trying to update AD user $($aduser.name)`n" + ($_ | Out-String)
                            WriteToLog $message
                        }
                    }
                    else{
                        $tableMarkup = $changedFieldsFriendly | Format-Table -Property Property, "Old Value", "New Value" | Out-String
                        $message = "Will update AD user: $($aduser.name) ($($adUser.SamAccountName))`n" + $tableMarkup
                        WriteToLog $message
                        
                        Write-Host "`n[PREVIEW] FIELD MISMATCH DETECTED FOR: $($aduser.name)" -ForegroundColor Cyan
                        $changedFieldsFriendly | Format-Table -Property Property, "Old Value", "New Value"
                    }
                }
                else{
                    WriteToLog ($adUser.name + " is in sync in AD")
                }
            }
       }
    }
} else {
    WriteToLog "Skipping account validation loops. (allowUpdate is configured to false)"
}

if($allowCreation){
    WriteToLog "###Creating users###"
    foreach($key in $allUsersSage.Keys){

        if($allUsersADHash.ContainsKey($key)){continue} # Skip if user already exists
        if($createUsersWithTheseEmploymentStatuses.Contains($allUsersSage.$key.fHCM2__Employment_Status__c)){
            $userprops = @{}
            $path = $newUserPath
            if($path){

                # Define the alphabet for the password
                $alphabet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*'
                #Generate a random 20-character string
                $randomPassword = -join ((1..20) | ForEach-Object { $alphabet[(Get-Random -Maximum $alphabet.Length)] })
                # Add it to your userprops hash

                $sageUser = $allUsersSage[$key]
                $newManager = $null
                $firstName = $allUsersSage.$key.fHCM2__First_Name__c
                $surname = $allUsersSage.$key.fHCM2__Surname__c
                $userprops.Add("givenname",$firstname)
                $userprops.Add("surname",$surname)
                $userprops.Add("displayName",$firstName + " " + $surname)
                $userprops.Add("Path",$path)
                $userprops.Add("accountPassword", (ConvertTo-SecureString -AsPlainText $randomPassword -Force))           

                foreach($prop in $SagefieldToADField.Keys){
                    if($userprops.ContainsKey($prop)){
                        continue
                    }
                    $userprops.Add($prop,([string]$allUsersSage.$key.($SagefieldToADField.$prop)).Trim())
                }

                if(!$userprops.accountexpirationdate){
                    $userprops.Remove("accountexpirationdate")
                }

                $userprops.Add($uniqueValueADandSage[0],$allUsersSage.$key.($uniqueValueADandSage[1]))

                $name = $firstName + " " + $userprops.Surname
                $userprops.Add("Enabled",$false)

                $dn = "CN=" + $firstName + " " + $userprops.Surname + "," + $userprops.Path
                if(Get-ADUser -Filter {DistinguishedName -eq $dn}){
                    $index = 1
                    while(1){
                        $dn = "CN=" + $firstName + " " + $userprops.Surname + $index + ", " + $userprops.Path
                        if(!(Get-ADUser -Filter {DistinguishedName -eq $dn})){
                            $name = $firstName + " " + $userprops.Surname + $index
                            break
                        }
                        $index++
                    } 
                } 
                $userprops.Add("name",$name)
                
                $sam = GetAvailablesAMAccountName $firstName $userprops.Surname
                if($sam -eq $null){
                    WriteToLog "Skipping user creation of $($name)"
                    continue
                }
                
                # Dynamic mapping evaluation logic for multi-suffix email domains
                $sageBusinessName = $allUsersSage.$key.BusinessName
                $assignedDomain = $domainMappingHash["Default"]
                if ($sageBusinessName -and $domainMappingHash.ContainsKey($sageBusinessName)) {
                    $assignedDomain = $domainMappingHash[$sageBusinessName]
                }

                $UPN = GetAvailableUPN $firstName $userprops.surname $assignedDomain
                if($UPN -eq $null){
                    WriteToLog "Skipping user creation of $($name)"
                    continue
                }
                $userprops.Add("samaccountname",$sam)
                $userprops.Add("userprincipalname",$UPN)
                $userprops.Add("EmailAddress",$UPN)

                if(!$diag){
                    try{
                        New-ADUser @userprops -ErrorAction Stop
                        $message = "Successfully created AD user $($userprops.name)`n" + ($userprops | Out-String)
                        $usersCreated += $userProps
                        WriteToLog $message
                    }
                    catch{
                        $message = "Error when trying to create AD user $($userprops.name)`n" + ($_ | Out-String)
                        WriteToLog $message
                    }
                }
                else{
                    $usersCreated += $userProps
                    WriteToLog ("Will create user $($userprops.name)`n" + ($userprops | out-string))
                }
            }
        }
    }
    WriteToLog "###Creation of new users done###"
} else {
    WriteToLog "Skipping user onboarding loops. (allowCreation is configured to false)"
}

if($disableUsersWhenLeft){
    WriteToLog "###Disabling accounts###"
    if ($allowDeactivation) {
        foreach($key in $allUsersSage.Keys){
            # ADD THIS LINE HERE:
        Write-Host "DEBUG: Processing Key: $key" -ForegroundColor Gray

            if($allUsersADHash[$key]){
                if($disableUsersWithTheseEmploymentStatuses.Contains($allUsersSage[$key].fHCM2__Employment_Status__c)){
                    $adUser = $allUsersADHash[$key]
                    if($adUser.enabled){
                        $disabledUsers += $adUser.SamAccountName
                        if(!$diag){
                            try{
                                $adUser | Disable-ADAccount -Confirm:$false -ErrorAction Stop
                                WriteToLog "Successfully disabled AD user $($adUser.name)"                      
                            }
                            catch{
                                $message = "Error when trying to disable AD user $($adUser.name)`n" + ($_ | Out-String)
                                WriteToLog $message
                            }
                        }
                        else{
                            WriteToLog "Will disable AD user $($adUser.name)"
                        }
                    }
                }
            }
        }
    } else {
        WriteToLog "Skipping account deactivation filters. (allowDeactivation is configured to false)"
    }
   WriteToLog "###Disable accounts done###"
}

if($SyncManagerFromSage){
    WriteToLog "###Updating managers in AD###"
    if ($allowUpdate) {
        foreach($key in $allUsersSage.Keys){
            if($allUsersADHash.$key){
                if(!($updateUsersWithTheseEmploymentStatuses.Contains($allUsersSage.$key.fHCM2__Employment_Status__c))){
                    continue
                }
                $newManager = $false
                $newManagerValue = $null
                if($allUsersSage.$key.fHCM2__Manager__r){
                    if($allUsersSage.$key.fHCM2__Manager__r.fHCM2__Unique_Id__c -eq $null){
                        continue
                    }
                    if($allUsersSageIDToSAM.ContainsKey($allUsersSage.$key.fHCM2__Manager__r.fHCM2__Unique_Id__c)){
                        $managerUniqueId = $allUsersSage.$key.fHCM2__Manager__r.fHCM2__Unique_Id__c
                        $manager = Get-ADUser -Filter "$($uniqueValueADandSage[0]) -eq '$managerUniqueId'" -Properties *
                        if($manager){
                            if($allUsersADHash.$key.manager -ne $manager.DistinguishedName){
                                $newManager = $true
                                $newManagerValue = $manager
                            }
                            else{
                                WriteToLog ("Manager is in sync for " + $allUsersADHash.$key.name)
                            }
                        }
                        else{
                            WriteToLog ("The manager for " + $($allUsersADHash.$key.name) + " (" + $managerUniqueId + ") doesn't exist in AD")
                        }
                    }
                    else{
                        WriteToLog ("The manager for " + $($allUsersADHash.$key.name) + " (" + $($allUsersSage.$key.fHCM2__Manager__r.fHCM2__Unique_Id__c) + ") either doesn't exist, have no unique identifier in Sage")
                    }
                }
                else{
                    if($allUsersADHash.$key.manager -ne $null){
                        $newManager = $true
                    }
                    else{
                        WriteToLog ("Manager is in sync for " + $($allUsersADHash.$key.name))
                    }
                }
                if($newManager){
                    if(!($modifiedUsers.contains($allUsersADHash.$key.SamAccountName))){
                        $modifiedUsers += $allUsersADHash.$key.SamAccountName
                    }
                    if($diag){
                        WriteToLog ("Will set " + $($newManagerValue.name) + " as the manager for " + $($allUsersADHash.$key.name))
                    }
                    else{
                        try{
                            $allUsersADHash.$key | Set-ADUser -Manager $newManagerValue -ErrorAction Stop
                            WriteToLog ("Succesfully set " + $($newManagerValue.name) + " as the manager for " + $($allUsersADHash.$key.name))
                        }
                        catch{
                            WriteToLog ("Error when trying to set " + $($newManagerValue.name) + " as the manager for " + $($allUsersADHash.$key.name) + $_)
                        }
                    }
                }
            }
        }
    } else {
        WriteToLog "Skipping manager mapping updates. (allowUpdate is configured to false)"
    }
    WriteToLog "###Manager update done###"
}

Function UpdateWorkEmailInSage($userId, $emailValue, $fedIdValue){
    $url = "/services/data/v58.0/sobjects/fHCM2__Team_Member__c/$userId"
    $body = @{
        "Work_Email_Address_Holding_Field__c" = $emailValue
        "Federation_ID_Holding_Field__c"      = $fedIdValue
    } | ConvertTo-Json

    try {
        $headers = @{
            Authorization  = "$($session.token_type) $($session.access_token)";
            "Content-Type" = "application/json"
        }
        Invoke-RestMethod ($salesforceUrl + $url) -Headers $headers -Body $body -Method Patch -ErrorAction Stop
        WriteToLog "Updated Sage ID $($userId): Email to '$emailValue', FedID to '$fedIdValue'"
    }
    catch {
        WriteToLog "Error updating Sage for ID $($userId): $($_.Exception.Message)"
    }
}

if($SyncEmailFromADToSage){
    WriteToLog "###Syncing Email to Sage###"
    if ($allowUpdate) {
        foreach($key in $allUsersSage.Keys){
            if($allUsersADHash.ContainsKey($key)){
                $adUser = $allUsersADHash[$key]
                $sageUser = $allUsersSage[$key]

                $adMail = $adUser.mail
                $adUPN  = $adUser.UserPrincipalName
                
                if ($adUser.mail -ne $null -and $adUser.mail -ne "") {
                    if (($sageUser.Work_Email_Address_Holding_Field__c -ne $adMail) -or ($sageUser.Federation_ID_Holding_Field__c -ne $adUPN)) {
                        WriteToLog "Updating Sage record for user $($adUser.name): Email to $($adMail), FedID to $($adUPN)"
                        if (!$diag) {
                            UpdateWorkEmailInSage $sageUser.Id $adMail $adUPN
                        }
                    }
                }
            }
        }
    } else {
        WriteToLog "Skipping reverse email synchronization to Sage. (allowUpdate is configured to false)"
    }
    WriteToLog "####Syncing Email to Sage Done###"
}

WriteToLog "###Statistics###"
WriteToLog ("Users Created: " + $usersCreated.Count)
WriteToLog ("Users Modified: " + $modifiedUsers.Count)
WriteToLog ("Users Disabled: " + $disabledUsers.Count)

SendErrorMail
WriteToLog "Script finished"
Stop-Transcript