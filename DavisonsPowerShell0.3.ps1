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
#############################################################################################################################

# Parameters
param(

    [string]$salesforceUrl = "https://davisonslaw.my.salesforce.com",
    [string]$client = "3MVG9WktJx3rjNu2vZrWMhn5uva5Rtmdrh94JcWuZzd_vIzxfaduYN32UZ62tf8S4YYoVYE4BEfBTc2ueqaL2",
    [string]$secret = "C9D43889FA069F1C75EFCC7D3CE66444BAD7A2FCEF57939E69A432EBAEC89F90",

    [string]$transcriptPath = $(if ($IsWindows) { "C:\scripts\HRSync\logs\transcript" } else { "$HOME/Downloads/HRSync/logs/transcript" }),
    [string]$logPath = $(if ($IsWindows) { "C:\scripts\HRSync\logs" } else { "$HOME/Downloads/HRSync/logs" }),
    
    [bool]$diag = $false,
    [array]$paths = @("OU=Users,OU=Davisons,DC=davisons,DC=law","OU=Accounts,OU=Data,DC=davisons,DC=law"),
    [string]$newUserPath = "OU=_NewJoiner,OU=Users,OU=Davisons,DC=davisons,DC=law",
    
    [bool]$allowCreation = $true,
    [bool]$allowUpdate = $false,
    [bool]$allowDeactivation = $false,
    
    [array]$createUsersWithTheseEmploymentStatuses = @(
        "Pre Joiner"
    ),
    [array]$updateUsersWithTheseEmploymentStatuses = @(
        "Active",
        "Active Employee",
        "Pre-Joiner"
    ),
    [bool]$disableUsersWhenLeft = $true,
    [array]$disableUsersWithTheseEmploymentStatuses = @(
        "Leaver",
        "Left"
    ),

    [array]$uniqueValueADandSage = @("employeeId","fHCM2__Unique_Id__c"),
    [string]$testUser = $null,
    [array]$SagefieldToADField = @{
        "Office"          = "fHCM2__Current_Employment__r.fHCM2__Work_Location__r.Name"
        "surname"         = "fHCM2__Surname__c"
        "company"         = "fHCM2__Current_Employment__r.fHCM2__Business_Name__c"
        "telephoneNumber" = "fHCM2__Phone_Number__c"
    },

    [bool]$SyncEmailFromADToSage = $false,
    $syncManagerFromSage = $true,

    $shouldMailOnError = $true,
    $smtpFrom          = "hris-sync@davisons.law",
    $smtpTo            = "support@davisons.law",
    $smtpServer        = "smtp.davisons.law",
    $smtpPort          = "587",
    $smtpUseSSL        = $true
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor
[Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

$sfLoginUrl = "$salesforceUrl/services/oauth2/token"
$workLocationAPIField = "fHCM2__Current_Employment__r.fHCM2__Work_Location__r.Name"

$domainMappingHash = @{
    "Davisons Law" = "@davisons.law"
    "DP Law"       = "@dplaw.law"
    "Default"      = "@davisons.law"
}

$allowedEmailDomains = @(
    "davisons.law",
    "dplaw.law"
)

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
    
    Write-Host $message
    $message | Out-File $logName -Force -Append
}

Function GetEmailDomainFromSageValue($emailValue) {
    if ([string]::IsNullOrWhiteSpace($emailValue)) {
        return $null
    }

    $normalizedEmail = $emailValue.ToLower().Trim()

    if ($normalizedEmail -match '@') {
        $candidateDomain = ($normalizedEmail.Split('@')[-1]).Trim()
    }
    elseif ($normalizedEmail -match '(davisons\.law|dplaw\.law)$') {
        $candidateDomain = $matches[1]
    }
    else {
        $candidateDomain = $null
    }

    if ($candidateDomain -and ($allowedEmailDomains -contains $candidateDomain)) {
        return "@" + $candidateDomain
    }

    return $null
}

Function GetSafeOUName($name) {
    if ([string]::IsNullOrWhiteSpace($name)) {
        return "Davisons Law"
    }

    $safeName = $name.Trim()
    $safeName = $safeName -replace '[\\\/\+\=\<\>\#\;\,\"]', '-'
    $safeName = $safeName -replace '\s+', ' '
    $safeName = $safeName.Trim()

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        return "Davisons Law"
    }

    return $safeName
}

Function EnsureOUExists($ouName, $parentPath) {
    $safeOUName = GetSafeOUName $ouName
    $ouDn = "OU=$safeOUName,$parentPath"

    try {
        $existingOU = Get-ADOrganizationalUnit -Identity $ouDn -ErrorAction Stop
        WriteToLog "OU already exists: $ouDn"
        return $ouDn
    }
    catch {
        if($diag){
            WriteToLog "Would create OU: $ouDn"
            return $ouDn
        }

        try {
            New-ADOrganizationalUnit -Name $safeOUName -Path $parentPath -ProtectedFromAccidentalDeletion $false -ErrorAction Stop | Out-Null
            WriteToLog "Created OU: $ouDn"
            return $ouDn
        }
        catch {
            WriteToLog "Failed to create OU: $ouDn. Error: $($_.Exception.Message)"
            return $null
        }
    }
}

Import-Module ActiveDirectory

if ($psise) {
    $scriptRoot = Split-Path $psise.CurrentFile.FullPath
}
else {
    $scriptRoot = $global:PSScriptRoot
}

$clearReplacementList = @{
    "state"   = "st"
    "city"    = "l"
    "Country" = "C"
    "Office"  = "physicalDeliveryOfficeName"
}

[array]$allSageValues = @()
if($SagefieldToADField.Keys.count -gt 0){
    $allSageValues += $SagefieldToADField.values
}
$allSageValues += $uniqueValueADandSage[1]

$queryValues = @(
    "id"
    "fHCM2__Hire_Date__c"
    "fHCM2__First_Name__c"
    "fHCM2__Middle_Name__c"
    "fHCM2__Surname__c"
    "fHCM2__Preferred_Name__c"
    "fHCM2__Email__c"
    "fHCM2__Phone_Number__c"
    "fHCM2__Unique_Id__c"
    "fHCM2__Job_Title__c"
    "fHCM2__Division__c"
    "fHCM2__Employment_Status__c"
    "fHCM2__Manager__r.fHCM2__Unique_Id__c"
    "fHCM2__Current_Employment__r.Name"
    "fHCM2__Country__c"
    "fHCM2__Current_Employment__r.fHCM2__Start_Date__c"
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
    New-Item -ItemType Directory -Path $transcriptPath | Out-Null
}

if(!(Test-Path $logPath)){
    New-Item -ItemType Directory -Path $logPath | Out-Null
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
    
    $firstInitial = $givenName[0]
    $name = ($firstInitial + $surname)[0..19] -join ""
    $counter = 1
    
    while($counter -le 20){
        try{
            if(!(Get-ADUser -Filter {sAMAccountName -eq $name} -ErrorAction Stop)){
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

    $firstInitial = $givenName[0]
    $name = ($firstInitial + $surname) + $targetDomain

    $counter = 1
    while($counter -le 20){
        try{
            if(!(Get-ADUser -Filter {userprincipalname -eq $name} -ErrorAction Stop)){
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
    $adUser | Get-Member -MemberType property | Select-Object -ExpandProperty name | ForEach-Object {
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
    $session = Invoke-RestMethod -Uri $sfLoginUrl `
        -Method Post `
        -Headers @{ "Content-Type" = "application/x-www-form-urlencoded" } `
        -Body "grant_type=client_credentials&client_id=$client&client_secret=$secret"

    WriteToLog "Authenticated. Instance: $($session.instance_url)"

    WriteToLog ("Session summary:`n" + (@{
        instance_url = $session.instance_url
        token_type   = $session.token_type
        issued_at    = $session.issued_at
        signature    = $session.signature
    } | ConvertTo-Json -Depth 5))
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

$statusCounts = @{}
$totalCount = 0
foreach ($record in $records) {
    $statusValue = $record.fHCM2__Employment_Status__c
    if ([string]::IsNullOrEmpty($statusValue)) { $statusValue = "[Blank/Null Status]" }
    $statusCounts[$statusValue]++
    $totalCount++
}

$breakdownString = "Employment Status Breakdown:`n"
foreach ($statusKey in $statusCounts.Keys | Sort-Object) {
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
        return
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

$allUsersAD = @()
try {
    if($paths.count -eq 0){
        if($testUser){
            $allUsersAD = @(Get-ADUser -Filter "$($uniqueValueADandSage[0]) -eq '$testUser'" -Properties * -ErrorAction Stop)
        }
        else{
            $allUsersAD = Get-ADUser -Filter * -Properties * -ErrorAction Stop | Where-Object { $_.($uniqueValueADandSage[0]) -ne $null }
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
                $allUsersAD += Get-ADUser -Filter * -SearchBase $path -Properties * -ErrorAction Stop | Where-Object { $_.($uniqueValueADandSage[0]) -ne $null }
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
    break
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
    $allUsersAD = @($allUsersAD | Where-Object { $_.($uniqueValueADandSage[0]) -eq $testuser })
    WriteToLog ("AD users filtered to one user - " + $testuser)
}

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
        if($allUsersADHash.ContainsKey($key)){ continue }

        if($createUsersWithTheseEmploymentStatuses.Contains($allUsersSage.$key.fHCM2__Employment_Status__c)){
            $userprops = @{}

            $sageUser = $allUsersSage[$key]
            $businessName = $sageUser.BusinessName
            $targetOuPath = EnsureOUExists $businessName $newUserPath

            if(-not $targetOuPath){
                WriteToLog "Skipping user creation because OU could not be resolved for business '$businessName'."
                continue
            }

            $path = $targetOuPath
            if($path){

                $alphabet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*'
                $randomPassword = -join ((1..20) | ForEach-Object { $alphabet[(Get-Random -Maximum $alphabet.Length)] })
                $securePassword = ConvertTo-SecureString -AsPlainText $randomPassword -Force

                $newManager = $null
                $firstName = $allUsersSage.$key.fHCM2__First_Name__c
                $surname = $allUsersSage.$key.fHCM2__Surname__c
                $userprops.Add("givenname",$firstname)
                $userprops.Add("surname",$surname)
                $userprops.Add("displayName",$firstName + " " + $surname)
                $userprops.Add("Path",$path)

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

                $sageEmail = $allUsersSage[$key].fHCM2__Email__c
                $assignedDomain = GetEmailDomainFromSageValue $sageEmail

                if (-not $assignedDomain) {
                    $assignedDomain = $domainMappingHash["Default"]
                    WriteToLog "Email domain not recognised for $($firstName) $($surname): '$sageEmail'. Falling back to default domain '$assignedDomain'."
                }
                else {
                    WriteToLog "Using email-derived domain '$assignedDomain' for $($firstName) $($surname) from Sage email '$sageEmail'."
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

                        Set-ADAccountPassword -Identity $sam `
                            -NewPassword $securePassword `
                            -Reset `
                            -ErrorAction Stop

                        Enable-ADAccount -Identity $sam -ErrorAction Stop

                        $userpropsForLog = $userprops.Clone()
                        $userpropsForLog["GeneratedPassword"] = $randomPassword

                        $message = "Successfully created AD user $($userprops.name)`n" + ($userpropsForLog | Out-String)
                        $usersCreated += $userProps
                        WriteToLog $message
                    }
                    catch{
                        $message = "Error when trying to create AD user $($userprops.name)`n" + ($_ | Out-String)
                        WriteToLog $message
                    }
                }
                else{
                    $userpropsForLog = $userprops.Clone()
                    $userpropsForLog["GeneratedPassword"] = $randomPassword
                    $usersCreated += $userProps
                    WriteToLog ("Will create user $($userprops.name)`n" + ($userpropsForLog | Out-String))
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
                    if(!($modifiedUsers.Contains($allUsersADHash.$key.SamAccountName))){
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
            Authorization  = "$($session.token_type) $($session.access_token)"
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
