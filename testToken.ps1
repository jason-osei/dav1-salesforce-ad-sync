[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor
[Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

Set-Location $PSScriptRoot
WriteToLog "Script started"

try {
    WriteToLog "Authenticating with Salesforce..."

    $sfLoginUrl = $salesforceUrl.TrimEnd('/') + "/services/oauth2/token"

    $session = Invoke-RestMethod -Uri $sfLoginUrl `
        -Method Post `
        -Headers @{ "Content-Type" = "application/x-www-form-urlencoded" } `
        -Body "grant_type=client_credentials&client_id=$client&client_secret=$secret" `
        -ErrorAction Stop

    WriteToLog "Authenticated. Instance: $($session.instance_url)"

    WriteToLog ("Session summary:`n" + (@{
        instance_url = $session.instance_url
        token_type   = $session.token_type
        issued_at    = $session.issued_at
        signature    = $session.signature
    } | ConvertTo-Json -Depth 5))
}
catch {
    $rawError = $_ | Out-String
    $responseText = $null
    $messageDetails = $null

    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
        $responseText = $_.ErrorDetails.Message
    }
    elseif ($_.Exception -and $_.Exception.Response) {
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $responseText = $reader.ReadToEnd()
        }
        catch {
            $responseText = $null
        }
    }

    if ($responseText) {
        try {
            $messageDetails = $responseText | ConvertFrom-Json
        }
        catch {
            $messageDetails = $responseText
        }
    }

    if ($messageDetails -and $messageDetails.error -eq "invalid_grant") {
        WriteToLog "Sages Credentials are invalid, please check username and password"
        WriteToLog ("Auth failed response:`n" + ($messageDetails | Out-String))
    }
    elseif ($responseText) {
        WriteToLog ("Auth failed:`n" + $responseText)
    }
    else {
        WriteToLog "Auth failed with no response body."
        WriteToLog $rawError
    }

    WriteToLog "Script finished"
    break
}
