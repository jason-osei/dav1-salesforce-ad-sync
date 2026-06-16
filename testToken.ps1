[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor
[Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

Import-Module ActiveDirectory

# --- Salesforce Credentials ---
$client     = "3MVG9WktJx3rjNu2vZrWMhn5uva5Rtmdrh94JcWuZzd_vIzxfaduYN32UZ62tf8S4YYoVYE4BEfBTc2ueqaL2"
$secret     = "C9D43889FA069F1C75EFCC7D3CE66444BAD7A2FCEF57939E69A432EBAEC89F90"
$sfLoginUrl = "https://davisonslaw.my.salesforce.com/services/oauth2/token"

# ============================================================
# STEP 1: Authenticate with Salesforce
# ============================================================
Write-Host "Authenticating with Salesforce..." -ForegroundColor Cyan

try {
    $session = Invoke-RestMethod -Uri $sfLoginUrl `
        -Method Post `
        -Headers @{ "Content-Type" = "application/x-www-form-urlencoded" } `
        -Body "grant_type=client_credentials&client_id=$client&client_secret=$secret"

    Write-Host "Authenticated. Instance: $($session.instance_url)" -ForegroundColor Green

    Write-Host ("Session summary:`n" + (@{
        instance_url = $session.instance_url
        token_type   = $session.token_type
        issued_at    = $session.issued_at
        signature    = $session.signature
    } | ConvertTo-Json -Depth 5))
}
catch {
    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
    Write-Host "Auth failed: $($reader.ReadToEnd())" -ForegroundColor Red
    exit
}