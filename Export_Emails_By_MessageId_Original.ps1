# Install-Module MSAL.PS -Scope CurrentUser
# Install-Module Microsoft.Graph -Scope CurrentUser

param(
    [Parameter(Mandatory)]
    [string]$TenantId,
 
    [Parameter(Mandatory)]
    [string]$ClientId,
 
    [Parameter(Mandatory)]
    [string]$ClientSecret,
 
    [Parameter(Mandatory)]
    [string]$SourceMailbox,
 
    [Parameter(Mandatory)]
    [string]$MessageId,
 
    [string]$ForwardToMailbox
) 
 
# Ensure message ID is properly wrapped 
$ClientSecret = $ClientSecret | ConvertTo-SecureString -AsPlainText -Force
if ($MessageId -notmatch "^<.*>$") {
    $MessageId = "<$MessageId>"
}
 
# Authenticate with Graph API
$TokenResponse = Get-MsalToken -ClientId $ClientId -ClientSecret $ClientSecret -TenantId $TenantId -Scopes "https://graph.microsoft.com/.default"
$AccessToken = $TokenResponse.AccessToken
 
# Search for the message by internetMessageId
$encodedMessageId = [System.Net.WebUtility]::UrlEncode($MessageId)
$searchUrl = "https://graph.microsoft.com/v1.0/users/$SourceMailbox/messages?`$filter=internetMessageId eq '$encodedMessageId'"
 
$response = Invoke-RestMethod -Method Get -Uri $searchUrl -Headers @{ Authorization = "Bearer $AccessToken" }
 
if ($response.value.Count -eq 0) {
    Write-Host "❌ No message found with that Message-ID." -ForegroundColor Red
    return
}
 
$message = $response.value[0]
Write-Host "✅ Found message: $($message.subject)" -ForegroundColor Green
 
# Download MIME content (EML)
#$downloadUrl = "https://graph.microsoft.com/v1.0/users/$SourceMailbox/messages/$($message.id)/$value"
#$emlFile = "$($message.id).eml"
#Invoke-RestMethod -Uri $downloadUrl -Method Get -Headers @{ Authorization = "Bearer $AccessToken" } -OutFile $emlFile
 
#Write-Host "📥 Email downloaded to: $emlFile"
 
# Optional: Forward it
if ($ForwardToMailbox) {
    Write-Host "✉️ Forwarding email to $ForwardToMailbox..."
 
    $forwardUrl = "https://graph.microsoft.com/v1.0/users/$SourceMailbox/messages/$($message.id)/forward"
 
    $body = @{
        Comment = $MessageId
        ToRecipients = @(@{
            EmailAddress = @{
                Address = $ForwardToMailbox
            }
        })
    } | ConvertTo-Json -Depth 10
 
    Invoke-RestMethod -Uri $forwardUrl -Method POST -Headers @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" } -Body $body
 
    Write-Host "✅ Email forwarded to $ForwardToMailbox"
}