# Version 1.0

# functions
function Initialize-ColorScheme
{
    Set-Variable -Name "successColor" -Value "Green" -Scope "Script" -Option "Constant"
    Set-Variable -Name "infoColor" -Value "DarkCyan" -Scope "Script" -Option "Constant"
    Set-Variable -Name "warningColor" -Value "Yellow" -Scope "Script" -Option "Constant"
    Set-Variable -Name "failColor" -Value "Red" -Scope "Script" -Option "Constant"
}

function Show-Introduction
{
    Write-Host "This script downloads a list of emails via their internet message ID..." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor $infoColor
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "^\s*y\s*$") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
            "1. Open Powershell as admin.`n" +
            "2. CD into script directory.`n" +
            "3. Run .\scriptname`n") -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        exit
    }
}

function PromptFor-ClientCredentials
{
    $clientId = (Read-Host "Enter your client ID").Trim()
    Write-Host "Enter your client secret in the prompt in the password field:"
    return Get-Credential -Credential $clientId
}

function TryConnect-MgGraph($tenantId, [PSCredential]$clientCredentials)
{    
    Write-Host "Connecting to Microsoft Graph in app context..." -ForegroundColor $infoColor

    try
    {
        Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $clientCredentials -ErrorAction "Stop" | Out-Null
    }
    catch
    {        
        $errorRecord = $_
        Write-Host "An error occurred when connecting to Microsoft Graph: `n$errorRecord" -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        exit
    }

    Write-Host "Successfully connected!" -ForegroundColor $successColor
}

function Test-ConnectedToMgGraph
{
    return $null -ne (Get-MgContext)
}

function PromptFor-EmailList
{
    Write-Host "Script requires CSV list of internet message IDs. Must include headers named `"Mailbox Address`", and `"Internet Message ID`"." -ForegroundColor $infoColor
    $csvPath = (Read-Host "Enter path to CSV (must be .csv)").Trim('"')
    return Import-Csv -Path $csvPath
}

function Confirm-CSVHasCorrectHeaders($importedCSV)
{
    $firstRecord = $importedCSV | Select-Object -First 1
    $validCSV = $true

    if ($null -eq $firstRecord)
    {
        Write-Host "CSV is empty" -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        Exit
    }

    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name "Internet Message ID"))
    {
        Write-Host "This CSV file is missing a header called `"Internet Message ID`"." -ForegroundColor $failColor
        $validCSV = $false
    }

    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name "Mailbox Address"))
    {
        Write-Host "This CSV file is missing a header called `"Mailbox Address`"." -ForegroundColor $failColor
        $validCSV = $false
    }

    if (-not($validCSV))
    {
        Write-Host "Please make corrections to the CSV." -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        Exit
    }
}

function Download-AllEmails($emailList, $exportDirectory)
{
    $i = 1
    foreach ($email in $emailList)
    {
        Write-Progress -Activity "Downloading emails..." -Status $i
        if ( ($null -eq $email.'Mailbox Address') -or ($null -eq $email.'Internet Message ID')) { continue }
        $emailContent = Get-Email -MailboxAddress $email.'Mailbox Address' -InternetMessageId $email.'Internet Message ID'
        Download-Email -MailboxAddress $email.'Mailbox Address' -MessageId $emailContent.id -InternetMessageId $email.'Internet Message ID' -Directory $exportDirectory
        $i++
    }
}

function Get-Email($mailboxAddress, $internetMessageId)
{
    if ($internetMessageId -inotmatch "^<.*>$")
    {
        $internetMessageId = "<$internetMessageId>"
    }

    # https://learn.microsoft.com/en-us/graph/api/user-list-messages
    $uri = "https://graph.microsoft.com/v1.0/users/$mailboxAddress/messages?`$filter=internetMessageId eq '$internetMessageId'"
    try
    {
        $email = Invoke-MgGraphRequest -Method "Get" -Uri $uri
    }
    catch
    {
        $errorRecord = $_
        Log-Warning "An error occurred when getting email. `nMailbox address: $mailboxAddress `nInternet Message Id: $internetMessageId `n$errorRecord"
    }
    return $email.value
}

function Log-Warning($message, $logPath = ".\logs.txt")
{
    $message = "[$(Get-Date -Format 'yyyy-MM-dd hh:mm tt') W] $message"
    Write-Output $message | Tee-Object -FilePath $logPath -Append | Write-Host -ForegroundColor $warningColor
}

function Download-Email($mailboxAddress, $messageId, $internetMessageId, $directory)
{
    # https://learn.microsoft.com/en-us/graph/outlook-get-mime-message
    $uri = "https://graph.microsoft.com/v1.0/users/$mailboxAddress/messages/$messageId/`$value"
    $internetMessageIdModified = $internetMessageId.Replace('<', '').Replace('>', '')
    $filePath = "$directory\$($mailboxAddress)_$internetMessageIdModified.eml"
    try
    {
        Invoke-MgGraphRequest -Method "Get" -Uri $uri -OutputFilePath $filePath
    }    
    catch
    {
        $errorRecord = $_
        Log-Warning "An error occurred when downloading email. Mailbox address: $mailboxAddress MessageId: $internetMessageId `n$errorRecord"
    }
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module "Microsoft.Graph.Authentication"
$tenantId = (Read-Host "Enter your tenant ID").Trim()
$clientCredentials = PromptFor-ClientCredentials
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor $infoColor
Disconnect-MgGraph -ErrorAction "SilentlyContinue" | Out-Null
TryConnect-MgGraph -TenantId $tenantId -ClientCredentials $clientCredentials
$emailList = PromptFor-EmailList
Confirm-CSVHasCorrectHeaders $emailList
$timeStamp = Get-Date -Format 'yyyy-MM-dd-hh-mmtt'
$exportDirectory = New-Item -Path "$PSScriptRoot\Email Export $timeStamp" -ItemType "Directory"
Download-AllEmails -EmailList $emailList -ExportDirectory $exportDirectory.FullName
Write-Host "All done! Emails exported to $exportDirectory" -ForegroundColor $successColor
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor $infoColor
Disconnect-MgGraph | Out-Null
Read-Host "Press Enter to exit"

# Should script be updated to provide option to forward them to a specific mailbox instead of (or in addition to) downloading them?
# Should script be updated to have file names with email subjects?

<#
Useful articles:
https://learn.microsoft.com/en-us/graph/outlook-get-mime-message
https://learn.microsoft.com/en-us/graph/api/message-get
https://learn.microsoft.com/en-us/graph/api/user-list-messages
#>

<#
Testing to do / edge cases to smooth out:
- Null inputs to functions
- User enters blank input on prompt
- Tenant ID invalid
- Client ID not found
- Client secret not found
- Client secret expired
- Service principal missing the required privileges
- Will it hit the rate limits of 1500 requests per 20 seconds?

Testing done:
- Trying to export email to directory but file with that name already exists
    - Outcome: The file is overwritten
- Was already connected to Microsoft Graph when you started the script
    - Outcome: Session is replaced for duration of script and then disconnected
- Input file has 0 emails
    - Outcome: Error is thrown that CSV is empty, and script exits.
- Input file has 1 email
    - Outcome: The email is downlaoded as expected
- Input file has 2+ emails
    - Outcome: All emails are downloaded as expected
- User inputs with leading/trailing spaces
    - Outcome: Spaces are trimmed
- Input file is missing the necessary headers
    - Outcome: Error is thrown and script exits
- Input file has internet message ID missing angle brackets <>
    - Outcome: Angle brackets are added automatically and email is downloaded as expected
- Input file has mailbox address not found
    - Outcome: 404 not found error is logged and other emails continue to download fine
- Input file has internet message ID not found
    - Outcome: Errors are logged for 404 not found or 400 bad request and other emails to continue to download fine
#>