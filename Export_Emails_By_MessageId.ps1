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

function TryConnect-MgGraph($scopes)
{
    $connected = Test-ConnectedToMgGraph
    while (-not($connected))
    {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor $infoColor

        if ($null -ne $scopes)
        {
            Connect-MgGraph -Scopes $scopes -ErrorAction SilentlyContinue | Out-Null
        }
        else
        {
            Connect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }

        $connected = Test-ConnectedToMgGraph
        if (-not($connected))
        {
            Read-Host "Failed to connect to Microsoft Graph. Press Enter to try again"
        }
        else
        {
            Write-Host "Successfully connected!" -ForegroundColor $successColor
        }
    }    
}

function Test-ConnectedToMgGraph
{
    return $null -ne (Get-MgContext)
}

function PromptFor-EmailList
{
    Write-Host "Script requires CSV list of internet message IDs. Must include headers named `"Mailbox Address`", and `"Internet Message ID`"." -ForegroundColor $infoColor
    $csvPath = Read-Host "Enter path to CSV (must be .csv)"
    $csvPath = $csvPath.Trim('"')
    return Import-Csv -Path $csvPath
}

function Confirm-CSVHasCorrectHeaders($importedCSV)
{
    $firstRecord = $importedCSV | Select-Object -First 1
    $validCSV = $true

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

function Download-AllEmails($emailList)
{
    $i = 1
    foreach ($email in $emailList)
    {
        Write-Progress -Activity "Downloading emails..." -Status $i
        $emailContent = Get-Email -MailboxAddress $email.'Mailbox Address' -InternetMessageId $email.'Internet Message ID'
        Download-Email -MailboxAddress $email.'Mailbox Address' -MessageId $emailContent.id -InternetMessageId $email.'Internet Message ID'
        $i++
    }
}

function Get-Email($mailboxAddress, $internetMessageId)
{
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

function Download-Email($mailboxAddress, $messageId, $internetMessageId)
{
    # https://learn.microsoft.com/en-us/graph/outlook-get-mime-message
    $uri = "https://graph.microsoft.com/v1.0/users/$mailboxAddress/messages/$messageId/`$value"
    $mailboxAddressModified = $mailboxAddress.Replace('@', '_').Replace('.', '_')
    $internetMessageIdModified = $internetMessageId.Replace('<', '').Replace('>', '')
    $filePath = "$PSScriptRoot\$($mailboxAddressModified)_$internetMessageIdModified.eml"
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
TryConnect-MgGraph
$emailList = PromptFor-EmailList
Confirm-CSVHasCorrectHeaders $emailList
Download-AllEmails $emailList
Write-Host "All done!" -ForegroundColor $successColor
Read-Host "Press Enter to exit"

# Should script be updated to provide option to forward them to a specific mailbox instead of (or in addition to) downloading them?

<#
Useful articles:
https://learn.microsoft.com/en-us/graph/outlook-get-mime-message
https://learn.microsoft.com/en-us/graph/api/message-get
https://learn.microsoft.com/en-us/graph/api/user-list-messages
#>