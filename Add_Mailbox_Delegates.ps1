# Version 1.0

# functions
function Show-Introduction
{
    Write-Host "This script adds a list of delegates to a mailbox." -ForegroundColor "DarkCyan"
    Read-Host "Press Enter to continue"
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor DarkCyan
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Read-Host -Prompt "Failed to connect to Exchange Online. Press Enter to try again"
        }
    }
}

function TryConnect-AzureAD
{
    $connected = Test-ConnectedToAzureAD

    while (-not($connected))
    {
        Write-Host "Connecting to Azure AD..." -ForegroundColor "DarkCyan"
        Connect-AzureAD -ErrorAction SilentlyContinue | Out-Null

        $connected = Test-ConnectedToAzureAD
        if (-not($connected))
        {
            Write-Warning "Failed to connect to Azure AD."
            Read-Host "Press Enter to try again"
        }
    }
}

function Test-ConnectedToAzureAD
{
    try
    {
        Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue | Out-Null
    }
    catch
    {
        return $false
    }
    return $true
}

function PromptFor-Mailbox
{
    $mailboxEmail = Read-Host "Enter the email address of the mailbox"
    $mailboxEmail = $mailboxEmail.Trim()
    if ($mailboxEmail -eq "" )
    {
        Read-Host "Mailbox email was blank. Press Enter to exit"
        exit
    }
    $mailbox = Get-ExoMailbox -Identity $mailboxEmail -ErrorAction "Stop"
    return $mailbox
}

function PromptFor-DelegateList
{
    Write-Host "Script requires CSV list of delegates and must include headers named `"DelegateEmail`" and `"AccessRights`"." -ForegroundColor "DarkCyan"
    Write-Host "AccessRights can accept the values `"FullAccess`" and `"SendAs`"." -ForegroundColor "DarkCyan"
    $csvPath = Read-Host "Enter path to CSV (must be .csv)"
    $csvPath = $csvPath.Trim('"')
    if ($csvPath -eq "")
    {
        Read-Host "CSV path is blank. Press Enter to exit"
        exit
    }
    return @(Import-Csv -Path $csvPath)
}

function Confirm-CSVHasCorrectHeaders($importedCSV)
{
    $firstRecord = $importedCSV | Select-Object -First 1
    $validCSV = $true

    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name "DelegateEmail"))
    {
        Write-Warning "This CSV file is missing a header called 'DelegateEmail'."
        $validCSV = $false
    }

    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name "FullAccess"))
    {
        Write-Warning "This CSV file is missing a header called 'FullAccess'."
        $validCSV = $false
    }

    if (-not($validCSV))
    {
        Write-Host "Please make corrections to the CSV."
        Read-Host "Press Enter to exit"
        Exit
    }
}

function Prompt-YesOrNo($question)
{
    Write-Host "$question`n[Y] Yes  [N] No"

    do
    {
        $response = Read-Host
        $validResponse = $response -imatch '^\s*[yn]\s*$' # regex matches y or n but allows spaces
        if (-not($validResponse)) 
        {
            Write-Warning "Please enter y or n."
        }
    }
    while (-not($validResponse))

    if ($response -imatch '^\s*y\s*$') # regex matches a y but allows spaces
    {
        return $true
    }
    return $false
}

function Add-Delegates($mailbox, $delegateList, $excludeDisabledUsers)
{
    if ( $null -eq $delegateList ) { return }
    $i = 0
    foreach ($delegate in $delegateList)
    {
        $delegateEmail = $delegate.DelegateEmail.Trim()
        Write-Progress -Activity "Adding delegates..." -Status "$i delegates added."
        if ($excludeDisabledUsers)
        {
            $userEnabled = Confirm-UserEnabled $delegateEmail
            if ($null -eq $userEnabled)
            { 
                Log-Warning "The user $delegateEmail was not found. Skipping user."
                continue 
            }

            if (-not($userEnabled)) 
            { 
                Log-Warning "The user $delegateEmail is disabled. Skipping user."
                continue 
            }
        }
        
        Grant-MailboxAccess -MailboxUpn $mailbox.UserPrincipalName -Delegate $delegate
        $i++
    }
    Write-Host "Finished granting access to users (if they didn't already have the access)." -ForegroundColor "Green"
}

function Confirm-UserEnabled($upn)
{
    $upn = $upn.Trim()
    try 
    {
        $user = Get-AzureADUser -ObjectId $upn -ErrorAction "SilentlyContinue"
    }
    catch 
    {
        # The try catch is just here because otherwise the errors are not suppressed with this cmdlet when a user is not found.
    }    
    if ($null -eq $user) { return }
    return $user.AccountEnabled
}

function Grant-MailboxAccess($mailboxUpn, $delegate)
{
    if ($null -eq $mailboxUpn) { return }
    if ($null -eq $delegate) { return }
    if ($null -eq $delegate.DelegateEmail) { return}
    $mailboxUpn = $mailboxUpn.Trim()
    $delegateEmail = $delegate.DelegateEmail.Trim()
    $accessRights = $delegate.AccessRights.Trim()
    
    try 
    {
        if ($accessRights -eq "FullAccess")
        {
            Add-MailboxPermission -Identity $mailboxUpn -User $delegateEmail -AccessRights "FullAccess" -Confirm:$false -ErrorAction "Stop" | Out-Null
            return
        }
        elseif ($accessRights -eq "SendAs")
        {
            Add-RecipientPermission -Identity $mailboxUpn -Trustee $delegateEmail -AccessRights "SendAs" -Confirm:$false -ErrorAction "Stop" | Out-Null
            return
        }
    }
    catch 
    {
        $errorRecord = $_
        Log-Warning "An error occurred when granting $accessRights permission to $delegateEmail : `n$errorRecord"
    }

    if ( ($accessRights -ne "FullAccess") -and ($accessRights -ne "SendAs")  )
    {
        Log-Warning "AccessRights of $accessRights is invalid for user $delegateEmail. Specify either `"Full Access`" or `"Send As`"."
    }
}

function Log-Warning($message, $logPath = "$PSScriptRoot\logs.txt")
{
    $message = "[$(Get-Date -Format 'yyyy-MM-dd hh:mm tt') W] $message"
    Write-Output $message | Tee-Object -FilePath $logPath -Append | Write-Host -ForegroundColor "Yellow"
}

# main
Show-Introduction
TryConnect-ExchangeOnline
$mailbox = PromptFor-Mailbox
$delegateList = PromptFor-DelegateList
$excludeDisabledUsers = Prompt-YesOrNo "Exclude disabled users from script inputs? (Takes longer)"
if ($excludeDisabledUsers) { TryConnect-AzureAD }
Add-Delegates -Mailbox $mailbox -DelegateList $delegateList -ExcludeDisabledusers $excludeDisabledUsers
Write-Host "All done!" -ForegroundColor "Green"
Read-Host -Prompt "Press Enter to exit"