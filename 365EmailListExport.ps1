# Check if the required modules are installed
$requiredModules = 'ExchangeOnlineManagement', 'MSOnline', 'ImportExcel'
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "$module module is not installed. Installing now..."
        Install-Module -Name $module -Force
    }
}

# Import the required modules
echo "Importing ExchangeOnlineManagement PowerShell Module"
Import-Module ExchangeOnlineManagement
echo "Importing Microsoft Online PowerShell Module"
Import-Module MSOnline
echo "Importing ImportExcel PowerShell Module"
Import-Module ImportExcel

# Store your username and app password (hackers please look away)
#$username = "your-username"
#$password = ConvertTo-SecureString "your-app-password" -AsPlainText -Force
#$credential = New-Object System.Management.Automation.PSCredential($username, $password)

# Connect to Exchange Online
Connect-ExchangeOnline #-Credential $credential

# Connect to Azure AD
Connect-MsolService #-Credential $credential

# Get the list of mailboxes excluding DiscoveryMailbox
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -ne "DiscoveryMailbox"}

# Create an empty array to hold the results
$results = @()

# Loop through each mailbox
foreach ($mailbox in $mailboxes) {
    # Get the user and their licenses
    $user = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName

    # Check MFA status
    $mfaStatus = if ($user.StrongAuthenticationMethods -ne $null) {"Enabled"} else {"Disabled"}

    # Add the mailbox, license, MFA status, and forwarding information to the results
    $results += New-Object PSObject -Property @{
        "Display Name" = $mailbox.DisplayName
        "Primary Email Address" = $mailbox.PrimarySmtpAddress
        Licenses = (($user.Licenses.AccountSkuId -join ", ").Replace("reseller-account:", "").Replace("FLOW_FREE, ", "").Replace("O365_BUSINESS_ESSENTIALS", "Microsoft 365 Business Basic").Replace("O365_BUSINESS_PREMIUM", "Microsoft 365 Business Standard").Replace("ENTERPRISEPACK", "Office 365 E3").Replace(", FLOW_FREE", ""))
        "MFA Status" = $mfaStatus
        RecipientTypeDetails = $mailbox.RecipientTypeDetails
        "Account Creation Date" = $mailbox.WhenMailboxCreated
        "Forwarding to" = ($mailbox.ForwardingSmtpAddress -replace "smtp:", "")
        "Keep mail if forwarding?" = $mailbox.DeliverToMailboxAndForward
        Aliases = ($aliases -join ", ")
    }
}

# Get the current date and time
$currentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Export the results to an XLSX file with auto-sized columns
echo "Exporting users_$currentDateTime.xlsx"
$results | Select-Object "Display Name", "Primary Email Address", Licenses, "Forwarding to", "Keep mail if forwarding?", RecipientTypeDetails, Aliases, "MFA Status", "Account Creation Date" | Export-Excel -Path "users_$currentDateTime.xlsx" -AutoSize

# Disconnect from Exchange Online
echo "Disconnecting from Microsoft 365 for ExchangeOnlineManagement PowerShell Module"
Disconnect-ExchangeOnline -Confirm:$false
