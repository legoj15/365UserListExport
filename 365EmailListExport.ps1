# Check if the required modules are installed
$requiredModules = 'ExchangeOnlineManagement', 'Microsoft.Graph', 'ImportExcel'
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "$module module is not installed. Installing now..."
        Install-Module -Name $module -Force
    }
}

# Import the required modules
echo "Importing ExchangeOnlineManagement PowerShell Module"
Import-Module ExchangeOnlineManagement
echo "Importing Microsoft Graph PowerShell Module"
Import-Module Microsoft.Graph
echo "Importing ImportExcel PowerShell Module"
Import-Module ImportExcel

# Connect to Exchange Online
Connect-ExchangeOnline

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All"

# Get the tenant's name
$tenantName = (Get-MgOrganization).DisplayName

# Get the list of mailboxes excluding DiscoveryMailbox
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -ne "DiscoveryMailbox"}

# Create a hashtable to store GUID to UserPrincipalName mapping
$guidToUserPrincipalName = @{}

# Populate the hashtable with GUID and UserPrincipalName
foreach ($mailbox in $mailboxes) {
    $guidToUserPrincipalName[$mailbox.Name.ToString()] = $mailbox.UserPrincipalName
}

# Output the hashtable for debugging
Write-Host "GUID to UserPrincipalName mapping:"
$guidToUserPrincipalName.GetEnumerator() | ForEach-Object { Write-Host "$($_.Key) : $($_.Value)" }

# Create an empty array to hold the results
$results = @()

# Loop through each mailbox
foreach ($mailbox in $mailboxes) {
    # Get the user and their licenses
    $user = Get-MgUser -UserId $mailbox.UserPrincipalName
    $userLicenses = Get-MgUserLicenseDetail -UserId $mailbox.UserPrincipalName

    # Check MFA status
    #$mfaStatus = if ($user.StrongAuthenticationMethods -ne $null) {"Enabled"} else {"Disabled"}

    # Determine the forwarding address
    $forwardingTo = if ($mailbox.ForwardingSmtpAddress -ne $null -and $mailbox.ForwardingSmtpAddress -ne "") {
        $mailbox.ForwardingSmtpAddress -replace "smtp:", ""
    } elseif ($mailbox.ForwardingAddress -ne $null -and $mailbox.ForwardingAddress -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        if ($guidToUserPrincipalName.ContainsKey($mailbox.ForwardingAddress)) {
            $guidToUserPrincipalName[$mailbox.ForwardingAddress]
        } else {
            Write-Host "GUID not found in hashtable: $($mailbox.ForwardingAddress)"
            $mailbox.ForwardingAddress
        }
    } else {
        $mailbox.ForwardingAddress
    }

    # Filter and format aliases
    $aliases = $mailbox.EmailAddresses | Where-Object {
        $_ -notmatch '^SPO:' -and $_ -notmatch [regex]::Escape($mailbox.PrimarySmtpAddress) -and $_ -notmatch 'onmicrosoft\.com'
    } | ForEach-Object { $_ -replace 'smtp:', '' }

    # Format licenses
    $licenses = $userLicenses.SkuPartNumber -join ", "

    # Output the forwarding address and aliases for debugging
    Write-Host "Mailbox: $($mailbox.UserPrincipalName), ForwardingTo: $forwardingTo, Aliases: $($aliases -join ', '), Licenses: $licenses"

    # Add the mailbox, license, MFA status, and forwarding information to the results
    $results += New-Object PSObject -Property @{
        "Display Name" = $mailbox.DisplayName
        "Primary Email Address" = $mailbox.PrimarySmtpAddress
        Licenses = (($licenses).Replace("O365_BUSINESS_PREMIUM", "Microsoft 365 Business Standard").Replace("FLOW_FREE", "Microsoft Power Automate Free").Replace("ENTERPRISEPACK","Office 365 E3").Replace("SPE_E3", "Microsoft 365 E3").Replace("POWER_BI_STANDARD",  "Microsoft Fabric (Free)").Replace("O365_BUSINESS_ESSENTIALS","Microsoft 365 Business Basic").Replace("EXCHANGEENTERPRISE", "Exchange Online (Plan 2)").Replace("RIGHTSMANAGEMENT_ADHOC", "Rights Management Adhoc").Replace("AAD_PREMIUM_P2", "Microsoft Entra ID P2").Replace("INTUNE_A", "Intune"))
        #"MFA Status" = $mfaStatus
        RecipientTypeDetails = $mailbox.RecipientTypeDetails
        "Account Creation Date" = $mailbox.WhenMailboxCreated
        "Forwarding to" = $forwardingTo
        "Keep mail if forwarding?" = $mailbox.DeliverToMailboxAndForward
        Aliases = ($aliases -join ", ")
    }
}

# Get the list of groups and distribution lists
$groups = Get-DistributionGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, ManagedBy

# Get the list of Microsoft 365 groups and Teams
$unifiedGroups = Get-UnifiedGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, ManagedBy

# Create an array to hold the group results
$groupResults = @()

# Loop through each group and format the results
foreach ($group in $groups) {
    # Exclude "All Company" and the group with the tenant's name
    if ($group.DisplayName -ne "All Company" -and $group.DisplayName -ne $tenantName) {
        # Get the members of the group
        $members = Get-DistributionGroupMember -Identity $group.PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySmtpAddress

        # Convert ManagedBy GUID to UserPrincipalName
        $managedBy = $group.ManagedBy | ForEach-Object {
            if ($guidToUserPrincipalName.ContainsKey($_.ToString())) {
                $guidToUserPrincipalName[$_.ToString()]
            } else {
                $_.ToString()
            }
        }

        $groupResults += New-Object PSObject -Property @{
            "Group Name" = $group.DisplayName
            "Primary Email Address" = $group.PrimarySmtpAddress
            "Managed By" = ($managedBy -join ", ")
            "Members" = ($members -join ", ")
        }
    }
}

# Loop through each unified group and format the results
foreach ($group in $unifiedGroups) {
    # Exclude "All Company" and the group with the tenant's name
    if ($group.DisplayName -ne "All Company" -and $group.DisplayName -ne $tenantName) {
        # Get the members of the group
        $members = Get-UnifiedGroupLinks -Identity $group.PrimarySmtpAddress -LinkType Members | Select-Object -ExpandProperty PrimarySmtpAddress

        # Convert ManagedBy GUID to UserPrincipalName
        $managedBy = $group.ManagedBy | ForEach-Object {
            if ($guidToUserPrincipalName.ContainsKey($_.ToString())) {
                $guidToUserPrincipalName[$_.ToString()]
            } else {
                $_.ToString()
            }
        }

        $groupResults += New-Object PSObject -Property @{
            "Group Name" = $group.DisplayName
            "Primary Email Address" = $group.PrimarySmtpAddress
            "Managed By" = ($managedBy -join ", ")
            "Members" = ($members -join ", ")
        }
    }
}

# Get the current date and time
$currentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Export the mailbox results to the first sheet
echo "Exporting users&groups_$currentDateTime.xlsx"
$results | Select-Object "Display Name", "Primary Email Address", Licenses, "Forwarding to", "Keep mail if forwarding?", RecipientTypeDetails, Aliases, "Account Creation Date" | Export-Excel -Path "users&groups_$currentDateTime.xlsx" -WorkSheetname "Mailboxes" -AutoSize -FreezeTopRow -AutoFilter

# Export the group results to the second sheet
$groupResults | Select-Object "Group Name", "Primary Email Address", "Managed By", "Members" | Export-Excel -Path "users&groups_$currentDateTime.xlsx" -WorkSheetname "Groups" -AutoSize -FreezeTopRow -AutoFilter -Append

# Disconnect from Exchange Online
echo "Disconnecting from Microsoft 365 for ExchangeOnlineManagement PowerShell Module"
Disconnect-ExchangeOnline -Confirm:$false
echo "Disconnecting from Microsoft Graph PowerShell Module"
Disconnect-MgGraph
