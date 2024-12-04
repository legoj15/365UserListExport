import subprocess
import sys
import os
import ctypes
import tempfile

def is_admin():
    try:
        return os.getuid() == 0
    except AttributeError:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0

def run_powershell_script(script_path):
    # Set the execution policy scope for the current process to RemoteSigned and run the PowerShell script
    subprocess.run(["powershell", "-Command", "Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned;" + script_path], check=True)

if __name__ == "__main__":
    if is_admin():
        # PowerShell script as a string
        ps_script = """
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

# Connect to Exchange Online
Connect-ExchangeOnline

# Connect to Azure AD
Connect-MsolService

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
    $user = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName

    # Check MFA status
    $mfaStatus = if ($user.StrongAuthenticationMethods -ne $null) {"Enabled"} else {"Disabled"}

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

    # Output the forwarding address and aliases for debugging
    Write-Host "Mailbox: $($mailbox.UserPrincipalName), ForwardingTo: $forwardingTo, Aliases: $($aliases -join ', ')"

    # Add the mailbox, license, MFA status, and forwarding information to the results
    $results += New-Object PSObject -Property @{
        "Display Name" = $mailbox.DisplayName
        "Primary Email Address" = $mailbox.PrimarySmtpAddress
        Licenses = (($user.Licenses.AccountSkuId -join ", ").Replace("KIS:", "").Replace("reseller-account:", "").Replace("PhilSmithAutomotiveGroup:", "").Replace(", FLOW_FREE", "").Replace("O365_BUSINESS_ESSENTIALS", "Microsoft 365 Business Basic").Replace("O365_BUSINESS_PREMIUM", "Microsoft 365 Business Standard").Replace("ENTERPRISEPACK", "Office 365 E3"))
        "MFA Status" = $mfaStatus
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
    # Get the members of the group
    $members = Get-DistributionGroupMember -Identity $group.PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySmtpAddress

    $groupResults += New-Object PSObject -Property @{
        "Group Name" = $group.DisplayName
        "Primary Email Address" = $group.PrimarySmtpAddress
        "Managed By" = ($group.ManagedBy -join ", ")
        "Members" = ($members -join ", ")
    }
}

# Loop through each unified group and format the results
foreach ($group in $unifiedGroups) {
    # Get the members of the group
    $members = Get-UnifiedGroupLinks -Identity $group.PrimarySmtpAddress -LinkType Members | Select-Object -ExpandProperty PrimarySmtpAddress

    $groupResults += New-Object PSObject -Property @{
        "Group Name" = $group.DisplayName
        "Primary Email Address" = $group.PrimarySmtpAddress
        "Managed By" = ($group.ManagedBy -join ", ")
        "Members" = ($members -join ", ")
    }
}

# Get the current date and time
$currentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Export the mailbox results to the first sheet
echo "Exporting users_$currentDateTime.xlsx"
$results | Select-Object "Display Name", "Primary Email Address", Licenses, "Forwarding to", "Keep mail if forwarding?", RecipientTypeDetails, Aliases, "MFA Status", "Account Creation Date" | Export-Excel -Path "users_$currentDateTime.xlsx" -WorkSheetname "Mailboxes" -AutoSize -FreezeTopRow -AutoFilter

# Export the group results to the second sheet
$groupResults | Select-Object "Group Name", "Primary Email Address", "Managed By", "Members" | Export-Excel -Path "users_$currentDateTime.xlsx" -WorkSheetname "Groups" -AutoSize -FreezeTopRow -AutoFilter -Append

# Disconnect from Exchange Online
echo "Disconnecting from Microsoft 365 for ExchangeOnlineManagement PowerShell Module"
Disconnect-ExchangeOnline -Confirm:$false

        """

        # Write the PowerShell script to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ps1") as temp:
            temp.write(ps_script.encode())
            temp_path = temp.name

        # Run the PowerShell script
        run_powershell_script(temp_path)

        # Delete the temporary file
        os.unlink(temp_path)
    else:
        # Re-run the program with admin rights
        if sys.platform == 'win32':
            subprocess.run(['runas', '/user:Administrator', 'python'] + sys.argv)
        else:
            print("Please run the script as root.")
