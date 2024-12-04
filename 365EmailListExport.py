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

    # Output the forwarding address for debugging
    Write-Host "Mailbox: $($mailbox.UserPrincipalName), ForwardingTo: $forwardingTo"

    # Add the mailbox, license, MFA status, and forwarding information to the results
    $results += New-Object PSObject -Property @{
        "Display Name" = $mailbox.DisplayName
        "Primary Email Address" = $mailbox.PrimarySmtpAddress
        Licenses = (($user.Licenses.AccountSkuId -join ", ").Replace("KIS:", "").Replace("reseller-account:", "").Replace("PhilSmithAutomotiveGroup:", "").Replace("FLOW_FREE, ", "").Replace("O365_BUSINESS_ESSENTIALS", "Microsoft 365 Business Basic").Replace("O365_BUSINESS_PREMIUM", "Microsoft 365 Business Standard").Replace("ENTERPRISEPACK", "Office 365 E3").Replace(", FLOW_FREE", ""))
        "MFA Status" = $mfaStatus
        RecipientTypeDetails = $mailbox.RecipientTypeDetails
        "Account Creation Date" = $mailbox.WhenMailboxCreated
        "Forwarding to" = $forwardingTo
        "Keep mail if forwarding?" = $mailbox.DeliverToMailboxAndForward
        Aliases = ($aliases -join ", ")
    }
}

# Get the current date and time
$currentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Export the results to an XLSX file with auto-sized columns, top row frozen, and filtering enabled
echo "Exporting users_$currentDateTime.xlsx"
$results | Select-Object "Display Name", "Primary Email Address", Licenses, "Forwarding to", "Keep mail if forwarding?", RecipientTypeDetails, Aliases, "MFA Status", "Account Creation Date" | Export-Excel -Path "users_$currentDateTime.xlsx" -AutoSize -FreezeTopRow -AutoFilter

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
