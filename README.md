# 365UserListExport
Exports a list of all mailboxes associated with a Microsoft 365 admin account, including assigned licenses, mailbox forwarding recipients, and MFA status, to an xlsx format Excel spreadsheet

## Full list of currently exported data, in order
- User's Display Name (the name that email clients show instead of the email address)
- User's Primary Email Address
- Licenses that the user has assigned (minus Automate Free)
- Who that user's mailbox is being forwarded to, if any
- Whether or not forwarded mail is kept in the mailbox after forwarding
- RecipientTypeDetails, i.e. User, DiscoveryMailbox, ScheduelingMailbox, etc
- Any mail aliases that the user has associated with them
- MFA status (True or False)
- Date account was created

The output XLSX file has its columns automatically sized, so that the final output is already clean and readable

Here is an example of the final output:
![image](https://github.com/legoj15/365UserListExport/assets/7399802/aa72498b-25ba-40f0-b422-11a13fde35ee)
