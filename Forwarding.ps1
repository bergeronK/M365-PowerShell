# Loop though the object returned by Get-Mailbox with each element represented by $mailbox
foreach ($mailbox in (Get-MailBox -ResultSize Unlimited -OrganizationalUnit "elant.local/SITES [NEW]/Choice"))
{
# Create the forwarding address string
$ForwardingAddress= $mailbox.SamAccountName + “@elantcare.org”
# Check there isn’t a contact, then add one
If (!(Get-MailContact $ForwardingAddress -ErrorAction SilentlyContinue))
{
New-MailContact $mailbox.displayName-ExternalEmailAddress $ForwardingAddress -OrganizationalUnit "elant.local/SITES [NEW]/Choice/Office365Contacts" | Set-MailContact -HiddenFromAddressListsEnabled $true
}
# Set the forwarding address
Set-Mailbox -Identity $mailbox.SamAccountName -ForwardingAddress $ForwardingAddress -DeliverToMailboxAndForward $true
}

