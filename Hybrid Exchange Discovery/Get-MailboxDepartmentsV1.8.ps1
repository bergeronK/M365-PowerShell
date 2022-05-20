$Users = Get-mailbox -resultsize unlimited
#Import-csv .\users.csv
$Results = ForEach ($User in $Users)

{   $Mailbox = Get-Mailbox $User.Name -ErrorAction SilentlyContinue

    If ($Mailbox)

    {   $Mail = $Mailbox | Get-MailboxStatistics -ErrorAction SilentlyContinue

        If ($Mail.TotalItemSize.Value -eq $null)

        {   $TotalSize = 0

        }

        Else

        {   $TotalSize = $Mail.TotalItemSize.Value.ToMB()

        }

        New-Object PSObject -Property @{

            Name = $User.Name

            SamAccountName = $User.SamAccountName

            Email = $User.Mail

            UserPrincipalName = $User.UserPrincipalName

		Department = $User.Department

		Database = $User.Database

            MailboxSize = $TotalSize

            RecipientTypeDetails = $User.RecipientTypeDetails

		IssueWarningQuota = $Mailbox.IssueWarningQuota

            ProhibitSendQuota = $Mailbox.ProhibitSendQuota

            ProhibitSendReceiveQuota = $Mailbox.ProhibitSendReceiveQuota

            MailboxItemCount = $Mail.ItemCount

        }

    }

}

$Results | Select Name,SamAccountName,Email,UserPrincipalName,Department,MailboxSize,MailboxItemCount,RecipientTypeDetails,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,Database | Export-Csv c:\Scripts\MailboxSizeByDepartment.csv -NoTypeInformation 
