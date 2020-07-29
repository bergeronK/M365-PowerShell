$Results = ForEach ($User in (Get-ADUser -Filter * -Properties Department,Mail))

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

            Department = $User.Department

            MailboxSize = $TotalSize

            IssueWarningQuota = $Mailbox.IssueWarningQuota

            ProhibitSendQuota = $Mailbox.ProhibitSendQuota

            ProhibitSendReceiveQuota = $Mailbox.ProhibitSendReceiveQuota

            MailboxItemCount = $Mail.ItemCount

        }

    }

}

$Results | Select Name,SamAccountName,Email,Department,MailboxSize,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,MailboxItemCount | Export-Csv c:\Scripts\7-2MailboxSizeByDepartment.csv -NoTypeInformation 
