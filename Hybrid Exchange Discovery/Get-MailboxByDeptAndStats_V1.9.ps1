########################################################################################################################################################################################################
<#	5/25/2022
	.VERSION 1.9

	.AUTHOR Bunch of folks
	
	.SYNOPSIS
		Get Exchange mailbox information for users on dept and other mailbox statistics, including folder counts, and export the information to csv file identified with 			timestamp of the output. This report is critical for planning Remote Move migrations in an Exchange hybrid deployment. 
	
	.DESCRIPTION
		The report output includes the following fields - Name,SamAccountName,Email,UserPrincipalName,Department,Database,MailboxSize,MailboxItemCount,
		TotalDeletedItemSize,DeletedItemCount,FolderCount,RecipientTypeDetails,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota

		Run on an Exchange server with Active Directory powershell module installed. 
		Report is output to the current working directory unless otherwise specified.
#>
########################################################################################################################################################################################################

$Results = ForEach ($User in (Get-ADUser -Filter * -Properties Department,Mail))

{   $Mailbox = Get-Mailbox $User.Name -ErrorAction SilentlyContinue

    If ($Mailbox)

    {   $Mail = $Mailbox | Get-MailboxStatistics -ErrorAction SilentlyContinue
	  $MbFolders = $Mailbox | Get-MailboxFolderStatistics -ErrorAction SilentlyContinue

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
		Database = $Mailbox.Database
            MailboxSize = $TotalSize
    		MailboxItemCount = $Mail.ItemCount
		TotalDeletedItemSize = $Mail.TotalDeletedItemSize
            DeletedItemCount = $Mail.DeletedItemCount
            FolderCount = $MbFolders.FolderPath.Count
            RecipientTypeDetails = $Mailbox.RecipientTypeDetails
		IssueWarningQuota = $Mailbox.IssueWarningQuota
            ProhibitSendQuota = $Mailbox.ProhibitSendQuota
            ProhibitSendReceiveQuota = $Mailbox.ProhibitSendReceiveQuota

        }

    }

}

$Results | Select Name,SamAccountName,Email,UserPrincipalName,Department,MailboxSize,MailboxItemCount,TotalDeletedItemSize,DeletedItemCount,FolderCount,RecipientTypeDetails,Database,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota | Export-Csv .\"$((Get-Date).ToString("yyyyMMdd_HHmmss"))_MbSizesByDept.csv" -NoTypeInformation 
