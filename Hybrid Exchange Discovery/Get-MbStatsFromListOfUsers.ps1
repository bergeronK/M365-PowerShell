
$Users = Import-csv .\import.csv
$Results = ForEach ($User in $Users)

{   $Mailbox = Get-Mailbox $User.email -ErrorAction SilentlyContinue

    If ($Mailbox)

    {   $Mail = $Mailbox | Get-MailboxStatistics -ErrorAction SilentlyContinue
        $MbFolders = $Mailbox | Get-MailboxFolderStatistics -ErrorAction SilentlyContinue

        If ($Mail.TotalItemSize.Value -eq $null)

        {   $TotalSize = 0

        }

        Else

        {   $TotalSize = $Mail.TotalItemSize.value

        }

        New-Object PSObject -Property @{

            Email = $mailbox.primarysmtpaddress
            
            MailboxSize = $TotalSize

            MailboxItemCount = $Mail.ItemCount

            TotalDeletedItemSize = $Mail.TotalDeletedItemSize

            DeletedItemCount = $Mail.DeletedItemCount

            FolderCount = $MbFolders.folderpath.count
        }

    }

}

$Results | Select Email,MailboxSize,MailboxItemCount,TotalDeletedItemSize,DeletedItemCount,FolderCount | Export-Csv .\"$((Get-Date).ToString("yyyyMMdd_HHmmss"))_MbStats.csv" -NoTypeInformation 