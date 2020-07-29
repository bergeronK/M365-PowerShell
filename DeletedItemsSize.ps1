$mailboxes = @(Get-Mailbox -ResultSize Unlimited)
$report = @()
 
foreach ($mailbox in $mailboxes)
{
    $inboxstats = Get-MailboxFolderStatistics $mailbox -FolderScope DeletedItems
 
    $mbObj = New-Object PSObject
    $mbObj | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $mailbox.DisplayName
    $mbObj | Add-Member -MemberType NoteProperty -Name "Inbox Size (Mb)" -Value $inboxstats.FolderandSubFolderSize.ToMB()
    $mbObj | Add-Member -MemberType NoteProperty -Name "Inbox Items" -Value $inboxstats.ItemsinFolderandSubfolders
    $report += $mbObj
}
 
$report
