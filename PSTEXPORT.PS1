$export = get-content c:\mailbox.csv
$Export|%{$_|New-MailboxExportRequest -FilePath "\\servername\pst\$($_.alias).pst"} 

