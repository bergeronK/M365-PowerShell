$OutFile = “.\Exported_List_of_ALL_Access_Permissions.csv”
“DisplayName” + “^” + “Email Address” + “^” + “Full Access” + “^” + “Send As” + “^” + “Send On Behalf Of” | Out-File $OutFile -Force
 
$Mailboxes = Get-Mailbox -resultsize unlimited | Select Identity, Alias, DisplayName, DistinguishedName, WindowsEmailAddress
ForEach ($Mailbox in $Mailboxes) {
#$SendOnBehalfOf = Get-mailbox $Mailbox.identity | select Alias, @{Name=’GrantSendOnBehalfTo’;Expression={[string]::join(“;”, ($_.GrantSendOnBehalfTo))}}

$SendOnBehalfOf = Get-mailbox $Mailbox.identity | % {$_.GrantSendOnBehalfTo}

$SendAs = Get-ADPermission $Mailbox.identity | where {($_.ExtendedRights -like “*Send-As*”) -and -not ($_.User -like “NT AUTHORITY\SELF”) -and -not ($_.User -like “s-1-5-21*”)} | % {$_.User}

#$FullAccess = Get-MailboxPermission $Mailbox.Identity | ? {$_.AccessRights -eq “FullAccess” -and !$_.IsInherited} | % {$_.User}
 
$FullAccess = Get-MailboxPermission $Mailbox.Identity | ?{($_.IsInherited -eq $False) -and -not ($_.User -match “NT AUTHORITY”)} |Select User,Identity,@{Name=”AccessRights”;Expression={$_.AccessRights}} | % {$_.User}

$Mailbox.DisplayName + “^” + $Mailbox.WindowsEmailAddress + “^” + $FullAccess + “^” + $SendAs + “^” + $SendOnBehalfOf  | Out-File $OutFile -Append    }