#get-mailbox -Identity hubbp | Select-Object ServerName, Database


$Users = Get-content -path "C:\Scripts\Users.txt"


foreach($user in $users) {
	get-mailbox -Identity $User | Select-Object displayName, PrimarySmtpAddress, ServerName, Database
}