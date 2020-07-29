Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0") `
-Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName`
|?{$_ -notmatch "_none_"} | select -First 1)

$EXOSession = New-ExoPSSession

Import-PSSession $EXOSession