#Import the module, requires that you are administrator and are able to run the script
Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
#connect specifying username, if you already have authenticated to another moduel, you actually do not have to authenticate
$credential = Get-Credential
Connect-EXOPSSession -Credential $Credential
#This will make sure when you need to reauthenticate after 1 hour that it uses existing token and you don't have to write password and stuff
$global:UserPrincipalName=$credential
#Disconnect
Get-Pssession | Remove-PSSession