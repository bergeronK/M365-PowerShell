Import-Csv C:\users\administrator.fiestajewelry\desktop\FiestaLogonCSVtest.csv | ForEach-Object {
$UPN = $_.UserPrincipalName; Get-ADUser -Filter { UserPrincipalName -Eq $UPN } | Set-ADUser -Add @{ProxyAddresses=("SMTP:"+$_.EmailAddress)}
}
  


