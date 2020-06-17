#Script to Change the UPN on the Active Directory#
#This script should run from an Active Directory Module for Windows PowerShell#
#Version 1.0 – 06/30/2015#
#Author: Antonio Vargas#

$UserCount = 0

Import-Csv -Path C:\FiestaLogonCSVTest.csv | ForEach-Object {

$UPN = $_.UserPrincipalName
Write-Host “Working on user:” $UPN
Get-ADUser -Filter {UserPrincipalName -Eq $UPN} | Set-AdUser -userprincipalname $_.EmailAddress
$usercount = $usercount +1
}
Write-Host “Number of users on your CSV: $UserCount”
Write-Host “UPN’s Changed”