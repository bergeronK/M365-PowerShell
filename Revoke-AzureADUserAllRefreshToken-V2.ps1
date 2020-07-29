
<#    

.NOTES

#=============================================
# Script      : Revoke-AzureADUserAllRefreshToken-V2.ps1
# Created     : ISE 3.0 
# Author(s)   : Casey.Dedeal 
# Date        : 03/01/2020 02:00:56 
# Org         : ETC Solutions
# File Name   : Revoke-AzureADUserAllRefreshToken-V2.ps1
# Comments    : Remove all user AzureAD Tokens O365
# Assumptions : O365 Tenant 
#==============================================

 
SYNOPSIS           : Revoke-AzureADUserAllRefreshToken-V2.ps1
DESCRIPTION        : Remove all user tokens dues to security concerns
Acknowledgements   : Open license
Limitations        : None
Known issues       : None
Credits            : Please visit: 
                          https://simplepowershell.blogspot.com
                          https://msazure365.blogspot.com
.EXAMPLE

  .\Revoke-AzureADUserAllRefreshToken-V2.ps1

  MAP:
  -----------
  #(1)_.Create a Log folder
  #(2)_.Check Folder existence
  #(3)_.Timestamp function
  #(4)_.Function Connect Azure AD
  #(5)_.Attempt to connect AzureAD PowerShell
  #(6)_.Capture User Instence Name
  #(7)_.Check AzureAD User existance 
  #(8)_.Create Custom PS Object store User data
  #(9)_.Provide Choice to Start AzureAD Revoke All Fresh Tokens, or cancel
  #(10)_.Revoke AzureAD All Fresh Tokens
  #(11)_.Create Custom PS Object store changed User data
  #(12)_.Present Results

#>

#(1)_.Create a Log folder
$name      = 'AzureAD-UserToken-Report'
$now       = Get-Date -format 'dd-MMM-yyyy-HH-mm'
$user      = $env:USERNAME
$desFol    = "C:\temp\Reports_\Report-$name\"
$filename  = "$name-$now.CSV"
$file      = $desFol + $filename

#(2)_.Check Folder existance
If (!(Test-Path $desFol)) {New-Item -ItemType Directory -Force $desFol | Out-Null}
#(3)_.Timestamp function
function Function-Get-TimeStamp {
    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

#(4)_.Function Connect Azure AD
function Function-Connect-AzureAD {

  [CmdletBinding()]
  param()
   
try { 
	$connection = Get-AzureADTenantDetail -ErrorAction Stop
    write-host 'Already connected AzureAD PowerShell' 
    Write-host 'Time:' -NoNewline; Function-Get-TimeStamp 
    Write-host $null
} 
catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]{ 
	Write-Host 'You are not connected to AzureAD PowerShell'   
    write-host 'Connecting AzureAD now' -f DarkCyan
    Connect-AzureAD -ErrorAction Stop
    }	  
	
}


#(5)_.Attempt to connect AzureAD PowerShell
try {
    Function-Connect-AzureAD -ErrorAction Stop
} 
catch{ 
	Write-host 'ERROR FOUND:'
    write-host "ERROR: $($PSItem.ToString())"
}	  
	

#(6)_.Capture User Instance Name
Clear-Host
$user = Read-host 'Provide user Instance name' 

#(7)_.Check AzureAD User existance 
Try{

 $Userdata =  Get-AzureADUser | ?{$_.DisPlayName -eq $user} 
 $ObjectId = (Get-AzureADUser | ?{$_.DisPlayName -eq $user}).Objectid 


 if ($Userdata){
 Write-host "Located user ($user)" -f DarkCyan

 #(8)_.Create Custom PS Object store User data
$objectProperty = [ordered]@{
    DisplayName         = $Userdata.DisplayName
    UPN                 = $Userdata.UserPrincipalName
    AccStaus            = $Userdata.AccountEnabled
    TokenRefreshTime    = $Userdata.RefreshTokensValidFromDateTime
}

$Results = New-Object -TypeName psobject -Property $objectProperty
$Results | fl

#(9)_.Provide Choice to Start AzureAD Revoke All Fresh Tokens, or cancel
 Read-host 'Press <ENTER> to Revoke all Tokens'

#(10)_.Revoke AzureAD All Fresh Tokens
 Revoke-AzureADUserAllRefreshToken -ObjectId  $ObjectId
 write-host 'Completed succesfully' -ForegroundColor DarkGray
$Userdata = Get-AzureADUser | ?{$_.DisPlayName -eq $user} 
$objectProperty = [ordered]@{
    DisplayName         = $Userdata.DisplayName
    UPN                 = $Userdata.UserPrincipalName
    AccStaus            = $Userdata.AccountEnabled
    TokenRefreshTime    = $Userdata.RefreshTokensValidFromDateTime
}

#(11)_.Create Custom PS Object store changed User data
$Results = New-Object -TypeName psobject -Property $objectProperty
$Results | export-Csv -Path $file -NoTypeInformation -Append
$Results | fl

 }else{

 Write-host "(a)_.CANNOT Locate user ($user)" -f DarkYellow
 Write-host '(b)_.Check user name, and try again' -f DarkYellow
 Write-Host '(c)_.Script will stop' -f DarkYellow
 break;

  }

}catch 
{

Write-host 'ERROR FOUND:'
write-host "ERROR: $($PSItem.ToString())"

}

#(12)_.Present Results
write-host 'Default access Token Lifetime expiration is upto <1hr>' -f DarkYellow
write-host '<TokenRefreshTime> value indicates when new token was issued' -f DarkYellow
Read-host 'Press <ENTER> to open report folder'
start $desFol
