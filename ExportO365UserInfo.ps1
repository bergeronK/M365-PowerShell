Param($migrationCSVFileName = "migration.csv")


function O365Logon

{

	#Check for current open O365 sessions and allow the admin to either use the existing session or create a new one

	$session = Get-PSSession | ?{$_.ConfigurationName -eq 'Microsoft.Exchange'}

	if($session -ne $null)

	{

		$a = Read-Host "An open session to Office 365 already exists.  Do you want to use this session?  Enter y to use the open session, anything else to close and open a fresh session."

		if($a.ToLower() -eq 'y')

		{

			Write-Host "Using existing Office 365 Powershell Session." -ForeGroundColor Green

			return	

		}

		$session | Remove-PSSession

	}

	Write-Host "Please enter your Office 365 credentials" -ForeGroundColor Green

	$cred = Get-Credential

	$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $cred -Authentication Basic -AllowRedirection

	$importresults = Import-PSSession $s

}


function Main

{


	#Verify the migration CSV file exists

	if(!(Test-Path $migrationCSVFileName))

	{

		Write-Host "File $migrationCSVFileName does not exist." -ForegroundColor Red

		Exit

	}


	#Import user list from migration.csv file

	$MigrationCSV = Import-Csv $migrationCSVFileName


	#Get mailbox list based on email addresses from CSV file

	$MailBoxList = $MigrationCSV | %{$_.EmailAddress} | Get-Mailbox

	$Users = @()


	#Get LegacyDN, Tenant, and On-Premise Email addresses for the users

	foreach($user in $MailBoxList)

	{

		$UserInfo = New-Object System.Object


		$CloudEmailAddress = $user.EmailAddresses | ?{($_ -match 'onmicrosoft') -and ($_ -cmatch 'smtp:')}	

		if ($CloudEmailAddress.Count -gt 1)

		{

			$CloudEmailAddress = $CloudEmailAddress[0].ToString().ToLower().Replace('smtp:', '')

			Write-Host "$user returned more than one cloud email address.  Using $CloudEmailAddress" -ForegroundColor Yellow

		}

		else

		{

			$CloudEmailAddress = $CloudEmailAddress.ToString().ToLower().Replace('smtp:', '')

		}


		$UserInfo | Add-Member -Type NoteProperty -Name LegacyExchangeDN -Value $user.LegacyExchangeDN	

		$UserInfo | Add-Member -Type NoteProperty -Name CloudEmailAddress -Value $CloudEmailAddress

		$UserInfo | Add-Member -Type NoteProperty -Name OnPremiseEmailAddress -Value $user.PrimarySMTPAddress.ToString()


		$Users += $UserInfo

	}


	#Check for existing csv file and overwrite if needed

	if(Test-Path ".\cloud.csv")

	{

		$delete = Read-Host "The file cloud.csv already exists in the current directory.  Do you want to delete it?  Enter y to delete, anything else to exit this script."

		if($delete.ToString().ToLower() -eq 'y')

		{

			Write-Host "Deleting existing cloud.csv file" -ForeGroundColor Red

			Remove-Item ".\cloud.csv"

		}

		else

		{

			Write-Host "Will NOT delete current cloud.csv file.  Exiting script." -ForeGroundColor Green

			Exit

		}

	}

	$Users | Export-CSV -Path ".\cloud.csv" -notype

	(Get-Content ".\cloud.csv") | %{$_ -replace '"', ''} | Set-Content ".\cloud.csv" -Encoding Unicode

	Write-Host "CSV File Successfully Exported to cloud.csv" -ForeGroundColor Green


}


O365Logon

Main
The following script converts on-premises Exchange 2007 mailboxes to MEUs. Run this script after you have ran the script to collect information from the cloud mailboxes.

Copy the script below to a .txt file and then save the file and give it a filename Exchange2007MBtoMEU.ps1.

 param($DomainController = [String]::Empty)


function Main

{

	#Script Logic flow

	#1. Pull User Info from cloud.csv file in the current directory

	#2. Lookup AD Info (DN, mail, proxyAddresses, and legacyExchangeDN) using the SMTP address from the CSV file

	#3. Save existing proxyAddresses

	#4. Add existing legacyExchangeDN's to proxyAddresses

	#5. Delete Mailbox

	#6. Mail-Enable the user using the cloud email address as the targetAddress

	#7. Disable RUS processing

	#8. Add proxyAddresses and mail attribute back to the object

	#9. Add msExchMailboxGUID from cloud.csv to the user object (for offboarding support)


	if($DomainController -eq [String]::Empty)

	{

		Write-Host "You must supply a value for the -DomainController switch" -ForegroundColor Red

		Exit

	}


	$CSVInfo = Import-Csv ".\cloud.csv"

	foreach($User in $CSVInfo)

	{

		Write-Host "Processing user" $User.OnPremiseEmailAddress -ForegroundColor Green

		Write-Host "Calling LookupADInformationFromSMTPAddress" -ForegroundColor Green

		$UserInfo = LookupADInformationFromSMTPAddress($User)


		#Check existing proxies for On-Premise and Cloud Legacy DN's as x500 proxies.  If not present add them.

		$CloudLegacyDNPresent = $false

		$LegacyDNPresent = $false

		foreach($Proxy in $UserInfo.ProxyAddresses)

		{

			if(("x500:$UserInfo.CloudLegacyDN") -ieq $Proxy)

			{

				$CloudLegacyDNPresent = $true

			}

			if(("x500:$UserInfo.LegacyDN") -ieq $Proxy)

			{

				$LegacyDNPresent = $true

			}

		}

		if(-not $CloudLegacyDNPresent)

		{

			$X500Proxy = "x500:" + $UserInfo.CloudLegacyDN

			Write-Host "Adding $X500Proxy to EmailAddresses" -ForegroundColor Green

			$UserInfo.ProxyAddresses += $X500Proxy

		}

		if(-not $LegacyDNPresent)

		{

			$X500Proxy = "x500:" + $UserInfo.LegacyDN

			Write-Host "Adding $X500Proxy to EmailAddresses" -ForegroundColor Green

			$UserInfo.ProxyAddresses += $X500Proxy

		}


		#Disable Mailbox

		Write-Host "Disabling Mailbox" -ForegroundColor Green

		Disable-Mailbox -Identity $UserInfo.OnPremiseEmailAddress -DomainController $DomainController -Confirm:$false


		#Mail Enable

		Write-Host "Enabling Mailbox" -ForegroundColor Green

		Enable-MailUser  -Identity $UserInfo.Identity -ExternalEmailAddress $UserInfo.CloudEmailAddress -DomainController $DomainController


		#Disable RUS

		Write-Host "Disabling RUS" -ForegroundColor Green

		Set-MailUser -Identity $UserInfo.Identity -EmailAddressPolicyEnabled $false -DomainController $DomainController


		#Add Proxies and Mail

		Write-Host "Adding EmailAddresses and WindowsEmailAddress" -ForegroundColor Green

		Set-MailUser -Identity $UserInfo.Identity -EmailAddresses $UserInfo.ProxyAddresses -WindowsEmailAddress $UserInfo.Mail -DomainController $DomainController


		#Set Mailbox GUID.  Need to do this via S.DS as Set-MailUser doesn't expose this property.

		$ADPath = "LDAP://" + $DomainController + "/" + $UserInfo.DistinguishedName

		$ADUser = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $ADPath

		$MailboxGUID = New-Object -TypeName System.Guid -ArgumentList $UserInfo.MailboxGUID

		[Void]$ADUser.psbase.invokeset('msExchMailboxGUID',$MailboxGUID.ToByteArray())

		Write-Host "Setting Mailbox GUID" $UserInfo.MailboxGUID -ForegroundColor Green

		$ADUser.psbase.CommitChanges()


		Write-Host "Migration Complete for" $UserInfo.OnPremiseEmailAddress -ForegroundColor Green

		Write-Host ""

		Write-Host ""

	}

}


function LookupADInformationFromSMTPAddress($CSV)

{

	$Mailbox = Get-Mailbox $CSV.OnPremiseEmailAddress -ErrorAction SilentlyContinue


	if($Mailbox -eq $null)

	{

		Write-Host "Get-Mailbox failed for" $CSV.OnPremiseEmailAddress -ForegroundColor Red

		continue

	}


	$UserInfo = New-Object System.Object


	$UserInfo | Add-Member -Type NoteProperty -Name OnPremiseEmailAddress -Value $CSV.OnPremiseEmailAddress

	$UserInfo | Add-Member -Type NoteProperty -Name CloudEmailAddress -Value $CSV.CloudEmailAddress

	$UserInfo | Add-Member -Type NoteProperty -Name CloudLegacyDN -Value $CSV.LegacyExchangeDN

	$UserInfo | Add-Member -Type NoteProperty -Name LegacyDN -Value $Mailbox.LegacyExchangeDN

	$ProxyAddresses = @()

	foreach($Address in $Mailbox.EmailAddresses)

	{

		$ProxyAddresses += $Address

	}

	$UserInfo | Add-Member -Type NoteProperty -Name ProxyAddresses -Value $ProxyAddresses

	$UserInfo | Add-Member -Type NoteProperty -Name Mail -Value $Mailbox.WindowsEmailAddress

	$UserInfo | Add-Member -Type NoteProperty -Name MailboxGUID -Value $CSV.MailboxGUID

	$UserInfo | Add-Member -Type NoteProperty -Name Identity -Value $Mailbox.Identity

	$UserInfo | Add-Member -Type NoteProperty -Name DistinguishedName -Value (Get-User $Mailbox.Identity).DistinguishedName


	$UserInfo

}


Main