<#
This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING 
BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. We grant You a nonexclusive, 
royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that 
You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and 
(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys' fees, 
that arise or result from the use or distribution of the Sample Code.
This posting is provided "AS IS" with no warranties, and confers no rights. 
Use of included script samples are subject to the terms specified at http://www.microsoft.com/info/cpyright.htm.

#>

##################   Function to Expand Group Membership ################
function getMemberExpanded
{
        param ($dn)
                
        $adobject = [adsi]"LDAP://$dn"
        $colMembers = $adobject.properties.item("member")
        Foreach ($objMember in $colMembers)
        {
                $objMembermod = $objMember.replace("/","\/")
                $objAD = [adsi]"LDAP://$objmembermod"
                $attObjClass = $objAD.properties.item("objectClass")
                if ($attObjClass -eq "group")
                {
			  getmemberexpanded $objMember           
                }   
                else
                {
			$colOfMembersExpanded += ,$objMember
		}
        }    
$colOfMembersExpanded 
}    

########################### Function to Calculate Password Age ##############
Function getUserAccountAttribs
{
                param($objADUser,$parentGroup)
		$objADUser = $objADUser.replace("/","\/")
                $adsientry=new-object directoryservices.directoryentry("LDAP://$objADUser")
                $adsisearcher=new-object directoryservices.directorysearcher($adsientry)
                $adsisearcher.pagesize=1000
                $adsisearcher.searchscope="base"
                $colUsers=$adsisearcher.findall()
                foreach($objuser in $colUsers)
                {
                	$dn=$objuser.properties.item("distinguishedname")
	                $sam=$objuser.properties.item("samaccountname")
        	        $attObjClass = $objuser.properties.item("objectClass")
			If ($attObjClass -eq "user")
			{
				$description=$objuser.properties.item("description")
                		$lastlogontimestamp=$objuser.properties.item("lastlogontimestamp")
	                	$accountexpiration=$objuser.properties.item("accountexpires")
        	        	$pwdLastSet=$objuser.properties.item("pwdLastSet")
                		if ($pwdLastSet -gt 0)
                        	{
                        		$pwdLastSet=[datetime]::fromfiletime([int64]::parse($pwdLastSet))
                                	$PasswordAge=((get-date) - $pwdLastSet).days
                        	}
                        	Else {$PasswordAge = "<Not Set>"}                                                                        
                		$uac=$objuser.properties.item("useraccountcontrol")
                        	$uac=$uac.item(0)
                		if (($uac -bor 0x0002) -eq $uac) {$disabled="TRUE"}
                        	else {$disabled = "FALSE"}
                        	if (($uac -bor 0x10000) -eq $uac) {$passwordneverexpires="TRUE"}
                        	else {$passwordNeverExpires = "FALSE"}
                        }                                                        
                        $record = "" | select-object SAM,DN,MemberOf,pwdAge,disabled,pWDneverExpires
                        $record.SAM = [string]$sam
                        $record.DN = [string]$dn
                        $record.memberOf = [string]$parentGroup
                        $record.pwdAge = $PasswordAge
                        $record.disabled= $disabled
                        $record.pWDneverExpires = $passwordNeverExpires
                                
                } 
$record
}
####### Function to find all Privileged Groups in the Forest ##########
Function getForestPrivGroups
{
                $colOfDNs = @()
                $Forest = [System.DirectoryServices.ActiveDirectory.forest]::getcurrentforest()
		$RootDomain = [string]($forest.rootdomain.name)
		$forestDomains = $forest.domains
		$colDomainNames = @()
		ForEach ($domain in $forestDomains)
		{
			$domainname = [string]($domain.name)
			$colDomainNames += $domainname
		}
		
                $ForestRootDN = FQDN2DN $RootDomain
		$colDomainDNs = @()
		ForEach ($domainname in $colDomainNames)
		{
			$domainDN = FQDN2DN $domainname
			$colDomainDNs += $domainDN	
		}

		$GC = $forest.FindGlobalCatalog()
                $adobject = [adsi]"GC://$ForestRootDN"
        	$RootDomainSid = New-Object System.Security.Principal.SecurityIdentifier($AdObject.objectSid[0], 0)
		$RootDomainSid = $RootDomainSid.toString()
		$colDASids = @()
		ForEach ($domainDN in $colDomainDNs)
		{
			$adobject = [adsi]"GC://$domainDN"
        		$DomainSid = New-Object System.Security.Principal.SecurityIdentifier($AdObject.objectSid[0], 0)
			$DomainSid = $DomainSid.toString()
			$daSid = "$DomainSID-512"
			$colDASids += $daSid
		}
			

		$colPrivGroups = @("S-1-5-32-544";"S-1-5-32-548";"S-1-5-32-549";"S-1-5-32-551";"$rootDomainSid-519";"$rootDomainSid-518")
		$colPrivGroups += $colDASids
                
		$searcher = $gc.GetDirectorySearcher()
		ForEach($privGroup in $colPrivGroups)
                {
                                $searcher.filter = "(objectSID=$privGroup)"
                                $Results = $Searcher.FindAll()
                                ForEach ($result in $Results)
                                {
                                                $dn = $result.properties.distinguishedname
                                                $colOfDNs += $dn
                                }
                }
$colofDNs
}

########################## Function to Generate Domain DN from FQDN ########
Function FQDN2DN
{
	Param ($domainFQDN)
	$colSplit = $domainFQDN.Split(".")
	$FQDNdepth = $colSplit.length
	$DomainDN = ""
	For ($i=0;$i -lt ($FQDNdepth);$i++)
	{
		If ($i -eq ($FQDNdepth - 1)) {$Separator=""}
		else {$Separator=","}
		[string]$DomainDN += "DC=" + $colSplit[$i] + $Separator
	}
	$DomainDN
}

########################## MAIN ###########################
$forestPrivGroups = GetForestPrivGroups
$colAllPrivUsers = @()

$rootdse=new-object directoryservices.directoryentry("LDAP://rootdse")

Foreach ($privGroup in $forestPrivGroups)
{
                Write-Host ""
		Write-Host "Enumerating $privGroup.." -foregroundColor yellow
                $uniqueMembers = @()
                $colOfMembersExpanded = @()
		$colofUniqueMembers = @()
                $members = getmemberexpanded $privGroup
                If ($members)
                {
                                $uniqueMembers = $members | sort-object -unique
				$numberofUnique = $uniqueMembers.count
				Foreach ($uniqueMember in $uniqueMembers)
				{
					 $objAttribs = getUserAccountAttribs $uniqueMember $privGroup
                                         $colOfuniqueMembers += $objAttribs      
	
				}
                                $colAllPrivUsers += $colOfUniqueMembers
                                
                                
                }
                Else {$numberofUnique = 0}
                
                If ($numberofUnique -gt 25)
                {
                                Write-host "...$privGroup has $numberofUnique unique members" -foregroundColor Red
                }
		Else { Write-host "...$privGroup has $numberofUnique unique members" -foregroundColor White }
                ForEach($user in $colOfuniquemembers)
                {
                                $userpwdAge = $user.pwdAge
                                $userSAM = $user.SAM
                                IF ($userpwdAge -gt 365)
                                {Write-host "......$userSAM has a password age of $userpwdage" -foregroundColor Green}
                }
                
                

}

$colAllPrivUsers | Export-CSV ".\AllPrivUsers.csv"
