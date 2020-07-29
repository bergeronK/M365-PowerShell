
  <#
   .SYNOPSIS

         This script will Query all DAGs in your environment and will create table with the following information
        - list DB copies hosted in each server
        - list the activation preference for each database in nice table presentation
        - Highlight the preferred server for the database to be mounted on according to the Activation Preference (Yellow Highlight)
        - Highlight the current server that the database is actually mounted on ( Red highlight if NOT the preferred)
        - Summary for :
                1. Total Number of databases copies on each server
                2. Ideal number of databases that should be mounted on each server
                3. Actual number of databases mounted on each server

        The script will send email with those info at the end




Script Name        :         Exchange DAG Database Distribution Table
Script Version     :         2.0
Author             :         Ammar Hasayen
Blog               :         http://ammarhasayen.com
Script Requirement :         Exchange View Only Administrator


Total Copies: 
represent the total number of database copies (Active and Passive) that are mounted on the server

Ideal Mounted DB Copies : 
According to the Activation Preference, how many databases should be mounted on the server.In other words, how many databases have this server with Activation preference = 1

Actual Mounted DB Copies : 
How many databases actually mounted on this server

Yellow cells:
represent the server on which the database is mounted and it happens that it is mounted on the server with Activation preference = 1

Red cells:
represent the server on which the database is mounted and it happens that it is mounted on the server with Activation preference not equal 1

Green cells:
represent database copy locations with activation preference


    New in Version 2 [April 2015]:

        - Support for Exchange 2013
        - Filter by DAG


.LINK
     My Blog
     http://ammarhasayen.com


 .PARAMETER ScriptFilesPath
     Path to store script files like ".\" to indicate current directory or full path like C:\myfiles
	
	.PARAMETER SendMail
	 Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory
	
	.PARAMETER MailFrom
	 Email address to send from. Passed directly to Send-MailMessage as -From
	
	.PARAMETER MailTo
	 Email address to send to. Passed directly to Send-MailMessage as -To
	
	.PARAMETER MailServer
	 SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer



    .EXAMPLE
     Generate the HTML report
     \Get-CorpDAGLayout.ps1 -ScriptFilesPath .\  

    .EXAMPLE
     Generate the HTML report with SMTP Email option
     \Get-CorpDAGLayout.ps1 -ScriptFilesPath .\  -SendMail:$true -MailFrom noreply@contoso.com  -MailTo me@contoso.com  -MailServer smtp.contoso.com

     .EXAMPLE
     Generate the HTML report with SMTP Email option and DAG Filter
     \Get-CorpDAGLayout.ps1 -ScriptFilesPath .\  -SendMail:$true -MailFrom noreply@contoso.com  -MailTo me@contoso.com  -MailServer smtp.contoso.com -InputDAGs DAG1,DAG2


#>


#region parameters

[cmdletbinding()]

param(
    [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Path to store script files like c:\ ')][string]$ScriptFilesPath,
	[parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Send Mail ($True/$False)')][bool]$SendMail=$false,
	[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail From')][string]$MailFrom,
	[parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail To')]$MailTo,
	[parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail Server')][string]$MailServer,	
    [parameter(Position=5,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Only include those DAGs (eg DAG1,DAG2...)',ParameterSetName="DAGFilter")][array]$InputDAGs = $null,	
    [switch]$DontCheckMountStatus
    )

#endregion parameters

#region functions

function _screenheadings {

        Cls
        write-host 
        write-host 
        write-host 
        write-host "--------------------------" 
        write-host "Script Info" -foreground Green
        write-host "--------------------------"
        write-host
        write-host " Script Name  : Get-CorpDAGLayout"  
        write-host " Author       : Ammar Hasayen (ammarhasayen)"        
        write-host " Version      : 2.0" 
        Write-Host " Release Date : April 2015"  
        write-host
        write-host "--------------------------" 
        write-host "Script Release Notes" -foreground Green
        write-host "--------------------------"
        write-host
        write-host "-Account Requirements :View Administrator Exchange Role"         
        write-host
        write-host "-Always check for newer version @ http://ammarhasayen.com."
        write-host 
        write-host "--------------------------" 
        write-host "Script Start" -foreground Green
        write-host "--------------------------"
        Write-Host
        sleep 1
} # function _screenheadings


function Write-CorpError {
            
            [cmdletbinding()]

            param(
                [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Error Variable')]$myError,	
	            [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Additional Info')][string]$Info,
                [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Log file full path')][string]$mypath,
	            [switch]$ViewOnly

                )

                Begin {
       
                    function get-timestamp {

                        get-date -format 'yyyy-MM-dd HH:mm:ss'
                    } 

                } #Begin

                Process {

                    if (!$mypath) {

                        $mypath = " "
                    }

                    if($myError.InvocationInfo.Line) {

                    $ErrorLine = ($myError.InvocationInfo.Line.Trim())

                    } else {

                    $ErrorLine = " "
                    }

                    if($ViewOnly) {

                        Write-warning @"
                        $(get-timestamp)
                        $(get-timestamp): $('-' * 60)
                        $(get-timestamp):   Error Report
                        $(get-timestamp): $('-' * 40)
                        $(get-timestamp):
                        $(get-timestamp): Error in $($myError.InvocationInfo.ScriptName).
                        $(get-timestamp):
                        $(get-timestamp): $('-' * 40)       
                        $(get-timestamp):
                        $(get-timestamp): Line Number: $($myError.InvocationInfo.ScriptLineNumber)
                        $(get-timestamp): Offset : $($myError.InvocationInfo.OffsetLine)
                        $(get-timestamp): Command: $($myError.invocationInfo.MyCommand)
                        $(get-timestamp): Line: $ErrorLine
                        $(get-timestamp): Error Details: $($myError)
                        $(get-timestamp): Error Details: $($myError.InvocationInfo)
"@

                        if($Info) {
                            Write-Warning -Message "More Custom Info: $info"
                        }

                        if ($myError.Exception.InnerException) {

                            Write-Warning -Message "Error Inner Exception: $($myError.Exception.InnerException.Message)"
                        }

                        Write-warning -Message " $('-' * 60)"

                     } #if($ViewOnly) 

                     else {
                     # if not view only 
        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): $('-' * 60)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):  Error Report"        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error in $($myError.InvocationInfo.ScriptName)."        
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp):"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Line Number: $($myError.InvocationInfo.ScriptLineNumber)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Offset : $($myError.InvocationInfo.OffsetLine)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Command: $($myError.invocationInfo.MyCommand)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Line: $ErrorLine"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error Details: $($myError)"
                        Log-Write -LogFullPath $mypath -LineValue "$(get-timestamp): Error Details: $($myError.InvocationInfo)"
                        if($Info) {
                            Log-Write -LogFullPath $mypath -LineValue  "$(get-timestamp): More Custom Info: $info"
                        }

                        if ($myError.Exception.InnerException) {

                            Log-Write -LogFullPath $mypath -LineValue  "$(get-timestamp) :Error Inner Exception: $($myError.Exception.InnerException.Message)"
            
                        }    

                     }# if not view only

               } # End Process

        } # function Write-CorpError


function Get-CorpDagMemberList{    
            <#    
            .SYNOPSIS
                Give me DAG name and will return string array of members

                OR will return $null if error happens

                Input should be a string that represent the name of the DAG

                Output is either array string or null

                Version                    :         1.0
                Author                     :         Ammar Hasayen (@ammarhasayen)(http://ammarhasayen.com)        


            .PARAMETER $DatabaseAvailabilityGroup
             String name of the database availability group.

             .EXAMPLE     
            .\Get-CorpDagMemberList -$DatabaseAvailabilityGroup "NYC"

            #>  

            [cmdletbinding()]

            param(

                [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,HelpMessage='String name of the database availability group')]
                [ValidateNotNullorEmpty()]
                [string]$DatabaseAvailabilityGroup
            )

            #region defining variables

                $DatabaseAvailabilityGroup_Members = @()

                $DatabaseAvailabilityGroup_Members_List = @()            

            #endregion #region defining variables

            try {
                $DatabaseAvailabilityGroup_Members = (Get-DatabaseAvailabilityGroup $DatabaseAvailabilityGroup -ErrorAction STOP).servers 

                $DatabaseAvailabilityGroup_Members_List = $DatabaseAvailabilityGroup_Members |foreach{$_.name}


            }catch {
                $DatabaseAvailabilityGroup_Members_List = $null
            }    
        

            Write-Output $DatabaseAvailabilityGroup_Members_List

     } # function Get-CorpDagMemberList


function _parameterGetExch {

    # This function will filter Exchange server based on 
    # the Exchange Server Objects matching that filter.

    param()        
   
            $myExchangeServers = @()
            $var = @()
            $var_dag= @()
    
            foreach ($InputDAG in $InputDAGs) {
                #string array of DAG server members
                $var = Get-CorpDagMemberList $InputDAG
                if ($var) {
                    $var_dag += $var
                }
            } #end foreach ($InputDAG in $InputDAGs)

                     
            foreach ($inputServer in $var_dag) {
                try {
                    $var_Srv = Get-ExchangeServer $inputServer -ErrorAction stop
                    $myExchangeServers += $var_Srv
                }catch {
                    Write-Verbose -Message "[_parameterGetExch] Error :Server $inputServer fails Get-ExchangeServer Command. Skipping"
                }
            } #end foreach ($inputServer in $var_dag)


            if (!$myExchangeServers) {
               throw "[_parameterGetExch] Error : No Exchange Servers matched."
            }
         
            #Get-Unique in case of duplicate user input for the same server.
            Write-Output ($myExchangeServers |Get-Unique)

} # function _parameterGetExch 


function _parameterGetDBs {

        param ($E2010, $E2013, $ExchangeServersList)
        
        $mydatabases = @()
 
        if ($E2010) {

            if ($E2013) {
                        	
                    $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2013 -Status)  | Where {$ExchangeServersList -contains $_.Server -and $_.recovery -eq $false  }
                              
            } # if ($E2013)

            elseif ($E2010) {
                        	
                    $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2010 -Status)  | Where {$ExchangeServersList -contains $_.Server -and $_.recovery -eq $false  } 

            } # elseif

            } # if ($E2010)
                     
            else { # if ($E2010)
                        
            $mydatabases = [array](Get-MailboxDatabase -IncludePreExchange2007 -Status) | Where {$ExchangeServersList -contains $_.Server }

            }#else 
         

        Write-Output $mydatabases

 } # function _parameterGetDBs


function _parameterGetDAG {

            param ($InputDAGs,$ExchangeServersList)
    
            
            $myDAGs = @()

            foreach ($inputDAG in $InputDAGs) {

                try{
                    $var = Get-DatabaseAvailabilityGroup $inputDAG -ErrorAction STOP
                    $myDAGs += $var
                }
                Catch {
                    Write-Verbose -Message "[_parameterGetDAG] Error: DAG $InputDAG : fails Get-DatabaseAvailabilityGroup Command. Skiping it"
                }


            }#foreach
                       

            Write-Output $myDAGs


} # function _parameterGetDAG

Function sendEmail 
{ param($from,$to,$subject,$smtphost,$htmlFileName) 


        $msg = new-object Net.Mail.MailMessage
        $smtp = new-object Net.Mail.SmtpClient($smtphost)
        $msg.From = $from
        $msg.To.Add($to)
        $msg.Subject = $subject
        $msg.Body = Get-Content $htmlFileName 
        $msg.isBodyhtml = $true 
        $smtp.Send($msg)
       

} 


function _GetDAG
{
	param($DAG)
	@{Name			= $DAG.Name.ToUpper()
	  MemberCount	= $DAG.Servers.Count
	  Members		= [array]($DAG.Servers | % { $_.Name })
	  Databases		= @()
	  }
}



function   _GetDB 
{
  param($Database)

  $DB_Act_pref = $Database.ActivationPreference
  $Mounted = $Database.Mounted
  $DB_Act_pref = $Database.ActivationPreference
  [array]$DBHolders =$null 
		( $Database.Servers) |%{$DBHolders  += $_.name}





     @{Name						= $Database.Name
	  ActiveOwner				= $Database.Server.Name.ToUpper()	 
	  Mounted                   = $Mounted
	  DBHolders			        = $DBHolders
	  DB_Act_pref               = $DB_Act_pref 	  
	  IsRecovery                = $Database.Recovery
	  }

}


function _GetDAG_DB_Layout
{
	param($Databases,$DAG)

	    $WarningColor                      = "#FF9900"
		$ErrorColor                        ="#980000"
		$BGColHeader                       ="#000099"
		$BGColSubHeader                    ="#0000FF"
		[Array]$Servers_In_DAG             = $DAG.Members
		
    $Output2 ="<table border=""0"" cellpadding=""3"" width=""50%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
	<col width=""5%"">
	<colgroup width=""25%"">"
	$Servers_In_DAG | Sort-Object | %{$Output2+="<col width=""3%"">"}
	$Output2 +="</colgroup>"
	$ServerCount = $Servers_In_DAG.Count
	
	$Output2 += "<tr bgcolor=""$($BGColHeader)""><th><font color=""#ffffff"">DatabaseCopies</font></th>	
	<th colspan=""$($ServerCount)""><font color=""#ffffff"">Mailbox Servers in $($DAG.name)</font></th>	
	</tr>"
	$Output2+="<tr bgcolor=""$($BGColSubHeader)""><th></th>"
	$Servers_In_DAG|Sort-Object | %{$Output2+="<th><font color=""#ffffff"">$($_)</font></th>"}
	
	$Output2 += "</tr>"
	
	#writing table content
	$AlternateRow=0
		foreach ($Database in $Databases)
	{
	$Output2+="<tr "
	if ($AlternateRow)
					{
						$Output2+=" style=""background-color:#dddddd"""
						$AlternateRow=0
					} else
					{
						$AlternateRow=1
					}
		
		$Output2+="><td><strong>$($database.name)</strong></td>"
		
		#copies
		
							$DatabaseServer   = $Database.ActiveOwner
							$DatabaseServers  = $Database.DBHolders
		$Servers_In_DAG|Sort-Object| 
			%{ 
									 $ActvPref =$Database.DB_Act_pref
									 $server_in_the_loop = $_
									 $Actv = $ActvPref  |where {$_.key -eq  $server_in_the_loop}
									 $Actv=  $Actv.value
									 $ActvKey= $ActvPref |Where {$_.value -eq 1}
									 $ActvKey = 	 $ActvKey.key.name
									  
									  
							$Output2+="<td"
							
								if (  ($DatabaseServers -contains $_) -and ( $_ -like $databaseserver)  )
										{
											if (  $ActvKey -like $databaseserver  )
											{$Output2+=" align=""center"" style=""background-color:#F7FB0B""><font color=""#000000f""><strong>$Actv</strong></font> "}
											else
											{$Output2+=" align=""center"" style=""background-color:#FB0B1B""><strong><font color=""#ffffff"">$Actv</strong></font> "}
										
										}
					
							
										elseif ($DatabaseServers -contains $_)
									{
								
									
									$Output2+=" align=""center"" style=""background-color:#00FF00"">$Actv "							 
									}
									else
								{ $Output2+=" align=""center"" style=""background-color:#dddddd"">"	}
								 
			
			}
				
		
		
		$Output2+="</tr >"
		}
		
	$Output2+="<tr></tr><tr></tr><tr></tr>"
	
	
	#Total Assigned copies
	
	$Output2 += "<tr bgcolor=""#440164""><th><font color=""#ffffff"">Total Copies</font></th>"
	
	$Servers_In_DAG|Sort-Object| 
			%{ 
		
								
								$this = $EnvironmentServers[$_]
									$Output2 += "<td align=""center"" style=""background-color:#E0ACF8""><font color=""#000000""><strong>$($this.DBCopyCount)</strong></font></td>"	
						
		}
	$Output2 +="</tr>"
	#Copies Assigned Ideal
	
	$Output2 += "<tr bgcolor=""#DB08CD""><th><font color=""#ffffff"">Ideal Mounted DB Copies</font></th>"
	
	$Servers_In_DAG|Sort-Object| 
			%{ 
			foreach ($this in $My_Hash_3.GetEnumerator())
						{				
									if ($this.key  -like $_)
									{$Output2 += "<td align=""center"" style=""background-color:#FBCCF9""><font color=""#000000""><strong>$($this.value)</strong></font></td>"}		
						}
		}
	$Output2 +="</tr>"
	
# Copies Actually Assigned
	
	$Output2 += "<tr bgcolor=""#440164""><th><font color=""#ffffff"">Actual Mounted DB Copies</font></th>"
	
	$Servers_In_DAG|Sort-Object| 
			%{ 
		
								
								$this = $EnvironmentServers[$_]
									$Output2 += "<td align=""center"" style=""background-color:#E0ACF8""><font color=""#000000""><strong>$($this.DBCopyCount_Assigned)</strong></font></td>"	
						
		}
	$Output2 +="</tr>"	
		
$Output2


}



#Gets all Mailbox Servers with extra info and return a hashtable
function _GetMailboxServerInfo
{
     param($Server, $databases)

     [int]$DBCopyCount           =  0 # This is considered not initialized
     [int]$DBs_Mountedcount       = 0 

     #Getting DBs mounted on this server
     $DBs_Mounted                      = @($databases | Where {$_.Server -ieq $Server})
     $DBs_MountedCount                 =  $DBs_Mounted.count

     #Getting DB copies on this server
     $MailboxServer                         = Get-MailboxServer $Server
     [array]$MailboxServer_DB_Copies        = @()     
     Try{$MailboxServer_DB_Copies           = $MailboxServer  | 
                                                Get-MailboxDatabaseCopyStatus -ErrorAction SilentlyContinue}
                                                  
                                                  Catch{}


    if ($MailboxServer_DB_copies) #if the server has copies
            {
              $DBCopyCount           = $MailboxServer_DB_copies.Count  
            }


     #Return hashtable
     @{ Name = $server
      DBCopyCount = $DBCopyCount
      DBCopyCount_Assigned = $DBs_MountedCount
      }


    

}


#Gets all Mailbox Servers in the organization
function _GetMailboxServers
{

$ExchServers = Get-ExchangeServer |Where {$_.ServerRole -Contains "Mailbox"}

Return $ExchServers 

}


 #endregion functions

 #_____________________________________________Script Preperation_______________________________

#region screen headings

     Write-Verbose -Message "Info : Starting $($MyInvocation.Mycommand)"  

     Write-verbose -Message ($PSBoundParameters | out-string)

     _screenheadings

#endregion screen headings

#region preperation 

    #region create directory

        try{
         $ScriptFilesPath = Convert-Path $ScriptFilesPath -ErrorAction Stop
        }catch {
            Write-CorpError -myError $_ -ViewOnly -Info "[Validating log files path] Sorry, please check the sript path name again"
        Exit
        throw " [Creating files] Validating log files path] Sorry, please check the sript path name again "
    }

    $ScriptFilesPath = Join-Path $ScriptFilesPath "DagLayout"

    if(Test-Path $ScriptFilesPath ) {
        try {
                Remove-Item $ScriptFilesPath -Force -Recurse -ErrorAction Stop

            }catch {

                Write-CorpError -myError $_ -ViewOnly -Info "[Deleting old working directory] Could not delete directory $ScriptFilesPath"
                Exit
                throw "[Deleting old working directory] Could not delete directory $ScriptFilesPath "

            }
    }
    
    
    if(!(Test-Path $ScriptFilesPath )) {
        try {
            New-Item -ItemType directory -Path $ScriptFilesPath -ErrorAction Stop
        }catch{
            Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Creating working directory] Could not delete directory $ScriptFilesPath"
            Exit
            throw "[Module Factory - Creating working directory] Could not delete directory $ScriptFilesPath "

        }
    }  

    #endregion create directory

    #region create html               
           
            try{                
                $HTMLFile = "DAGLayout.html"    
                $HTMLFileFullPath = Join-Path $ScriptFilesPath  $HTMLFile  -ErrorAction Stop
            }catch {
                Write-CorpError -myError $_ -ViewOnly -Info "[Validating log files] Sorry, please check the sript path name again"
                Exit
                throw " [Creating files] Sorry, please check the sript path name again "

            }
             
            #Check if file exists and delete if it does

            If((Test-Path -Path $HTMLFileFullPath )){
                try {
                Remove-Item -Path $HTMLFileFullPath  -Force -ErrorAction Stop
                }catch {
                    Write-CorpError -myError $_ -ViewOnly -Info "[Deleting old log files]"
                    Exit
                    throw " [Creating files] Sorry, but the script could not delete log file on $HTMLFileFullPath "
                }
            }
    
            #Create HTML file

            try {
                New-Item -Path $ScriptFilesPath -Name $HTMLFile –ItemType File -ErrorAction Stop
            }
            catch {
                 Write-CorpError -myError $_ -ViewOnly -Info "[Module Factory - Creating log files]"
                 Exit
                throw "  [Module Factory - Creating files] Sorry, but the script could not create log file on $ScriptFilesPath [Creating files] " }  

    #endregion create html           

    #region Check Exchange Management Shell, attempt to load  
     
             #Quote : part of the PowerShell Loading code block is inspired from Steve Goodman  
            
             Write-Verbose -Message "Checking PowerShell Environment"
             Write-host
             Write-host  "Checking PowerShell Environment" -ForegroundColor Cyan

             if ((Get-Host).Version.Major -eq 1) {
                    
	             throw "Powershell Version 1 not supported";
             }       

                               
            # Sometime it is tricky to load Exchange Management Shell specially if Exchange was installed on a drive other than the C drive.
            #So we will get the Exchange Installation Path
            [string]$Exch_InstallPath = $env:exchangeinstallpath              
            $Exch_InstallDrive = $Exch_InstallPath.Substring(0,3)                         
            $loadscript1 = Join-Path $Exch_InstallDrive "Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1"  
            $loadscript2 = Join-Path $Exch_InstallDrive "Program Files\Microsoft\Exchange Server\bin\Exchange.ps1"   
            $loadscript3 = Join-Path $Exch_InstallPath  "bin\RemoteExchange.ps1"

            if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
                    
	            if (Test-Path "C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1"){	
                                                     
                    . 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
		            Connect-ExchangeServer -auto
	            }  
                            
                elseif (Test-Path "C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1") {
		            Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.Admin
		            .'C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1'
	            }
                            
                elseif (Test-Path $loadscript1 ) {	
		            . $loadscript1
		            Connect-ExchangeServer -auto
	            }
                            
                elseif (Test-Path $loadscript2) {
		            Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.Admin
		            . $loadscript2
	            }
                            
                elseif (Test-Path $loadscript3 ) { #Exchange 2013
                    . $loadscript3
		            Connect-ExchangeServer -auto
	            }    

                else {
                    throw "Exchange Management Shell cannot be loaded"                            
	            }

            }


            # Check if -SendMail parameter set and if so check -MailFrom, -MailTo and -MailServer are set
            if ($SendMail)
            {
	            if (!$MailFrom -or !$MailTo -or !$MailServer)
	            {

                    throw "If -SendMail specified, you must also specify -MailFrom, -MailTo and -MailServer"
	            }
            }

            # Check Exchange Management Shell Version
            if ((Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue))
            {
	            $E2010 = $false;
	            if (Get-ExchangeServer | Where {$_.AdminDisplayVersion.Major -gt 14})
	            {
		            Write-Warning "Exchange 2010 or higher detected. You'll get better results if you run this script from an Exchange 2010/2013 management shell"

	            }
            }else{
    
                $E2010 = $true

                $varPS = $ErrorActionPreference
                $ErrorActionPreference = "Stop"
                try {
                    $localserver = get-exchangeserver $Env:computername -ErrorAction Stop
                    $localversion = $localserver.admindisplayversion.major
                    if ($localversion -eq 15) { $E2013 = $true }
                }catch {
                Write-Warning -Message " [Module Factory] You are not running the script from an Exchange Server"
                Write-Warning -Message " [Module Factory] The script logic cannot determine if PS version is E2013 or E2010"
                Write-Warning -Message " [Module Factory] Knowing this info is so important to determine the command set to use"
                Write-Warning -Message " [Module Factory] Command failing is (Get-ExchangeServer `$Env:computername) "
                Write-Warning -Message " [Module Factory] `$Env:computername in this case evaluates to $($Env:computername)"
                Write-Warning -Message " [Module Factory] Please run the script from an Exchange Server"
                Write-Warning -Message " [Module Factory] Existing script"
                Write-Host -ForegroundColor Red "Terminating Script.."
                Exit
                Throw " You are not running the script from an Exchange Server so the code decided to Exit"
                }finally {

                $ErrorActionPreference = $varPS
                }

                if(!$localversion) {
                    Write-Warning -Message " [Module Factory] You are not running the script from an Exchange Server"
                    Write-Warning -Message " [Module Factory] The script logic cannot determine if PS version is E2013 or E2010"
                    Write-Warning -Message " [Module Factory] Knowing this info is so important to determine the command set to use"
                    Write-Warning -Message " [Module Factory] Command failing is (Get-ExchangeServer `$Env:computername) "
                    Write-Warning -Message " [Module Factory] `$Env:computername in this case evaluates to $($Env:computername)"
                    Write-Warning -Message " [Module Factory] Please run the script from an Exchange Server"
                    Write-Warning -Message " [Module Factory] Existing script"
                    Write-Host -ForegroundColor Red "Terminating Script.. "
                     
                    Exit
                    Throw "You are not running the script from an Exchange Server so the code decided to Exit"

                }

            }

        #endregion Check Exchange Management Shell, attempt to load 
                

#endregion preperation

#_____________________________________________Script Body_______________________________

#region variables



$SRV                 = @() #Holds name of Exchange Mailbox Servers participating in DAGs
$EnvironmentServers  = @{} # Hashtable to hold Exchange Mailbox Server custom info 
$EnvironmentDAGS     = @() # Hashtable to hold DAG custom info

#endregion variables

#region get databases

     Write-host
     Write-host  "Getting DB Info" -ForegroundColor Cyan

    if ($InputDAGs -ne $null){
        
        Write-host "    DAG Filters detected" -ForegroundColor Yellow

        #region get Exch server info
            $ExchangeServers = [array](_parameterGetExch )    
            Write-host
            write-host "         - Exchange Servers detected : $($ExchangeServers.count)" 
            Write-host  
            $ExchangeServersList = @($ExchangeServers | foreach{ $_.name})

        #endregion get Exch server info

        #region get DB info
            $Databases =  [array](_parameterGetDBs  $E2010 $E2013 $ExchangeServersList) 
            Write-Host "Number of DBs collected : $($Databases.count)"  
        #endregion get DB info
    }


    if ($InputDAGs -eq $null){

         Write-host  "    No DAG Filters detected" -ForegroundColor Yellow
        if ($E2010) {
    
            if ($E2013) {    
                            	
                $Databases = [array](Get-MailboxDatabase -IncludePreExchange2013 -Status)  | 
                              where {$_.recovery -eq $false} 
                Write-Verbose -Message "PowerShell Environment is Exchange 2013"
                                            
            } # if ($E2013)

            elseif ($E2010) {   
                             	
                $Databases = [array](Get-MailboxDatabase  -Status)  | 
                             where {$_.recovery -eq $false} 
                Write-Verbose -Message "PowerShell Environment is Exchange 2010"

            } # elseif

        } # if ($E2010) 

        Write-Host "Number of DBs collected : $($Databases.count)"        
    }


    if ($Databases.count -eq 0) {

            throw "OPS !! Get-MailboxDatabase -stuaus command did not return any database. Exiting"
    }

    
    $My_Databases = $Databases

#endregion get databases


#region get DAG

     Write-host
     Write-host  "Getting DAG Info" -ForegroundColor Cyan

     if ($InputDAGs -ne $null) {

            if ($E2010) {                    

                $DAGs =  [array](_parameterGetDAG $InputDAGs $ExchangeServersList)

                $Dag_Count = $DAGs.count

            }else {
                $DAGs = $null                
            } 
    }
     

     if ($InputDAGs -eq $null) {
        $DAGs   = [array](Get-DatabaseAvailabilityGroup)
     }




     if ($DAGs) {
        Foreach ($DAG in $DAGS){                
          $EnvironmentDAGS += _GetDAG $DAG
        }
     }
     else {
        throw "no Exchange DAGs are detected.. exitings  "
     }



#endregion get dag



#-------------------- START Collecting DB Info-----------------------

for ($i=0; $i -lt $My_Databases.Count; $i++){

     $database = _GetDB $My_Databases[$i]


    for ($j=0; $j -lt $EnvironmentDAGS.Count; $j++)
			{
				if ($EnvironmentDAGS[$j].Members -contains $Database.ActiveOwner)
				{
					$EnvironmentDAGS[$j].Databases += $Database
				}
			}
}

#-------------------- END Collecting DB Info-----------------------


#-------------------- START Collecting Exchange Server Info---------

#collect DAG Exchange Servers

Foreach($DAG in $DAGS){
 Foreach ($Server in $DAG.Servers){
  $SRV+=$Server.name}}
   
 
foreach ($server in $SRV)

        { $SRV_Info  =_GetMailboxServerInfo $server $My_Databases

         $EnvironmentServers.Add($SRV_Info.Name, $SRV_Info) 

         }


#-------------------- END Collecting Exchange Server Info-----------


#-------------------- Start Creating HTML Table-----------


#Hold Info in temp hashtables
$My_Hash_1 = $null
$My_Hash_1 = @{}

$My_Hash_2 = $null
$My_Hash_2 = @{}

$myobjects =$null
$myobjects = @()

$My_Hash_3 =$null
$My_Hash_3 = @{}





        foreach ($My_database in $My_databases)

                {
	                $Var1 =  $My_database.activationPreference |where{$_.value -eq 1}
	                $Var2 = $Var1.key.name
	                $My_Hash_1.add($My_Database.Name,$Var2)
                }

        foreach ($Var3 in $My_Hash_1.GetEnumerator())
                {
                $objx = New-Object System.Object
                $objx | Add-Member -type NoteProperty -name Name -value $Var3.key
                $objx | Add-Member -type NoteProperty -name count -value $Var3.value
                $myobjects +=     $objx
                }


$mydata = $myobjects |Group-Object -Property count

        foreach ($counting in $mydata)
                {

                $My_Hash_3.add($counting.name,$counting.count)
                }

        foreach ($Server in $SRV)
        {
                if(!( $My_Hash_3.ContainsKey($Server)))
                {
                  $My_Hash_3.Add($Server, 0)
                }

        }




$Output ="<html>
<body>
<font size=""1"" face=""Arial,sans-serif"">
<h3 align=""center"">DAG Database Copies layout</h3>
<h5 align=""center"">Generated $((Get-Date).ToString())</h5>
</font>"




        foreach ($DAG in $EnvironmentDAGS )
                {
	                if ($DAG.Membercount -gt 0)
	                        {
		                        # Database Availability Group Header
		                        $Output +="<table border=""0"" cellpadding=""3"" width=""50%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
		                        <col width=""20%""><col width=""10%""><col width=""70%"">
		                        <tr align=""center"" bgcolor=""#FC8E10""><th><font color=""#ffffff"">Database Availability Group Name</font></th><th><font color=""#ffffff"">Member Count</font></th>
		                        <th><font color=""#ffffff"">Database Availability Group Members</font></th></tr>
		                        <tr><td>$($DAG.Name)</td><td align=""center"">
		                        $($DAG.MemberCount)</td><td>"
		                        $DAG.Members | % { $Output+="$($_) " }
		                        $Output +="</td></tr></table>"
		
		
		
		                        # Get Table HTML

                            $Output += _GetDAG_DB_Layout -Databases $DAG.Databases -DAG $DAG
	                       

                            }
	
                }

$Output += "</table>"

$Output+="</body></html>";
Add-Content $HTMLFileFullPath $Output


#-------------------- END Creating HTML Table-----------


#-------------------- START Send Email-------------------

if ($SendMail){
SendEmail  $MailFrom $MailTo "DAG Layout Report _$(Get-Date -f 'yyyy-MM-dd')" $MailServer $HTMLFileFullPath}

#-------------------- END Send Email---------------------

