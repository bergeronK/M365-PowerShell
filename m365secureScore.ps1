# This script needs to be run by an admin account in your Office 365 tenant
# This script will create an Azure AD app in your organisation with permission
# to access resources in yours and your customers' tenants.
   
$externalCSS = "<link rel=`"stylesheet`" href=`"https://dl.dropbox.com/s/vpx9ysgr11cah4u/reports.css?dl=0`">"
$yourLogo = "https://pbs.twimg.com/profile_images/1103610066291314688/pClI5wTS_400x400.png"
$TableHeaderColour = "#00a1f1"
$applicationName = "GCITS Secure Score Exporter"
      
# Change this to true if you would like to overwrite any existing applications with matching names. 
$removeExistingAppWithSameName = $false
# Modify the homePage and logoutURI values to any valid URI that you like.
# They don't need to be actual addresses, so feel free to make something up.
# Set the $appIDUri variable to a use a valid domain in your tenant. eg. https://yourdomain.com/$((New-Guid).ToString())
 
$homePage = "https://daymarklab.com"
$appIdURI = "https://daymarklab.com/$((New-Guid).ToString())"
$logoutURI = "https://portal.office.com"
$ApplicationPermissions = "SecurityEvents.Read.All Directory.Read.All"
   
   
function Confirm-FolderPath($Path) {
    $folder = Test-Path -Path $Path
    if ($folder) {
        Write-Host "Path exists"
    }
    else {
        Write-Host "Creating Temp folder"
        New-Item -Path $Path -ItemType directory
    }
}
   
function New-GCITSITGTableFromArray($Array, $HeaderColour) {
    # Remove any empty properties from table
    $properties = $Array | get-member -ErrorAction SilentlyContinue | Where-Object {$_.memberType -contains "NoteProperty"}
    foreach ($property in $properties) {
        try {
            $members = $Array.$($property.name) | Get-Member -ErrorAction Stop
        }
        catch {
            $Array = $Array | Select-Object -Property * -ExcludeProperty $property.name
        }
    }
    $Table = $Array | ConvertTo-Html -Fragment
    if ($Table[2] -match "<tr>") {
        $Table[2] = $Table[2] -replace "<tr>", "<tr style=`"background-color:$HeaderColour`">"
    }
    return $Table
}
Function Add-ResourcePermission($requiredAccess, $exposedPermissions, $requiredAccesses, $permissionType) {
    foreach ($permission in $requiredAccesses.Trim().Split(" ")) {
        $reqPermission = $null
        $reqPermission = $exposedPermissions | Where-Object {$_.Value -contains $permission}
        Write-Host "Collected information for $($reqPermission.Value) of type $permissionType" -ForegroundColor Green
        $resourceAccess = New-Object Microsoft.Open.AzureAD.Model.ResourceAccess
        $resourceAccess.Type = $permissionType
        $resourceAccess.Id = $reqPermission.Id    
        $requiredAccess.ResourceAccess.Add($resourceAccess)
    }
}
      
Function Get-RequiredPermissions($requiredDelegatedPermissions, $requiredApplicationPermissions, $reqsp) {
    $sp = $reqsp
    $appid = $sp.AppId
    $requiredAccess = New-Object Microsoft.Open.AzureAD.Model.RequiredResourceAccess
    $requiredAccess.ResourceAppId = $appid
    $requiredAccess.ResourceAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]
    if ($requiredDelegatedPermissions) {
        Add-ResourcePermission $requiredAccess -exposedPermissions $sp.Oauth2Permissions -requiredAccesses $requiredDelegatedPermissions -permissionType "Scope"
    } 
    if ($requiredApplicationPermissions) {
        Add-ResourcePermission $requiredAccess -exposedPermissions $sp.AppRoles -requiredAccesses $requiredApplicationPermissions -permissionType "Role"
    }
    return $requiredAccess
}
Function New-AppKey ($fromDate, $durationInYears, $pw) {
    $endDate = $fromDate.AddYears($durationInYears) 
    $keyId = (New-Guid).ToString()
    $key = New-Object Microsoft.Open.AzureAD.Model.PasswordCredential($null, $endDate, $keyId, $fromDate, $pw)
    return $key
}
      
Function Test-AppKey($fromDate, $durationInYears, $pw) {
      
    $testKey = New-AppKey -fromDate $fromDate -durationInYears $durationInYears -pw $pw
    while ($testKey.Value -match "\+" -or $testKey.Value -match "/") {
        Write-Host "Secret contains + or / and may not authenticate correctly. Regenerating..." -ForegroundColor Yellow
        $pw = Initialize-AppKey
        $testKey = New-AppKey -fromDate $fromDate -durationInYears $durationInYears -pw $pw
    }
    Write-Host "Secret doesn't contain + or /. Continuing..." -ForegroundColor Green
    $key = $testKey
      
    return $key
}
      
Function Initialize-AppKey {
    $aesManaged = New-Object "System.Security.Cryptography.AesManaged"
    $aesManaged.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aesManaged.Padding = [System.Security.Cryptography.PaddingMode]::Zeros
    $aesManaged.BlockSize = 128
    $aesManaged.KeySize = 256
    $aesManaged.GenerateKey()
    return [System.Convert]::ToBase64String($aesManaged.Key)
}
function Confirm-MicrosoftGraphServicePrincipal {
    $graphsp = Get-AzureADServicePrincipal -All $true | Where-Object {$_.displayname -eq "Microsoft Graph"}
    if (!$graphsp) {
        $graphsp = Get-AzureADServicePrincipal -SearchString "Microsoft.Azure.AgregatorService"
    }
    if (!$graphsp) {
        Login-AzureRmAccount -Credential $credentials
        New-AzureRmADServicePrincipal -ApplicationId "00000003-0000-0000-c000-000000000000"
        $graphsp = Get-AzureADServicePrincipal -All $true | Where-Object {$_.displayname -eq "Microsoft Graph"}
    }
    return $graphsp
}
   
function Get-GCITSMSGraphResource($Resource) {
    $graphBaseUri = "https://graph.microsoft.com/beta"
    $values = @()
    $result = Invoke-RestMethod -Uri "$graphBaseUri/$resource" -Headers $headers
    if ($result.value) {
        $values += $result.value
        if ($result."@odata.nextLink") {
            do {
                $result = Invoke-RestMethod -Uri $result."@odata.nextLink" -Headers $headers
                $values += $result.value
            } while ($result."@odata.nextLink")
        }
    }
    else {
        $values = $result
    }
    return $values
}
function Get-GCITSAccessToken($appCredential, $tenantId) {
    $client_id = $appCredential.appID
    $client_secret = $appCredential.secret
    $tenant_id = $tenantid
    $resource = "https://graph.microsoft.com"
    $authority = "https://login.microsoftonline.com/$tenant_id"
    $tokenEndpointUri = "$authority/oauth2/token"
    $content = "grant_type=client_credentials&client_id=$client_id&client_secret=$client_secret&resource=$resource"
    $response = Invoke-RestMethod -Uri $tokenEndpointUri -Body $content -Method Post -UseBasicParsing
    $access_token = $response.access_token
    return $access_token
}
   
Confirm-FolderPath -Path C:\temp
Confirm-FolderPath -Path C:\temp\SecureScoreReports
Write-Host "Connecting to Azure AD. The login window may appear behind Visual Studio Code."
Connect-AzureAD
      
Write-Host "Creating partner application in tenant: $((Get-AzureADTenantDetail).displayName)"
      
# Check for the Microsoft Graph Service Principal. If it doesn't exist already, create it.
$graphsp = Confirm-MicrosoftGraphServicePrincipal
   
$existingapp = $null
$existingapp = get-azureadapplication -SearchString $applicationName
if ($existingapp -and $removeExistingAppWithSameName) {
    Remove-Azureadapplication -ObjectId $existingApp.objectId
}
   
$rsps = @()
if ($graphsp) {
      
    $rsps += $graphsp
    $tenantInfo = Get-AzureADTenantDetail
    $tenant_id = $tenantInfo.ObjectId
    $tenantName = $tenantInfo.DisplayName
    $initialDomain = ($tenantInfo.verifiedDomains | Where-Object {$_.Initial}).name
      
    # Add Required Resources Access (Microsoft Graph)
    $requiredResourcesAccess = New-Object System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess]
    $microsoftGraphRequiredPermissions = Get-RequiredPermissions -reqsp $graphsp -requiredApplicationPermissions $ApplicationPermissions -requiredDelegatedPermissions $DelegatedPermissions
    $requiredResourcesAccess.Add($microsoftGraphRequiredPermissions)
      
    # Get an application key
    $pw = Initialize-AppKey
    $fromDate = [System.DateTime]::Now
    $appKey = Test-AppKey -fromDate $fromDate -durationInYears 99 -pw $pw
      
    Write-Host "Creating the AAD application $applicationName" -ForegroundColor Blue
    $aadApplication = New-AzureADApplication -DisplayName $applicationName `
        -HomePage $homePage `
        -ReplyUrls $homePage `
        -IdentifierUris $appIdURI `
        -LogoutUrl $logoutURI `
        -RequiredResourceAccess $requiredResourcesAccess `
        -PasswordCredentials $appKey `
        -AvailableToOtherTenants $true
          
    # Creating the Service Principal for the application
    $servicePrincipal = New-AzureADServicePrincipal -AppId $aadApplication.AppId
      
    Write-Host "Assigning Permissions" -ForegroundColor Yellow
        
    # Assign application permissions to the application
    foreach ($app in $requiredResourcesAccess) {
        $reqAppSP = $rsps | Where-Object {$_.appid -contains $app.ResourceAppId}
        Write-Host "Assigning Application permissions for $($reqAppSP.displayName)" -ForegroundColor DarkYellow
        foreach ($resource in $app.ResourceAccess) {
            if ($resource.Type -match "Role") {
                New-AzureADServiceAppRoleAssignment -ObjectId $serviceprincipal.ObjectId `
                    -PrincipalId $serviceprincipal.ObjectId -ResourceId $reqAppSP.ObjectId -Id $resource.Id
            }
        }
    }
        
    # This provides the application with access to your customer tenants.
    # If you are running this for a single tenant, comment out the following two lines:
    #$group = Get-AzureADGroup -Filter "displayName eq 'Adminagents'"
    #Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $servicePrincipal.ObjectId
    
    Write-Host "App Created" -ForegroundColor Green
   
    [array]$contracts = @{
        DisplayName = $tenantName
        CustomerContextID = $tenant_id
        DefaultDomainName = $initialDomain
    }
    try{
        $contracts += Get-AzureADContract -All $true
    }catch{
        Write-host "Couldn't retrieve customer tenants. Generating report for a single tenant."
    }
       
    $appCredential = @{
        AppId  = $aadApplication.AppId
        Secret = $appkey.value
    }
    foreach ($contract in $contracts) {
        Write-Host "Retrieving Secure Score for $($contract.DisplayName)"
        $tenant_id = $contract.customercontextid
        # Try to execute the API call 6 times
      
        $Stoploop = $false
        [int]$Retrycount = "0"
        do {
            try {
                   
                $access_token = Get-GCITSAccessToken -appCredential $appCredential -tenantId $tenant_id
                $headers = @{
                    Authorization = "Bearer $access_token"
                }
                Write-Host "Retrieved Access Token" -ForegroundColor Green
   
   
                $scores = $null
      
                $scores = Get-GCITSMSGraphResource -Resource security/securescores
                $profiles = Get-GCITSMSGraphResource -Resource "security/secureScoreControlProfiles"
                   
                if ($scores) {
                    $latestScore = $scores[0]
                    $HTMLCollection = @()
         
                    foreach ($control in $latestScore.controlScores) {
                        $controlReport = $null
                        $launchButton = $null
                        $controlProfile = $profiles | Where-Object {$_.id -contains $control.controlname}
                        $controlTitle = "<h3>$($controlProfile.title)</h3>"
                        [int]$controlScoreInt = $control.score
                        [int]$maxScoreInt = $controlProfile.maxScore
                        [string]$controlScore = "<h4>Score: $controlScoreInt/$maxScoreInt</h4>"
                        $assessment = "<strong>Assessment</strong><br>$($control.description)<br>"
                        $remediation = "<strong>Remediation</strong><br>$($controlprofile.remediation)<br>"
                        $remediationImpact = "<strong>Remediation Impact</strong><br>$($controlprofile.remediationImpact)<br>"
                        if ($controlProfile.actionUrl) {
                            $launchButton = "<a class=`"button`" href=`"$($controlProfile.actionUrl)`">Launch</a><br>"
                        }
                        $userImpact = "<strong>User Impact:</strong> $($controlprofile.userImpact)"
                        $implementationCost = "<strong>Implementation Cost:</strong> $($controlprofile.implementationCost)"
                        $threats = "<strong>Threats:</strong> $($controlprofile.threats -join ", ")"
                        $tier = "<strong>Tier:</strong> $($controlprofile.tier)"
                        $hr = "<hr>"
                        [array]$controlElements = $assessment, $remediation, $remediationImpact
                        if ($launchButton) {
                            $controlElements += $launchButton
                        }
                        $controlReport = "<div>$($controlElements -join "</div><div><br></div><div>")</div><div><br></div>$($userImpact,$implementationCost,$threats,$tier,$hr -join "</div><div>")</div>"
                        $controlReport = "$($controlTitle)$($controlScore)<div><br></div>$($controlReport)"
                        $HTMLCollection += [pscustomobject]@{
                            category      = $controlProfile.controlCategory
                            controlReport = [string]$controlReport
                            rank          = $controlProfile.rank
                            deprecated    = $controlProfile.deprecated
                            score         = $control.score
                 
                        }
                    }
         
                    $HTMLCollection = $HTMLCollection | Where-Object {!$_.deprecated} | Sort-Object rank
         
                    $identityControls = $HTMLCollection | Where-Object {$_.category -contains "Identity"}
                    $DataControls = $HTMLCollection | Where-Object {$_.category -contains "Data"}
                    $DeviceControls = $HTMLCollection | Where-Object {$_.category -contains "Device"}
                    $AppsControls = $HTMLCollection | Where-Object {$_.category -contains "Apps"}
                    $InfrastructureControls = $HTMLCollection | Where-Object {$_.category -contains "Infrastructure"}
       
         
                    $identityScore = 0
                    $dataScore = 0
                    $deviceScore = 0
                    $appsScore = 0
                    $infrastructureScore = 0
                    $identityControls | ForEach-Object {$identityScore += $_.score}
                    $DataControls | ForEach-Object {$dataScore += $_.score}
                    $DeviceControls | ForEach-Object {$deviceScore += $_.score}
                    $AppsControls | ForEach-Object {$appsScore += $_.score}
                    $InfrastructureControls | ForEach-Object {$infrastructureScore += $_.score}
         
                    [int]$identityScore = $identityScore
                    [int]$dataScore = $dataScore
                    [int]$deviceScore = $deviceScore
                    [int]$appsScore = $appsScore
                    [int]$infrastructureScore = $infrastructureScore
                    $categoryScores = @()
         
                    $allTenantScores = $latestScore.averageComparativeScores | Where-Object {$_.basis -contains "AllTenants"}
                    $similarCompanyScores = $latestScore.averageComparativeScores | Where-Object {$_.basis -contains "TotalSeats"}
         
                    [int]$maxScore = $latestScore.maxScore
                    [int]$similarCompanyAverage = $similarCompanyScores.averageScore
                    [int]$globalAverage = $allTenantScores.averageScore
                    $minSeat = $similarCompanyScores.seatSizeRangeLowerValue
                    $maxSeat = $similarCompanyScores.seatSizeRangeUpperValue
         
                    $categoryScores += [pscustomobject][ordered]@{
                        Identity = "Tenant score: $($identityScore)"
                        Data     = "Tenant score: $($dataScore)"
                        Device   = "Tenant score: $($deviceScore)"
                    }
                    $categoryScores += [pscustomobject][ordered]@{
                        Identity = "Global average: $($allTenantScores.identityScore)"
                        Data     = "Global average: $($allTenantScores.dataScore)"
                        Device   = "Global average: $($allTenantScores.deviceScore)"
                    }
                    $categoryScores += [pscustomobject][ordered]@{
                        Identity = "Similar sized company average: $($similarCompanyScores.identityScore)"
                        Data     = "Similar sized company average: $($similarCompanyScores.dataScore)"
                        Device   = "Similar sized company average: $($similarCompanyScores.deviceScore)"
                    }
                    # Add Apps and Infrastructure scores to the overview table if they exist. 
                    if ($allTenantScores) {
                        if (($allTenantScores | get-member).name -contains "appsScore") {
                            $categoryScores[0] | Add-Member Apps "Tenant score: $appsScore"
                            $categoryScores[1] | Add-Member Apps "Global average: $($allTenantScores.appsScore)"
                            $categoryScores[2] | Add-Member Apps "Similar sized company average: $($similarCompanyScores.appsScore)"
                        }
                        if (($allTenantScores | get-member).name -contains "infrastructureScore") {
                            $categoryScores[0] | Add-Member Infrastructure "Tenant score: $infrastructureScore"
                            $categoryScores[1] | Add-Member Infrastructure "Global average: $($allTenantScores.infrastructureScore)"
                            $categoryScores[2] | Add-Member Infrastructure "Similar sized company average: $($similarCompanyScores.infrastructureScore)"
                        }
                    }
                 
                    $reportByLine = "<div>Secure Score report compiled by <strong>$tenantName</strong> on $((Get-Date).ToLongDateString())</div>"
                    [int]$currentScore = $($latestScore.currentScore)
                    $customerHeading = "<h1>$($contract.displayname)</h1>$reportByLine<br>"
                    $scoreheading = "<h2>Microsoft Secure Score: $currentScore</h2>"
                    $maxScoreTitle = "<strong>Maximum attainable score:</strong> $maxScore"
                    $similarCompanyTitle = "<strong>Similar sized company average ($minSeat - $maxSeat users):</strong> $similarCompanyAverage"
                    $globalAverageTitle = "<strong>Global average:</strong> $globalAverage"
                    $scoreBreakDownTitle = "<strong>Score Breakdown:</strong>"
                    $scoreBreakdownTable = New-GCITSITGTableFromArray -Array $categoryScores -HeaderColour $TableHeaderColour
                    $subHeadings = "<div>$($maxScoreTitle,$similarCompanyTitle,$globalAverageTitle -join "</div><div>")</div>"
                    $overviewHTML = "$($customerHeading,$scoreheading,$subHeadings,$scoreBreakDownTitle -join "<div><br></div>")$scoreBreakdownTable<br><br>"
                    $identityHTML = "<h2>Identity Controls</h2>$($identityControls.controlReport -join "<div><br></div>")"
                    $dataHTML = "<h2>Data Controls</h2>$($dataControls.controlReport -join "<div><br></div>")"
                    $deviceHTML = "<h2>Device Controls</h2>$($deviceControls.controlReport -join "<div><br></div>")"
                    $appsHTML = "<h2>Apps Controls</h2>$($appsControls.controlReport -join "<div><br></div>")"
                    $infrastructureHTML = "<h2>Infrastructure Controls</h2>$($infrastructureControls.controlReport -join "<div><br></div>")"
                       
                    [array]$completeReport = $overviewHTML
       
                    if ($identityControls) {
                        $completeReport += $identityHTML
                    }
                    if ($dataControls) {
                        $completeReport += $dataHTML
                    }
                    if ($deviceControls) {
                        $completeReport += $deviceHTML
                    }
                    if ($AppsControls) {
                        $completeReport += $appsHTML
                    }
                    if ($InfrastructureControls) {
                        $completeReport += $infrastructureHTML
                    }
       
                    "$externalCSS <img class=`"float-right`" src=$yourLogo> $($completeReport -join "<p></p>")<img class=`"float-right`" src=$yourLogo><br><br><div>Report compiled by <strong>$tenantName</strong> on $((Get-Date).ToLongDateString())</div>" | Out-file C:\temp\securescorereports\$($contract.DefaultDomainName).html
                    [pscustomobject][ordered]@{
                        CustomerName = $contract.DisplayName
                        TenantId = $contract.CustomerContextId
                        CreatedDateTime = $latestScore.createdDateTime
                        SecureScore = $currentScore
                        MaxScore = $latestScore.maxScore
                        LicensedUserCount = $latestScore.licensedUserCount
                        SimilarCompanyAverage = $similarCompanyAverage
                        IdentityScore = $identityScore
                        DataScore = $dataScore
                        DeviceScore = $deviceScore
                        AppsScore = $appsScore
                        InfrastructureScore = $infrastructureScore
                    } | Export-csv C:\temp\SecureScoreReports\AllTenantOverview.csv -NoTypeInformation -Append
                    Write-Host "Exported Secure Score Report" -ForegroundColor Green
                }
                   
                $Stoploop = $true
            }
            catch {
                if ($Retrycount -gt 5) {
                    Write-Host "Could not get secure score, or complete report after 6 retries." -ForegroundColor Red
                    $Stoploop = $true
                }
                else {
                    Write-Host "Could not get secure score, or complete report. Retrying in 5 seconds..." -ForegroundColor DarkYellow
                    Start-Sleep -Seconds 5
                    $Retrycount ++
                }
            }
        }
        While ($Stoploop -eq $false)
    }
    Remove-AzureADApplication -ObjectId $aadApplication.ObjectId
}
else {
    Write-Host "Microsoft Graph Service Principal could not be found or created" -ForegroundColor Red
}
