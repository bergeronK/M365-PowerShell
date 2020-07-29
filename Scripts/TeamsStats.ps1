<#
|------------------------------------------------------------ Microsoft Teams Report ------------------
| Written by Shaun Wilkinson                                                                                                                        
| Generates a report showing all teams and users                                                                                                    
| Requirements: Access to all teams and Team PS module - https://www.powershellgallery.com/packages/MicrosoftTeams                                  
|------------------------------------------------------------------------------------------------------
#>


$exportLocation = ".\exportTEams.csv"


Connect-MicrosoftTeams

# Get all of the team Groups IDs
$AllTeamsInOrg = (Get-Team).GroupID
$TeamReport = @()

# Will hold a basic count of user types and teams
$unavailableTeamCount = 0
$knownOwnersCount = 0
$knownMemberCount = 0
$knownGuestCount = 0

# Loop through all Group IDs
$currentIndex = 1
ForEach($Team in $AllTeamsInOrg) {
    # Show a nice progress bar as this can take a while
    Write-Progress -Id 0 -Activity "Building report from Microsoft Teams" -Status "$currentIndex of $($allTeamsInOrg.Count)" -PercentComplete (($currentIndex / $allTeamsInOrg.Count) * 100)

    # Get properties of the team
    $team = Get-Team -GroupId $Team

    # Attempt to get team users, throw error message if no access
    try {
        # Get team members
        $users = Get-TeamUser -GroupId $team.groupID

        # foreach user create a line in the report
        ForEach($user in $users) {
            # Maintain a count of user types
            switch($user.Role) {
                "owner" { $knownOwnersCount++ }
                "member" { $knownMemberCount++ }
                "guest" { $knownGuestCount++ }
            }

            # Create an object to hold all values
            $teamReportObject = New-Object PSObject -Property @{
                TeamName = $team.DisplayName
                Description = $team.Description
                Archived = $team.Archived
                Visibility = $team.Visibility
                User = $user.Name
                Email = $user.User
                Role = $user.Role
            }

            # Add to the report
            $TeamReport += $teamReportObject
        }
    } catch [Microsoft.TeamsCmdlets.PowerShell.Custom.ErrorHandling.ApiException] {
        Write-Host -ForegroundColor Yellow "No access to $($team.DisplayName) team, cannot generate report"
        $unavailableTeamCount++
    }

    
    $currentIndex++
}
Write-Progress -Id 0 -Activity " " -Status " " -Completed

# Disconnect from the teams service
Disconnect-MicrosoftTeams

# Provide some nice output
Write-Host -ForegroundColor Green "============================================================"
Write-Host -ForegroundColor Green "                Microsoft Teams User Report                 "
Write-Host -ForegroundColor Green ""
Write-Host -ForegroundColor Green "  Count of All Teams - $($AllTeamsInOrg.Count)                "
Write-Host -ForegroundColor Green "  Count of Inaccesible Teams - $($unavailableTeamCount)         "
Write-Host -ForegroundColor Green ""
Write-Host -ForegroundColor Green "  Count of Known Users - $($AllTeamsInOrg.Count)                "
Write-Host -ForegroundColor Green "  Count of Known Owners - $($knownOwnersCount)                "
Write-Host -ForegroundColor Green "  Count of Known Members - $($knownMemberCount)                "
Write-Host -ForegroundColor Green "  Count of Known Guests - $($knownGuestCount)                "


$TeamReport | Export-CSV $exportLocation -NoTypeInformation
Write-Host -ForegroundColor Green "Exported report to $($exportLocation)"
