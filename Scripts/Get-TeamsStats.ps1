#generate array for custom object
$teamsinfo = @()
#get all teams from organisation
$teams = get-team 
#find members, owner, guest
foreach($team in $teams){
  $displayname = ($team.DisplayName)
  $Description = ($team.Description)
  $groupid = $team.groupid
  $members = (Get-TeamUser -GroupId $groupid -Role Member).User
  $owner = (Get-TeamUser -GroupId $groupid -Role Owner).User
  $guests = (Get-TeamUser -GroupId $groupid -Role Guest).User
   #custom object for output
  $teamsinfo += [pscustomobject]@{
    DisplayName   = $displayname
    Owner = ("$owner")
    Members = ("$members")
    Guests = ("$guests")
    Description = ("$Description")
  }
}
#show teaminformation in OutGrid-View
$teamsinfo | Sort-Object DisplayName | Out-GridView -Title "All Office365 Groups created in MS Teams"
$teamsinfo | Sort-Object DisplayName | export-csv .\TeamsExp.csv -NoTypeInformation