Import-Csv .\ConfRooms.csv|%{New-Mailbox -Name $_.DisplayName -DisplayName $_.DisplayName -UserPrincipalName $_.Email -Password (ConvertTo-SecureString $_.Password -AsPlainText -force) -Room -OrganizationUnit $_.OU } | %{Set-CalendarProcessing $_.DisplayName -AllowConflicts $false -AddOrganizerToSubject $true}


New-DistributionGroup -Name $_.roomList -OrganizationalUnit [OU] -RoomList
Add-DistributionGroupMember -Identity $_.roomList -Member $_.roomName
