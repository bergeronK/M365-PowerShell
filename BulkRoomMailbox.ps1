#Script creates roombailboxes from csv
#http://thatlazyadmin.com
#twitter:@ShaunHardneck
###################################################################################################
Read-Host "Nerd Life"
###################################################################################################
Import-Csv .\BulkRoomMailbox.csv|%{New-Mailbox -Name $_.Name -DisplayName $_.Name -Database $_.Database -ResourceCapacity $_.Capacity-UserPrincipalName $_.UPN -Password (ConvertTo-SecureString $_.Password -AsPlainText -force) -Room } | %{Set-CalendarProcessing $_.Name -AutomateProcessing AutoAccept -AllowConflicts $false -BookingWindowInDays 10 -AddOrganizerToSubject $true -ResourceDelegates $_.Delegate}
