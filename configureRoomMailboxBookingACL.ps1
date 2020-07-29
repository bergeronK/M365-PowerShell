


#.Reconf meeting rooms allowed users ACL  
  $meetingRoomName = "UK-Lon-Ton-14-Training"  
 
#.This group should be a mail enabled group 
# recipientType : MailUniversalSecurityGroup 
  $BookInPolicy = "Trainers@ABC.onmicrosoft.com" 
 
  Set-CalendarProcessing -Identity $meetingRoomName -AllRequestInPolicy $false   
  Set-CalendarProcessing -Identity $meetingRoomName -BookInPolicy $BookInPolicy 
 


 