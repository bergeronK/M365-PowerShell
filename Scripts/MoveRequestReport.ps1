[CmdletBinding()]
param (
    [bool]$CompletedMoves = $false, 
    [bool]$Synced = $false,    
    [bool]$Inprogress = $false,
    [bool]$AllMoves = $true,
    [bool]$IncludeAllBadItems = $true,
    [bool]$MoveDataBasic = $true,
    [bool]$MoveDataFull = $true,
    [bool]$GenereExportInfo = $true
)

#region Functions
function generateBadItemInformation ($BadItem,$User)
{
    $ExportEntry = New-Object PSObject
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Identity -Value $User.Alias
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name BatchName -Value $User.Batchname
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name DisplayName -Value $User.DisplayName
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Date -Value $BadItem.Date
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Subject -Value $BadItem.Subject
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Kind -Value $BadItem.Kind
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name FolderName -Value $BadItem.FolderName
    
    
    $ExportEntry   
}

function generateBasicMoveInformation ($User)
{
    $ExportEntry = New-Object PSObject
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Identity -Value $User.Alias
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name BatchName -Value $User.Batchname
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name DisplayName -Value $User.DisplayName
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Status -Value $User.Status
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name SyncStage -Value $User.SyncStage
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RemoteDatabase -Value $User.RemoteDatabase
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RemoteHostName -Value $User.RemoteHostName
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Message -Value $User.Message
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RecipientTypeDetails -Value $User.RecipientTypeDetails
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalMailboxSize -Value $User.TotalMailboxSize
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalMailboxItemCount -Value $User.TotalMailboxItemCount
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalArchiveSize -Value $User.TotalArchiveSize
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalArchiveItemCount -Value $User.TotalArchiveItemCount
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalPrimarySize -Value $User.TotalPrimarySize
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalPrimaryItemCount -Value $User.TotalPrimaryItemCount
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name BytesTransferred -Value $User.BytesTransferred

    $ExportEntry   
}

function generateFullMoveInformation ($User)
{
    $ExportEntry = New-Object PSObject
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Identity -Value $User.Alias
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name BatchName -Value $User.Batchname
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name DisplayName -Value $User.DisplayName
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Status -Value $User.Status
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name SyncStage -Value $User.SyncStage
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RemoteDatabase -Value $User.RemoteDatabase
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RemoteHostName -Value $User.RemoteHostName
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Message -Value $User.Message
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RecipientTypeDetails -Value $User.RecipientTypeDetails
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalMailboxSize -Value $User.TotalMailboxSize
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalMailboxItemCount -Value $User.TotalMailboxItemCount
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalArchiveSize -Value $User.TotalArchiveSize
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalArchiveItemCount -Value $User.TotalArchiveItemCount
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalPrimarySize -Value $User.TotalPrimarySize
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalPrimaryItemCount -Value $User.TotalPrimaryItemCount
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name BytesTransferred -Value $User.BytesTransferred
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name OverallDuration -Value $User.OverallDuration
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalSuspendedDuration -Value $User.TotalSuspendedDuration
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalFailedDuration -Value $User.TotalFailedDuration
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalQueuedDuration -Value $User.TotalQueuedDuration
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalInProgressDuration -Value $User.TotalInProgressDuration
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name StatusDetail -Value $User.StatusDetail
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name SourceVersion -Value $User.SourceVersion
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name FailureCode -Value $User.FailureCode
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name FailureType -Value $User.FailureType
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name FailureSide -Value $User.FailureSide
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name QueuedTimestamp -Value $User.QueuedTimestamp
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name StartTimestamp -Value $User.StartTimestamp
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name LastUpdateTimestamp -Value $User.LastUpdateTimestamp
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name LastSuccessfulSyncTimestamp -Value $User.LastSuccessfulSyncTimestamp
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name InitialSeedingCompletedTimestamp -Value $User.InitialSeedingCompletedTimestamp
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToContentIndexingDuration -Value $User.TotalStalledDueToContentIndexingDuration
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToMdbReplicationDuration -Value $User.TotalStalledDueToMdbReplicationDuration
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToMailboxLockedDuration -Value $User.TotalStalledDueToMailboxLockedDuration
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToReadThrottle -Value $User.TotalStalledDueToReadThrottle
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToWriteThrottle -Value $User.TotalStalledDueToWriteThrottle
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToReadCpu -Value $User.TotalStalledDueToReadCpu
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToWriteCpu -Value $User.TotalStalledDueToWriteCpu
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToReadUnknown -Value $User.TotalStalledDueToReadUnknown
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalStalledDueToWriteUnknown -Value $User.TotalStalledDueToWriteUnknown
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalTransientFailureDuration -Value $User.TotalTransientFailureDuration

    
    $ExportEntry   
}

function GenerateExcelInformation ($BadItems,$User)
{
    $ExportEntry = New-Object PSObject
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Identity -Value $User.Alias
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name BatchName -Value $User.Batchname
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name DisplayName -Value $User.DisplayName
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Status -Value $User.Status
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name SyncStage -Value $User.SyncStage
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RemoteDatabase -Value $User.RemoteDatabase
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RemoteHostName -Value $User.RemoteHostName
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name Message -Value $User.Message
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name RecipientTypeDetails -Value $User.RecipientTypeDetails
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalMailboxSize -Value $User.TotalMailboxSize
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name TotalMailboxItemCount -Value $User.TotalMailboxItemCount
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name BadItems -Value $BadItems.Count
    $ExportEntry | Add-Member -MemberType "NoteProperty" -Name BytesTransferred -Value $User.BytesTransferred
    
    $ExportEntry   
}


#endregion

#region Global Vales
$LogPath = "C:\Scripts\"
$FileTimeStamp = (Get-Date -Format yyyyMMdd-HHmmss).ToString()
$ExportBadItems = @()
$ExportBasicInformation = @()
$ExportFullInformation = @()
$ExportExcelInfo = @()
$Moves = @()
#endregion


#region Collect Movedata

If($AllMoves){
    $Moves = Get-MoveRequest | Get-MoveRequestStatistics -IncludeReport
}

If($CompletedMoves){
    $Moves = Get-MoveRequest -MoveStatus Completed | Get-MoveRequestStatistics
}

If($Synced){
    $Moves += Get-MoveRequest -MoveStatus Synced | Get-MoveRequestStatistics
}

If($Inprogress){
    $Moves += Get-MoveRequest -MoveStatus Inprogress | Get-MoveRequestStatistics
}

#endregion

#region Process moves Data

If($IncludeAllBadItems){

    $Moves2 = get-migrationuser | where{$_.SkippedItemCount -gt 0} | Get-MoveRequestStatistics -IncludeReport
    Foreach ($Move in $Moves2){
        $Baditems = ""
        $Baditems = ($Move.Report).Baditems
          
            If($Baditems -eq ""){
                Write-Host "No Errors found for " -BackgroundColor Green
            }
                Else{
                    Foreach($BadItem in $Baditems){
                        $ExportBadItems += generateBadItemInformation -BadItem $Baditem -User $Move
                    }
                }
    }

}

If($GenereExportInfo){

    Foreach ($Move in $Moves){
        $Baditems = ""
        $Baditems = ($Move.Report).Baditems
          
            If($Baditems -eq ""){
                Write-Host "No Errors found for " -BackgroundColor Green
            }
            Else{
                $ExportExcelInfo += GenerateExcelInformation -BadItems $Baditems -User $Move
            }
    }

}



If($MoveDataBasic){

     Foreach ($Move in $Moves){
        $ExportBasicInformation += generateBasicMoveInformation -User $Move
     }           
}


If($MoveDataFull){

    Foreach ($Move in $Moves){
        $ExportFullInformation += generateFullMoveInformation -User $Move           
    }
}

#endregion

#region Export data
If($IncludeAllBadItems){
    $ExportBadItems | ogv
    $FileName = $LogPath+"BadItems"+$FileTimeStamp+".csv"
    $ExportBadItems | Export-Csv -Path $FileName -NoTypeInformation
}

If($MoveDataBasic){
    $ExportBasicInformation | ogv
    $FileName = $LogPath+"BasicMoveData"+$FileTimeStamp+".csv"
    $ExportBasicInformation |Export-Csv -Path $FileName -NoTypeInformation
}

If($MoveDataFull){
    $ExportFullInformation | ogv
    $FileName = $LogPath+"FullMoveData"+$FileTimeStamp+".csv"
    $ExportFullInformation | Export-Csv -Path $FileName -NoTypeInformation
}


if($ExportExcelInfo){
    $ExportExcelInfo | ogv 
    $FileName = $LogPath+"ExcelData"+$FileTimeStamp+".csv"
    $ExportExcelInfo | Export-Csv -Path $FileName -NoTypeInformation
}

#endregion

#-Delimiter ";"
