# .DESCRIPTION
#   This is a script to analyze the performance of MRS move requests.
#   It outputs some important performance statistics of a given set of move request statistics.
#   It also generates two files. One for the failure list, and the other for individual move stats.
#   For more information, please visit http://aka.ms/MailboxMigrationPerfScript
#
#   Copyright (c) Microsoft Corporation. All rights reserved.
#
#   THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#   OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

function ProcessStats([array] $stats, [string] $name, [int] $percentile=90)
{

if($stats.count -eq 0)
{
  return
}

$startTimestamp = ($stats | sort QueuedTimeStamp | select -first 1).QueuedTimeStamp
$lastCompleted = $stats | sort completiontimestamp -Descending | select -First 1
$lastTimestamp = $lastCompleted.completiontimestamp

if($lastTimestamp -eq $null)
{
	$lastSuspended = $stats | sort SuspendedTimestamp -Descending | select -First 1
	$lastTimestamp = $lastCompleted.SuspendedTimestamp
	if($lastTimestamp -eq $null)
	{
		$lastTimestamp = Get-Date 
	}
}

$moveDuration = $lastTimestamp - $startTimestamp
$MoveDurationInTicks = [math]::Truncate(($lastTimestamp - $startTimestamp).Ticks)

$perMoveInfo = $stats  | 
	select alias, TotalInProgressDuration, TotalIdleDuration, OverallDuration,TotalQueuedDuration,TotalDataReplicationWaitDuration,TotalSuspendedDuration, `
	@{Name="SourceProviderDuration"; Expression={$_.Report.SessionStatistics.SourceProviderInfo.TotalDuration + $_.Report.ArchiveSessionStatistics.SourceProviderInfo.TotalDuration}}, `
	@{Name="DestinationProviderDuration"; Expression={$_.Report.SessionStatistics.DestinationProviderInfo.TotalDuration + $_.Report.ArchiveSessionStatistics.DestinationProviderInfo.TotalDuration}}, `
	@{Name="RelinquishedDurationInTicks"; Expression={$_.OverallDuration.Ticks - $(($_.TotalInProgressDuration, $_.TotalQueuedDuration, $_.TotalFailedDuration, $_.TotalSuspendedDuration)|Measure-Object Ticks -Sum).Sum }}, `

	@{Name="TotalStalledDueToCIInTicks"; Expression={$_.TotalStalledDueToCIDuration.Ticks}},
	@{Name="TotalStalledDueToHAInTicks"; Expression={$_.TotalStalledDueToHADuration.Ticks}},
	@{Name="TotalStalledDueToTargetCpuInTicks"; Expression={$_.TotalStalledDueToWriteCpu.Ticks}},
	@{Name="TotalStalledDueToSourceCpuInTicks"; Expression={$_.TotalStalledDueToReadCpu.Ticks}},
	@{Name="TotalStalledDueToMailboxLockedDurationInTicks"; Expression={$_.TotalStalledDueToMailboxLockedDuration.Ticks}},
	@{Name="TotalStalledDueToSourceProxyUnknownInTicks"; Expression={$_.TotalStalledDueToReadUnknown.Ticks}},
	@{Name="TotalStalledDueToTargetProxyUnknownInTicks"; Expression={$_.TotalStalledDueToWriteUnknown.Ticks}},

	@{Name="SourceLatencySampleCount"; Expression={$_.Report.SessionStatistics.SourceLatencyInfo.NumberOfLatencySamplingCalls}},
	@{Name="AverageSourceLatency"; Expression={$_.Report.SessionStatistics.SourceLatencyInfo.Average}},
	@{Name="TotalNumberOfSourceSideRemoteCalls"; Expression={$_.Report.SessionStatistics.SourceLatencyInfo.TotalNumberOfRemoteCalls}},
	@{Name="DestinationLatencySampleCount"; Expression={$_.Report.SessionStatistics.DestinationLatencyInfo.NumberOfLatencySamplingCalls}},
	@{Name="AverageDestinationLatency"; Expression={$_.Report.SessionStatistics.DestinationLatencyInfo.Average}},
	@{Name="TotalNumberOfDestinationSideRemoteCalls"; Expression={$_.Report.SessionStatistics.DestinationLatencyInfo.TotalNumberOfRemoteCalls}},

	@{Name="WordBreaking_TotalTimeProcessingMessagesInTicks"; Expression={$_.Report.SessionStatistics.TotalTimeProcessingMessages.Ticks}},
	@{Name="TotalTransientFailureDurationInTicks"; Expression={$_.TotalTransientFailureDuration.Ticks}},
	@{Name="TotalInProgressDurationInTicks"; Expression={$_.TotalInProgressDuration.Ticks}},
	@{Name="MailboxSizeInMB"; Expression={ (ToMB $_.TotalMailboxSize) + (ToMB (GetArchiveSize -size $_.TotalArchiveSize -flags $_.Flags))}},
	@{Name="TransferredMailboxSizeInMB"; Expression={((ToMB $_.TotalMailboxSize) + (ToMB (GetArchiveSize -size $_.TotalArchiveSize -flags $_.Flags))) * $_.PercentComplete / 100}},
	@{Name="PerMoveRate"; Expression={(((ToKB $_.TotalMailboxSize) + (ToKB (GetArchiveSize -size $_.TotalArchiveSize -flags $_.Flags)))* $_.PercentComplete / 100 /$_.TotalInProgressDuration.TotalSeconds) * 3600 / 1024}},
    @{Name="BytesTransferredInMB"; Expression={ToMB $_.BytesTransferred}},
    @{Name="PerMoveTransferRate"; Expression={((ToKB $_.BytesTransferred) / $_.TotalInProgressDuration.TotalSeconds) * 3600 / 1024}}

$perMoveInfo = $perMoveInfo| sort permovetransferrate -desc
$perMoveInfo = @($perMoveInfo);
$perMoveInfo | Export-Csv "$($name)PerMoveInfo.csv"  -NoTypeInformation

$perMoveInfo = $perMoveInfo| select -first ($perMoveinfo.count * ($percentile/100))
$totalInProgressDurationInTicks = ($perMoveInfo  | measure -property TotalInProgressDurationInTicks -Sum).Sum
$TotalStalledDueToCIInTicks = ($perMoveInfo  | measure -property TotalStalledDueToCIInTicks -Sum).Sum
$TotalStalledDueToHAInTicks = ($perMoveInfo  | measure -property TotalStalledDueToHAInTicks -Sum).Sum

$MeasuredOverallDurationInTicks = ($perMoveInfo | select @{Name="OverallDurationInTicks"; Expression={$_.OverallDuration.Ticks}} | measure -property OverallDurationInTicks -Sum -Maximum -Minimum -Average)
$MeasuredIdleDurationInTicks = ($perMoveInfo | select @{Name="TotalIdleDurationInTicks"; Expression={$_.TotalIdleDuration.Ticks}} | measure -property TotalIdleDurationInTicks -Sum -Maximum -Minimum -Average)
$MeasuredSourceProviderDurationInTicks = ($perMoveInfo | select @{Name="SourceProviderDurationInTicks"; Expression={$_.SourceProviderDuration.Ticks}} | measure -property SourceProviderDurationInTicks -Sum -Maximum -Minimum -Average)
$MeasuredDestinationProviderDurationInTicks = ($perMoveInfo | select @{Name="DestinationProviderDurationInTicks"; Expression={$_.DestinationProviderDuration.Ticks}} | measure -property DestinationProviderDurationInTicks -Sum -Maximum -Minimum -Average)
$MeasuredRelinquishedDurationInTicks = ($perMoveInfo |  measure -property RelinquishedDurationInTicks -Sum -Maximum -Minimum -Average)

$TotalOverallDurationInTicks = $MeasuredOverallDurationInTicks.Sum
$TotalIdleDurationInTicks = $MeasuredIdleDurationInTicks.Sum
$TotalSourceProviderDurationInTicks = $MeasuredSourceProviderDurationInTicks.Sum
$TotalDestinationProviderDurationInTicks = $MeasuredDestinationProviderDurationInTicks.Sum

$IdlePercent = $("{0:p2}" -f ((Nullify $TotalIdleDurationInTicks) / $TotalInProgressDurationInTicks))
$SourcePercent = $("{0:p2}" -f ((Nullify $TotalSourceProviderDurationInTicks) / $TotalInProgressDurationInTicks))
$DestinationPercent = $("{0:p2}" -f ((Nullify $TotalDestinationProviderDurationInTicks) / $TotalInProgressDurationInTicks))

$TotalStalledDueToTargetCpuInTicks = ($perMoveInfo  | measure -property TotalStalledDueToTargetCpuInTicks -Sum).Sum
$TotalStalledDueToSourceCpuInTicks = ($perMoveInfo  | measure -property TotalStalledDueToSourceCpuInTicks -Sum).Sum
$TotalStalledDueToMailboxLockedDurationInTicks = ($perMoveInfo  | measure -property TotalStalledDueToMailboxLockedDurationInTicks -Sum).Sum
$TotalStalledDueToSourceProxyUnknownInTicks = ($perMoveInfo  | measure -property TotalStalledDueToSourceProxyUnknownInTicks -Sum).Sum
$TotalStalledDueToTargetProxyUnknownInTicks = ($perMoveInfo  | measure -property TotalStalledDueToTargetProxyUnknownInTicks -Sum).Sum

$WordBreaking_TotalTimeProcessingMessagesInTicks = ($perMoveInfo  | measure -property WordBreaking_TotalTimeProcessingMessagesInTicks -Sum).Sum

$CIStallPercent = $($TotalStalledDueToCIInTicks/$totalInProgressDurationInTicks)
$HAStallPercent = $($TotalStalledDueToHAInTicks/$totalInProgressDurationInTicks)
$SourceCPUStallPercent = $($TotalStalledDueToSourceCpuInTicks/$totalInProgressDurationInTicks)
$TargetCPUStallPercent = $($TotalStalledDueToTargetCpuInTicks/$totalInProgressDurationInTicks)
$MailboxLockedStallPercent = $($TotalStalledDueToMailboxLockedDurationInTicks/$totalInProgressDurationInTicks)
$ProxyUnknownStallPercent = $(($TotalStalledDueToSourceProxyUnknownInTicks + $TotalStalledDueToTargetProxyUnknownInTicks)/$totalInProgressDurationInTicks)

$totalStalledTimeInTicks = $TotalStalledDueToCIInTicks + $TotalStalledDueToHAInTicks + $TotalStalledDueToSourceCpuInTicks + $TotalStalledDueToTargetCpuInTicks +  $TotalStalledDueToMailboxLockedDurationInTicks + $TotalStalledDueToTargetProxyUnknownInTicks + $TotalStalledDueToSourceProxyUnknownInTicks
$TotalTransientFailureDurationInTicks = ($perMoveInfo  | measure -property TotalTransientFailureDurationInTicks -Sum).Sum
$TransientFailurePercent = $($TotalTransientFailureDurationInTicks/$totalInProgressDurationInTicks)
$delayRatio = $($totalStalledTimeInTicks/$totalInProgressDurationInTicks)
$totalMailboxSizeInMB = ($perMoveInfo  | measure -property MailboxSizeInMB -Sum).Sum
$totalTransferredMailboxSizeInMB = ($perMoveInfo  | measure -property TransferredMailboxSizeInMB -Sum).Sum
$MeasuredPerMoveTransferRate = ($perMoveInfo  | measure -property PerMoveTransferRate -Average -Maximum -Minimum) 
$totalMegabytesTransferred = ($perMoveInfo  | measure -property BytesTransferredInMB -Sum).Sum 
$perMoveRateInMBPerHour = ($perMoveInfo | measure -Property PerMoveRate -average).Average

$averageSourceLatency = ($perMoveInfo | ? {$_.SourceLatencySampleCount -gt 0} | measure -Property AverageSourceLatency -average).Average
$averageNumberOfSourceSideRemoteCalls = ($perMoveInfo | measure -Property TotalNumberOfSourceSideRemoteCalls -average).Average
$averageDestinationLatency = ($perMoveInfo | ? {$_.DestinationLatencySampleCount -gt 0} | measure -Property AverageDestinationLatency -average).Average
$averageNumberOfDestinationSideRemoteCalls = ($perMoveInfo | measure -Property TotalNumberOfDestinationSideRemoteCalls -average).Average
$WordBreakingVsInProgressRatio  = $("{0:p2}" -f $($WordBreaking_TotalTimeProcessingMessagesInTicks/$totalInProgressDurationInTicks))

$mailboxCount = $batch.Count

$nl = [System.Environment]::NewLine

$failures = ""

$stats | % {  if($_.Report.Failures -ne $null) { $failures += $_.Alias + ": " + $_.Report.Failures + $nl}}

if($failures.Length -gt 0){$failures | Add-Content -Path "$($name)failures.txt" -Encoding ASCII ;   write-warning "Move reports contains failures. Check $($name)failures.txt"}

return New-Rec -name $name -mailboxCount $stats.Count -moveDuration $(GetTimeSpan($MoveDurationInTicks)) -startTime $($startTime.QueuedTimeStamp) -completionTime $lastTimestamp `
   -TotalMailboxSizeInGB $(RoundIt($($totalMailboxSizeInMB / 1024))) -TotalTransferredMailboxSizeInGB $(RoundIt($($totalTransferredMailboxSizeInMB / 1024))) `
   -TotalThroughputGBPerHour $(RoundIt($($totalTransferredMailboxSizeInMB / $MoveDuration.TotalHours / 1024))) -PerMoveThroughputGBPerHour $(RoundIt($perMoveRateInMBPerHour / 1024)) `
   -StalledVsInProgressRatio $("{0:p2}" -f $delayRatio) -WordBreakingVsInProgressRatio $WordBreakingVsInProgressRatio  `
   -CIStallVsInProgressRatio $("{0:p2}" -f $CIStallPercent) -HAStallVsInProgressRatio $("{0:p2}" -f $HAStallPercent) `
   -TargetCPUStallVsInProgressRatio $("{0:p2}" -f $TargetCPUStallPercent ) -SourceCPUStallVsInProgressRatio $("{0:p2}" -f $SourceCPUStallPercent ) `
   -MailboxLockedStallVsInProgressRatio $("{0:p2}" -f $MailboxLockedStallPercent ) -ProxyUnknownStallVsInProgressRatio $("{0:p2}" -f $ProxyUnknownStallPercent  ) `
   -TransientFailurePercent $("{0:p2}" -f $TransientFailurePercent) `
   -IdlePercent $IdlePercent -SourcePercent $SourcePercent -DestinationPercent $DestinationPercent `
   -MeasuredOverallDuration $MeasuredOverallDurationInTicks -MeasuredIdleDuration $MeasuredIdleDurationInTicks -MeasuredSourceProviderDuration $MeasuredSourceProviderDurationInTicks -MeasuredDestinationProviderDuration $MeasuredDestinationProviderDurationInTicks `
   -MeasuredRelinquishedDuration $MeasuredRelinquishedDurationInTicks `
   -TotalGBTransferred $(RoundIt($($totalMegabytesTransferred/1024))) -MeasuredPerMoveTransferRate $MeasuredPerMoveTransferRate `
   -AverageSourceLatency $(RoundIt($($averageSourceLatency))) -AverageNumberOfSourceSideRemoteCalls $(RoundIt($($averageNumberOfSourceSideRemoteCalls))) `
   -AverageDestinationLatency $(RoundIt($($averageDestinationLatency))) -AverageNumberOfDestinationSideRemoteCalls $(RoundIt($($averageNumberOfDestinationSideRemoteCalls))) `
}

function New-Rec()
{
param ([string]$name, [int]$MailboxCount, $MoveDuration, $StartTime, $CompletionTime, 
$TotalMailboxSizeInGB, $TotalTransferredMailboxSizeInGB, $TotalThroughputGBPerHour,$PerMoveThroughputGBPerHour,
$StalledVsInProgressRatio,$WordBreakingVsInProgressRatio,
$CIStallVsInProgressRatio, $HAStallVsInProgressRatio,$TargetCPUStallVsInProgressRatio,$SourceCPUStallVsInProgressRatio,$MailboxLockedStallVsInProgressRatio, $ProxyUnknownStallVsInProgressRatio,
$TransientFailurePercent, $IdlePercent, $SourcePercent, $DestinationPercent, $MeasuredOverallDuration, $MeasuredIdleDuration, $MeasuredSourceProviderDuration, $MeasuredDestinationProviderDuration, $MeasuredRelinquishedDuration, $TotalGBTransferred, $MeasuredPerMoveTransferRate,
$AverageSourceLatency, $AverageNumberOfSourceSideRemoteCalls, $AverageDestinationLatency, $AverageNumberOfDestinationSideRemoteCalls)

 $rec = new-object PSObject

  $rec | add-member -type NoteProperty -Name Name -Value $Name
  $rec | add-member -type NoteProperty -Name StartTime -Value $StartTime 
  $rec | add-member -type NoteProperty -Name EndTime -Value $CompletionTime
  $rec | add-member -type NoteProperty -Name MigrationDuration -Value $MoveDuration

  $rec | add-member -type NoteProperty -Name MailboxCount -Value $MailboxCount
  $rec | add-member -type NoteProperty -Name TotalGBTransferred -Value $TotalGBTransferred
  $rec | add-member -type NoteProperty -Name PercentComplete -Value $(RoundIt($TotalTransferredMailboxSizeInGB / $TotalMailboxSizeInGB * 100))
  
  $rec | add-member -type NoteProperty -Name MaxPerMoveTransferRateGBPerHour -Value $(RoundIt($MeasuredPerMoveTransferRate.Maximum / 1024))
  $rec | add-member -type NoteProperty -Name MinPerMoveTransferRateGBPerHour -Value $(RoundIt($MeasuredPerMoveTransferRate.Minimum / 1024))
  $rec | Add-Member -Type NoteProperty -Name AvgPerMoveTransferRateGBPerHour -Value $(RoundIt($MeasuredPerMoveTransferRate.Average / 1024))

  #transfer size is greater than the source mailbox size due to transient failures and other factors. This shows how close these numbers are.
  $rec | add-member -type NoteProperty -Name MoveEfficiencyPercent -Value $(RoundIt($TotalTransferredMailboxSizeInGB / $TotalGBTransferred * 100))
  
  $rec | add-member -type NoteProperty -Name AverageSourceLatency -Value $AverageSourceLatency #applies to onboarding
  $rec | add-member -type NoteProperty -Name AverageDestinationLatency -Value $AverageDestinationLatency #applies to offboarding
  
  $rec | add-member -type NoteProperty -Name IdleDuration -Value $IdlePercent

  $rec | add-member -type NoteProperty -Name SourceSideDuration -Value $SourcePercent
  $rec | add-member -type NoteProperty -Name DestinationSideDuration -Value $DestinationPercent

  $rec | add-member -type NoteProperty -Name WordBreakingDuration -Value $WordBreakingVsInProgressRatio
  $rec | add-member -type NoteProperty -Name TransientFailureDurations -Value $TransientFailurePercent

  $rec | add-member -type NoteProperty -Name OverallStallDurations -Value $StalledVsInProgressRatio
  $rec | add-member -type NoteProperty -Name ContentIndexingStalls -Value $CIStallVsInProgressRatio
  $rec | add-member -type NoteProperty -Name HighAvailabilityStalls -Value $HAStallVsInProgressRatio
  $rec | add-member -type NoteProperty -Name TargetCPUStalls -Value $TargetCPUStallVsInProgressRatio
  $rec | add-member -type NoteProperty -Name SourceCPUStalls -Value $SourceCPUStallVsInProgressRatio
  $rec | add-member -type NoteProperty -Name MailboxLockedStall -Value $MailboxLockedStallVsInProgressRatio
  $rec | add-member -type NoteProperty -Name ProxyUnknownStall -Value $ProxyUnknownStallVsInProgressRatio
  
  return $rec
}

#utility functions

function LogIt($str)
{
   $currentTime = Get-Date -Format "hh:mm:ss"
   $loggedText = "[{0}] {1}" -f $currentTime,$str
   write-host $loggedText
}

function GetTimeSpan($ticks)
{
  if($seconds -eq 0)
  {
    return "0"
  }
  $a = [TimeSpan]::FromTicks($ticks)
  if($a.Days -eq 0)
  {
    return "{0:00}:{1:00}:{2:00}" -f $a.hours,$a.minutes,$a.seconds
  }
  else
  {
    return "{0} day(s) {1:00}:{2:00}:{3:00}" -f $a.days,$a.hours,$a.minutes,$a.seconds
  }
}

function RoundIt($num)
{
  return "{0:N2}" -f $num
}


 function Nullify($var)
 {
   if($var -eq $null)
   {
	 return 0
   }
   else
   {
     return $var
   }
 }

function ByteStrToBytes($str)
{
   if($str -eq $null)
   {
     return 0;
   }

   $str = $str.ToString()
   return [int64]$str.substring($str.IndexOf('(') + 1, $str.IndexOf(' bytes)')-$str.IndexOf('(')-1)
}

function ToMB($str)
{
	return (ByteStrToBytes $str)/1024/1024
}

function ToKB($str)
{
	return (ByteStrToBytes $str)/1024
}

function GetArchiveSize($size, $flags)
{
	if($flags.ToString().Contains("MoveOnlyArchiveMailbox"))
    {
        return $null
    }
    
    return $size;
}