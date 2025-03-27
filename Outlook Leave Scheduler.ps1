$LeaveStartDay = Get-Date -Format "dd/MM/yyyy"
$LeaveEndDay = Get-Date -Format "dd/MM/yyyy"
$Subjectline = "Planned leave"
#==========================================================
$ol = New-Object -comObject Outlook.Application
$meeting = $ol.CreateItem(1) #('olAppointmentItem')

#$mSubject = Read-Host 'Enter meeting Subject line '
$meeting.Subject = $Subjectline
# $meeting.Location = 'Office Closed'
$meeting.ReminderSet = $false
$meeting.MeetingStatus = 1
$meeting.ResponseStatus = 0
$meeting.ResponseRequested = $false
$meeting.Start =  $LeaveStartDay 
$meeting.End =  $LeaveEndDay  
# $meeting.Duration = 540
#$datediff = $date2 - $date
$datediff = $meeting.End - $meeting.Start
#$meeting.Duration = [Math]::Round($datediff.TotalMinutes)  # written before 'if block' logic
if ([Math]::Round($datediff.TotalMinutes))
{
	$OneDay = New-TimeSpan -Days 1 #timespan https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/new-timespan?view=powershell-7.2
	$datediff = $datediff + $OneDay
	$meeting.Duration = [Math]::Round($datediff.TotalMinutes)
	
}
else
{
	$meeting.Duration = 1440
}
$meeting.BusyStatus = 0
$meeting.AllDayEvent = $true
$meeting.location = "NA"
#Write-Host "`r`nEnter leave category" -ForegroundColor Green 
#$meeting.Categories = 'New Yr Holiday' #sick oof, Holiday
#$meeting.Save()
$meeting.display()
#$meeting.Duration
#[Math]::Round($d.TotalMinutes)
