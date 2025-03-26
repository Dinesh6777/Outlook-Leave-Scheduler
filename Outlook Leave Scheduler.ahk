; This program is free software: you can redistribute it and/or modify
; it under the terms of the GNU General Public License as published by
; the Free Software Foundation, either version 3 of the License, or
; (at your option) any later version.

; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
; GNU General Public License for more details.

; You should have received a copy of the GNU General Public License
; along with this program. If not, see <https://www.gnu.org/licenses/>.

; Script Name: Outlook Leave Scheduler
; Author: Dinesh Kadali
; Description: Outlook Leave Scheduler will help send “out of office” invite without blocking other’s calendars.


#NoEnv
#SingleInstance force

SetTimer, Message, 36000
Message:
    progress, r0-30,  , Loading Outlook - Please wait...
    loop 20 {
        ;progress, % a_index, % 30 - a_index
        sleep 100
    }
    progress, hide
;return

    ;################################# FantasticGuru https://autohotkey.com//boards/viewtopic.php?f=5&t=39576
    ;MsgBox, beginning of script

    olApp := ComObjCreate("Outlook.Application")
    meeting := olApp.CreateItem(1)	; olAppointmentItem := 1
	meeting.ReminderSet := false
	meeting.MeetingStatus := 1
	meeting.ResponseStatus := 0
	meeting.ResponseRequested := false
    meeting.Subject := "Planned OOF#"
	meeting.BusyStatus := 0
	meeting.AllDayEvent := true
	meeting.location := "NA"
    meeting.Recipients.Add("aivpindia@microsoft.com")
    meeting.Display


	;olAppoint.Location := "Online"
    ;olAppoint.Start := "1/1/2021 8:32:00 AM"
    ;olAppoint.Duration := 45
    ;olAppoint.Body := "Don't forget to breath"
    ;olAppoint.Display ; Remove this line to have it all happen in the background
    ;~ olAppoint.Save ; Uncomment to automatically save
    ;MsgBox, end of Appointment script


    ;Esc:: ExitApp

; Date picker - https://tdalon.blogspot.com/2020/09/autohotkey-insert-date.html
/*
Gui, -Caption +AlwaysOnTop
;Gui SetTimer, Add, 1000
Gui, Font, s50 w700 q4, Arial
Gui, Color, White
Gui, Margin, 10, 5
Gui, Add, Text, Center, Loading Outlook, please wait.
Gui, Show, NA
*/



/*
SetTimer, Message, 3600000
Message:
    progress, r0-300, 300, Take A Break`, Do it!
    loop 300 {
        progress, % a_index, % 300 - a_index
        sleep 1000
    }
    progress, hide
return
*/
