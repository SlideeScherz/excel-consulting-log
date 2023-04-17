Option Explicit

Sub logTime()

  ' confirm execution
  Dim feedback
  feedback = MsgBox("Log time?", vbYesNo + vbQuestion, "Confirm")

  If feedback <> vbYes Then 
    Exit Sub
  End If

  ' init constants
  globals.init

  ' init and assign date time objects
  Dim dateNow, timeNow, weekdayIndex, weekdayNow 
  ' for goals and accomplished notes
  Dim sessionNotes as String
  dateNow = Date
  timeNow = Now

  ' set the weekday
  weekdayIndex = Weekday(dateNow)
  weekdayNow = WeekdayName(weekdayIndex, False)
      
  If IsEmpty(timeDate) = True Then
    timeDate.Value = dateNow
  End If

  If IsEmpty(timeWeekDay) = True Then
    timeWeekDay.Value = weekdayNow
  End If
  
  ' Test if the value is cell is blank/empty, and mark time for this correct slot
  If IsEmpty(timeStart) = True Then
    timeStart.Value = timeNow
    timeStampButton.Caption = "Clock Out"
    'createEntryButton.Enabled = False
    sessionNotes = Application.InputBox(prompt := "Goals", type := 2)
    goals.Value = sessionNotes
  ElseIf IsEmpty(timeEnd) = True Then
    timeEnd.Value = timeNow
    'createEntryButton.Enabled = True
    sessionNotes = Application.InputBox(prompt := "Accomplished", type := 2)
    accomplished.Value = sessionNotes
  Else
    MsgBox "Export this data before logging more time."
  End If

End Sub

