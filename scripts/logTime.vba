Option Explicit

Sub logTime()

  ' confirm execution
  Dim feedback
  feedback = MsgBox("Log time?", vbYesNo + vbQuestion, "Proceed?")

  If feedback <> vbYes Then 
    Exit Sub
  End If

  ' init and assign date time objects
  Dim dateNow, timeNow, weekdayIndex, weekdayNow
  dateNow = Date
  timeNow = Now

  ' set the weekday
  weekdayIndex = Weekday(dateNow)
  weekdayNow = WeekdayName(weekdayIndex, False)

  Set timeDate = Range("A3")
  Set timeWeekDay = Range("B3")
  Set timeStart = Range("C3")
  Set timeEnd = Range("D3")
      
  timeDate.Value = dateNow
  timeWeekDay.Value = weekdayNow

  ' Test if the value is cell is blank/empty, and mark time for this correct slot
  If IsEmpty(timeStart) = True Then
    timeStart.Value = timeNow
  ElseIf IsEmpty(timeEnd) = True Then
    timeEnd.Value = timeNow
  Else
    MsgBox "Export this data before logging more time."
  End If

End Sub

