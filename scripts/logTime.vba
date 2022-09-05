Sub logTime()

  ' init and assign date time objects
  Dim dateNow, timeNow, weekdayIndex, weekdayNow
  dateNow = Date
  timeNow = Now

  ' set the weekday
  weekdayIndex = Weekday(dateNow)
  weekdayNow = WeekdayName(weekdayIndex, False)

  Dim timeDate As Range
  Dim timeWeekDay As Range
  Dim timeStart As Range
  Dim timeEnd As Range

  Set timeDate = Range("A3")
  Set timeWeekDay = Range("B3")
  Set timeStart = Range("C3")
  Set timeEnd = Range("D3")
  
  ' confirm execution
  Dim feedback
  feedback = MsgBox("Log time?", vbYesNo + vbQuestion, "Proceed?")
    
  timeDate.Value = dateNow
  timeWeekDay.Value = weekdayNow

  ' Test if the value is cell is blank/empty, and mark time for this correct slot
  If IsEmpty(timeStart) = True And feedback = vbYes Then
    timeStart.Value = timeNow
  ElseIf IsEmpty(timeEnd) = True And feedback = vbYes Then
    timeEnd.Value = timeNow
  Else
    MsgBox "Export this data before logging more time."
  End If

End Sub
