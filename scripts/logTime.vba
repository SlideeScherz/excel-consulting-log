Sub clockIn()

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

  Set timeDate = Range("C3")
  Set timeWeekDay = Range("C4")
  Set timeStart = Range("C5")
  Set timeEnd = Range("C6")
  
  ' if date is available
  Dim proceed As Boolean
  proceed = True
  
  MsgBox "Log time?"

  timeDate.Value = dateNow
  timeWeekDay.Value = weekdayNow

  ' Test if the value is cell is blank/empty, and mark time for this correct slot
  If IsEmpty(timeStart) = True And proceed = True Then
    MsgBox "Marking time start 1"
    timeStart.Value = timeNow
  ElseIf IsEmpty(timeEnd) = True And proceed = True Then
    MsgBox "Marking time end 1"
    timeEnd.Value = timeNow
  Else
    MsgBox "Export this data before logging more time."
  End If

End Sub
