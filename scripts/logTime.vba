Sub clockIn()

  ' init and assign date time objects
  Dim dateNow
  Dim timeNow
  dateNow = Date
  timeNow = Now

  Dim timeStart1 As Range
  Dim timeEnd1 As Range
  Dim timeStart2 As Range
  Dim timeEnd2 As Range

  Set timeStart1 = Range("B5:H5")
  Set timeEnd1 = Range("B6:H6")
  Set timeStart2 = Range("B7:H7")
  Set timeEnd2 = Range("B8:H8")

  ' possible feature
  'Dim columnDate As Range
  'Set columnDate = Range("B3:H3")

  Dim columnIndex As Integer
  columnIndex = 1
  
  ' Assign correct range based on day
  If Range("B3") = dateNow Then
    MsgBox "Log time for Sunday?"
    columnIndex = 1
  ElseIf Range("C3") = dateNow Then
    MsgBox "Log time for Monday?"
    columnIndex = 2
  ElseIf Range("D3") = dateNow Then
    MsgBox "Log time for Tuesday?"
    columnIndex = 3
  ElseIf Range("E3") = dateNow Then
    MsgBox "Log time for Wednesday"
    columnIndex = 4
  ElseIf Range("F3") = dateNow Then
    MsgBox "Log time for Thursday?"
    columnIndex = 5
  ElseIf Range("G3") = dateNow Then
    MsgBox "Log time for Friday?"
    columnIndex = 6
  ElseIf Range("H3") = dateNow Then
    MsgBox "Log time for Saturday?"
    columnIndex = 7
  Else
    MsgBox ("Cannot pick correct date column")
  End If

  ' change the column in the range to the correct day
  Set timeStart1 = timeStart1.Columns(columnIndex)
  Set timeEnd1 = timeEnd1.Columns(columnIndex)
  Set timeStart2 = timeStart2.Columns(columnIndex)
  Set timeEnd2 = timeEnd2.Columns(columnIndex)

  ' Test if the value is cell is blank/empty, and mark time for this correct slot
  If IsEmpty(timeStart1) = True Then
    MsgBox "makring time start 1"
    timeStart1.Value = timeNow
  ElseIf IsEmpty(timeEnd1) = True Then
    MsgBox "makring time end 1"
    timeEnd1.Value = timeNow
  ElseIf IsEmpty(timeStart2) = True Then
    MsgBox "makring time start 2"
    timeStart2.Value = timeNow
  ElseIf IsEmpty(timeEnd2) = True Then
    MsgBox "makring time end 2"
    timeEnd2.Value = timeNow
  Else
    MsgBox "No open spaces. Use bonus time"
  End If

End Sub