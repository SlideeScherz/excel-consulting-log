Sub clearTimeLog()

' confirm execution
  Dim feedback
  feedback = MsgBox("Are you sure you want to clear?", vbYesNo + vbQuestion, "Proceed?")

  If feedback = vbNo Then
    Exit Sub
  End If  

  Dim timeDate As Range
  Dim timeWeekDay As Range
  Dim timeStart As Range
  Dim timeEnd As Range
  Dim netTime As Range
  Dim netPay As Range
  Dim goals As Range
  Dim accomplished As Range

  Set timeDate = Range("A3")
  Set timeWeekDay = Range("B3")
  Set timeStart = Range("C3")
  Set timeEnd = Range("D3")
  Set netTime = Range("E3")
  Set netPay = Range("F3")
  Set goals = Range("G3")
  Set accomplished = Range("H3")
  
  timeDate.Value = ""
  timeWeekDay.Value = ""
  timeStart.Value = ""
  timeEnd.Value = ""
  netTime.Value = ""
  netPay.Value = ""
  goals.Value = ""
  accomplished.Value = ""

End Sub  