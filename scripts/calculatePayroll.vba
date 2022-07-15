Sub calculatePayroll()

  ' confirm execution
  Dim feedback
  feedback = MsgBox("Calculate time and pay?", vbYesNo + vbQuestion, "Proceed?")
    
  If feedback = vbYes Then
    ' init ranges with times
    Dim timeStart1 As Range
    Dim timeEnd1 As Range
    Dim timeStart2 As Range
    Dim timeEnd2 As Range
    Dim bonusTime As Range
    Dim netTime As Range
    Dim netPay As Range
  
    Set timeStart1 = Range("B5:H5")
    Set timeEnd1 = Range("B6:H6")
    Set timeStart2 = Range("B7:H7")
    Set timeEnd2 = Range("B8:H8")
    Set bonusTime = Range("B9:H9")

    ' result containers
    Set netTime = Range("B10:H10")
    Set netPay = Range("B11:H11")

    Dim i As Integer
    For i = 1 To 7
      netTime.Columns(i).Value = bonusTime.Columns(i) + (24 * ((timeEnd1.Columns(i) - timeStart1.Columns(i)) + (timeEnd2.Columns(i) - timeStart2.Columns(i))))
      netPay.Columns(i).Value = netTime.Columns(i) * 25
    Next i

  End If
End Sub
