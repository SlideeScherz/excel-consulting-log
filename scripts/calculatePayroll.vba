Option Explicit

Sub calculatePayroll()

  ' confirm execution
  Dim feedback
  feedback = MsgBox("Calculate time and pay?", vbYesNo + vbQuestion, "Proceed?")
    
  If feedback <> vbYes Then
    Exit Sub
  End If  

  Set timeStart = Range("C3")
  Set timeEnd = Range("D3")
  Set netTime = Range("E3")
  Set netPay = Range("F3")

  ' Test if the value is cell is blank/empty, and mark time for this correct slot
  If IsEmpty(timeStart) = True Then
    MsgBox "Begin logging time before running payroll"
  ElseIf IsEmpty(timeEnd) = True Then
    MsgBox "End logging time before running payroll"
  Else
    netTime.Value = 24 * (timeEnd.Value - timeStart.Value)
    netPay.Value = netTime * 25
  End If
End Sub
