Option Explicit

Sub createEntry()

  Public Const DAY_HOURS As Integer = 24

  ' confirm execution
  Dim feedback
  feedback = MsgBox("Calculate time and pay?", vbYesNo + vbQuestion, "Proceed?")
    
  If feedback <> vbYes Then
    Exit Sub
  End If  

  ' init constants
  globals.init

  ' Test if the value is cell is blank/empty, and mark time for this correct slot
  If IsEmpty(timeStart) = True Then
    MsgBox "Begin logging time before running payroll"
  ElseIf IsEmpty(timeEnd) = True Then
    MsgBox "End logging time before running payroll"
  Else
    netTime.Value = DAY_HOURS * (timeEnd.Value - timeStart.Value)
    netPay.Value = netTime * BILLING_RATE
  End If
End Sub
