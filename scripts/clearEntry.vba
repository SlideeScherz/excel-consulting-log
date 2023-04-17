Option Explicit

Sub clearEntry()

  ' confirm execution
  Dim feedback
  feedback = MsgBox("Clear this entry?", vbYesNo + vbQuestion, "Confirm")

  If feedback <> vbYes Then
    Exit Sub
  End If  

  ' init constants
  globals.init
  
  timeDate.Value = ""
  timeWeekDay.Value = ""
  timeStart.Value = ""
  timeEnd.Value = ""
  netTime.Value = ""
  netPay.Value = ""
  goals.Value = ""
  accomplished.Value = ""
  timeStampButton.Caption = "Clock In"

End Sub  