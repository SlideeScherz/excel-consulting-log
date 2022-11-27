Option Explicit

Sub clearTimeLog()

  ' confirm execution
  Dim feedback
  feedback = MsgBox("Are you sure you want to clear?", vbYesNo + vbQuestion, "Proceed?")

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

End Sub  