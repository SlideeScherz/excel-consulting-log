Option Explicit

Public timeDate As Range
Public timeWeekDay As Range
Public timeStart As Range
Public timeEnd As Range
Public netTime As Range
Public netPay As Range
Public goals As Range
Public accomplished As Range
Public timeStampButton as Button
Public createEntryButton as Button

Public Sub init()
  Set timeDate = Range("A2")
  Set timeWeekDay = Range("B2")
  Set timeStart = Range("C2")
  Set timeEnd = Range("D2")
  Set netTime = Range("E2")
  Set netPay = Range("F2")
  Set goals = Range("G2")
  Set accomplished = Range("H2")
  Set timeStampButton = ActiveSheet.Buttons("timeStampButton")
  'Set createEntryButton = ActiveSheet.Buttons("createEntryButton")
End Sub  