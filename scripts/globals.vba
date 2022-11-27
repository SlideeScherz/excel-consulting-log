Option Explicit

Public timeDate As Range
Public timeWeekDay As Range
Public timeStart As Range
Public timeEnd As Range
Public netTime As Range
Public netPay As Range
Public goals As Range
Public accomplished As Range

Public Sub init()
  Set timeDate = Range("A3")
  Set timeWeekDay = Range("B3")
  Set timeStart = Range("C3")
  Set timeEnd = Range("D3")
  Set netTime = Range("E3")
  Set netPay = Range("F3")
  Set goals = Range("G3")
  Set accomplished = Range("H3")
End Sub  