Sub calculatePayroll()

  ' init ranges with times
  Dim suStart1 As Range
  Dim suEnd1 As Range
  Dim suStart2 As Range
  Dim suEnd2 As Range
  Dim suNetTime as Range
  Dim suNetPay as Range
 
  Dim mStart1 As Range
  Dim mEnd1 As Range
  Dim mStart2 As Range
  Dim mEnd2 As Range
  Dim mNetTime as Range
  Dim mNetPay as Range

  Dim tStart1 As Range
  Dim tEnd1 As Range
  Dim tStart2 As Range
  Dim tEnd2 As Range
  Dim tNetTime as Range
  Dim tNetPay as Range

  Dim wStart1 As Range
  Dim wEnd1 As Range
  Dim wStart2 As Range
  Dim wEnd2 As Range
  Dim wNetTime as Range
  Dim wNetPay as Range

  Dim thStart1 As Range
  Dim thEnd1 As Range
  Dim thStart2 As Range
  Dim thEnd2 As Range
  Dim thNetTime as Range
  Dim thNetPay as Range

  Dim fStart1 As Range
  Dim fEnd1 As Range
  Dim fStart2 As Range
  Dim fEnd2 As Range
  Dim fNetTime as Range
  Dim fNetPay as Range

  Dim saStart1 As Range
  Dim saEnd1 As Range
  Dim saStart2 As Range
  Dim saEnd2 As Range
  Dim saNetTime as Range
  Dim saNetPay as Range

  ' set the ranges
  Set suStart1 = Range("B5")
  Set suEnd1 = Range("B6")
  Set suStart2 = Range("B7")
  Set suEnd2 = Range("B8")
  Set suBonus = Range("b9")
  Set suNetTime = Range("B10")
  Set suNetPay = Range("B11")

  Set mStart1 = Range("C5")
  Set mEnd1 = Range("C6")
  Set mStart2 = Range("C7")
  Set mEnd2 = Range("C8")
  Set mBonus = Range("c9")
  Set mNetTime = Range("C10")
  Set mNetPay = Range("C11")

  Set tStart1 = Range("D5")
  Set tEnd1 = Range("D6")
  Set tStart2 = Range("D7")
  Set tEnd2 = Range("D8")
  Set tBonus = Range("d9")
  Set tNetTime = Range("D10")
  Set tNetPay = Range("D11")

  Set wStart1 = Range("E5")
  Set wEnd1 = Range("E6")
  Set wStart2 = Range("E7")
  Set wEnd2 = Range("E8")
  Set wBonus = Range("e9")
  Set wNetTime = Range("E10")
  Set wNetPay = Range("E11")

  Set thStart1 = Range("F5")
  Set thEnd1 = Range("F6")
  Set thStart2 = Range("F7")
  Set thEnd2 = Range("F8")
  Set thBonus = Range("f9")
  Set thNetTime = Range("F10")
  Set thNetPay = Range("F11")

  Set fStart1 = Range("G5")
  Set fEnd1 = Range("G6")
  Set fStart2 = Range("G7")
  Set fEnd2 = Range("G8")
  Set fBonus = Range("g9")
  Set fNetTime = Range("G10")
  Set fNetPay = Range("G11")
  
  Set saStart1 = Range("H5")
  Set saEnd1 = Range("H6")
  Set saStart2 = Range("H7")
  Set saEnd2 = Range("H8")
  Set saBonus = Range("h9")
  Set saNetTime = Range("H10")
  Set saNetPay = Range("H11")

  suNetTime.Value = netTimeCalculator(suStart1, suEnd1, suStart2, suEnd2, suBonus)
  MsgBox suNetTime

End Sub

' calcuates pay for the time period
Function netTimeCalculator(start1 As Time, end1 As Time, start2 As Time, end2 As Time, bonus as Double) As Double 
  netTimeCalculator = 24 * (bonus + ((end1 - start1) + (end2 - end1)) )  
End Function

' calcuates pay for the time period
Function netPayCalculator(netTime as Double, rate as Double) As Double 
  netPayCalculator = (netTime * rate)  
End Function
