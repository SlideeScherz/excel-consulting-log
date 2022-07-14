Sub clockIn()

    ' init and assign date time objects
    Dim dateNow
    Dim timeNow
    dateNow = Date
    timeNow = Now
    
    ' store the date in a cell if needed for new sheet
    Dim dateCell As Range
    Set dateCell = Range("B17")
    dateCell.Value = dateNow
    
    ' init ranges to be assigned based on correct date
    Dim start1 As Range
    Dim end1 As Range
    Dim start2 As Range
    Dim end2 As Range
    
    ' Assign correct range based on day
    If Range("B3") = dateNow Then
        MsgBox "Log time for Sunday?"
        Set start1 = Range("B5")
        Set end1 = Range("B6")
        Set start2 = Range("B7")
        Set end2 = Range("B8")
    ElseIf Range("C3") = dateNow Then
        MsgBox "Log time for Monday?"
        Set start1 = Range("C5")
        Set end1 = Range("C6")
        Set start2 = Range("C7")
        Set end2 = Range("C8")
    ElseIf Range("D3") = dateNow Then
        MsgBox "Log time for Tuesday?"
        Set start1 = Range("D5")
        Set end1 = Range("D6")
        Set start2 = Range("D7")
        Set end2 = Range("D8")
    ElseIf Range("E3") = dateNow Then
        MsgBox "Log time for Wednesday"
        Set start1 = Range("E5")
        Set end1 = Range("E6")
        Set start2 = Range("E7")
        Set end2 = Range("E8")
    ElseIf Range("F3") = dateNow Then
        MsgBox "Log time for Thursday?"
        Set start1 = Range("F5")
        Set end1 = Range("F6")
        Set start2 = Range("F7")
        Set end2 = Range("F8")
    ElseIf Range("G3") = dateNow Then
        MsgBox "Log time for Friday?"
        Set start1 = Range("G5")
        Set end1 = Range("G6")
        Set start2 = Range("G7")
        Set end2 = Range("G8")
    ElseIf Range("H3") = dateNow Then
        MsgBox "Log time for Saturday?"
        Set start1 = Range("H5")
        Set end1 = Range("H6")
        Set start2 = Range("H7")
        Set end2 = Range("H8")
    Else
        MsgBox ("Cannot pick correct date column")
    End If

    ' Test if the value is cell is blank/empty
    ' mark time for this correct slot
    If IsEmpty(start1) = True Then
        MsgBox "start1"
        start1.Value = timeNow
    ElseIf IsEmpty(end1) = True Then
        MsgBox "end1"
        end1.Value = timeNow
    ElseIf IsEmpty(start2) = True Then
        MsgBox "start2"
        start2.Value = timeNow
    ElseIf IsEmpty(end2) = True Then
        MsgBox "end2"
        end2.Value = timeNow
    Else
        MsgBox "No open spaces. Use bonus time"
    End If

End Sub
