Sub NextDateArray()

    'For this macro, I had hourly data that I was trying to combine into days. I could not get Excel's autofill to work as I wanted, so I wrote this little macro instead.

    'Set the active sheet as a variable for easier reference
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim cell As Range
    Dim O As Long
    Dim P As Long
    Dim X As Long
    Dim NewFormula As String

    'Find lastRow'
    lastRow = ws.Cells(ws.Rows.Count, 17).End(xlUp).Row
    O = 25
    P = 48

    'Add 24 to both cell addresses'
    For i = 11 To lastRow
        Set cell = ws.Cells(i, 19)
        cell.Value = "=MAX(O" & O & ":P" & P & ")"
        O = O + 24
        P = P + 24
    Next i

End Sub

Sub NextNextDateArray()

    'Set the active sheet as a variable for easier reference
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim cell As Range
    Dim O As Long
    Dim P As Long
    Dim X As Long
    Dim NewFormula As String

    'Find lastRow'
    lastRow = ws.Cells(ws.Rows.Count, 17).End(xlUp).Row
    O = 25
    P = 48

    'Add 24 to both cell addresses'
    For i = 11 To lastRow
        Set cell = ws.Cells(i, 22)
        cell.Value = "=MAX(O" & O & ":O" & P & ")"
        O = O + 24
        P = P + 24
    Next i
    
End Sub
