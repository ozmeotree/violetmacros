Sub isolate_isbns()

    'This macro finds the column titled "ISBN", moves it to column A, and deletes all other information in the sheet.'

    'We were using this macro to upload our data into Alma to make an itemized set from the ISBNs'
    
    'Set the active sheet as a variable for easier reference
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Check if a filter is currently applied
    If ws.AutoFilterMode Then
        ' Remove the filter
        ws.AutoFilterMode = False
    End If

    'Unfreeze panes'
    ActiveWindow.FreezePanes = False

    ' Add a new column A and make it wider
    ws.Columns("A").Insert
    ws.Columns("A").ColumnWidth = 20

    Dim searchRange As Range
    Set searchRange = ws.UsedRange ' Set the range to search within the used range of the worksheet

    Dim c As Range
    Set c = searchRange.Find("ISBN", LookIn:=xlValues)

    If Not c Is Nothing Then
        c.Select
    Else
        MsgBox "ISBN not found in the worksheet."
    End If

    c.EntireColumn.Copy
    ws.Range("A:A").PasteSpecial xlPasteAll

    'Delete everything else; you can increase column AA if you are working with a larger sheet'
    ws.Range("B:AA").EntireColumn.Delete

    'Fix number format and remove duplicates'
    ws.Columns(1).NumberFormat = "0"
    ws.Range("$A$1:$A$10000").RemoveDuplicates Columns:=1, Header:= _
        xlYes

    'Insert ISBN as heading'
    If ws.Cells(1, 1).Value <> "ISBN" Then
    ws.Cells(1, 1).Insert
    ws.Cells(1, 1).Value = "ISBN"
    End If

    ws.Columns("A").ColumnWidth = 20

End Sub