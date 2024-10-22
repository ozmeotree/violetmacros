Sub ExpiredHoldShelf()
    'This macro formats the Expired Hold Shelf export from Alma into a printable pull list for staff.

    ' Set the active sheet as a variable for easier reference
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Check if a filter is currently applied
    If ws.AutoFilterMode Then
        ' Remove the filter
        ws.AutoFilterMode = False
    End If

    ' Add a new column A with the header 'ü' and 5px wide
    ws.Columns("A").Insert
    ws.Cells(1, 1).Value = "ü"
    ws.Columns("A").ColumnWidth = 5

    ' Add a new column B with the header '#' and 5px wide, and number the rows (but as values)
    ws.Columns("B").Insert
    ws.Cells(1, 2).Value = "#"
    ws.Columns("B").ColumnWidth = 5
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    ws.Range("B2:B" & lastRow).FormulaArray = "=ROW() - 1"
    ws.Columns("C").Insert
    ws.Range("B:B").Copy
    ws.Range("C:C").PasteSpecial xlPasteValues
    ws.Range("B:B").EntireColumn.Delete

    'Make title column C'
    ' Find the column number of the "Title" header
    Dim TitleCol As Long
    TitleCol = ws.Rows(1).Find("Title").Column
   
    If TitleCol <> 3 Then
        ' Cut and insert the "Title" column to be column F
        ws.Columns(TitleCol).Cut
        ws.Columns("C").Insert Shift:=xlToRight
    End If

    'Make Barcode column D'
    ' Find the column number of the "Barcode" header
    Dim BarcodeCol As Long
    BarcodeCol = ws.Rows(1).Find("Barcode").Column
   
    If BarcodeCol <> 4 Then
        ' Cut and insert the "Barcode" column to be column D
        ws.Columns(BarcodeCol).Cut
        ws.Columns("D").Insert Shift:=xlToRight
    End If

    'Make Held For column E'
    ' Find the column number of the "Held For" header
    Dim HeldCol As Long
    HeldCol = ws.Rows(1).Find("Held For").Column
   
    If HeldCol <> 5 Then
        ' Cut and insert the "Held For" column to be column E
        ws.Columns(HeldCol).Cut
        ws.Columns("E").Insert Shift:=xlToRight
    End If

    'Make Preferred Identifier column F'
    ' Find the column number of the "Preferred Identifier" header
    Dim IDCol As Long
    IDCol = ws.Rows(1).Find("Preferred Identifier").Column
   
    If IDCol <> 6 Then
        ' Cut and insert the "Preferred Identifier" column to be column F
        ws.Columns(IDCol).Cut
        ws.Columns("F").Insert Shift:=xlToRight
    End If

    'Make Held Until column G'
    ' Find the column number of the "Held Until" header
    Dim UntilCol As Long
    UntilCol = ws.Rows(1).Find("Held Until").Column
   
    If UntilCol <> 7 Then
        ' Cut and insert the "Location" column to be column G
        ws.Columns(UntilCol).Cut
        ws.Columns("G").Insert Shift:=xlToRight
    End If

    'Change header of Column C to Title
    ws.Cells(1, 3).Value = "Title"

    'Change header of Column F to Username'
    ws.Cells(1, 6).Value = "Username"

    'Format Column Widths'
    ws.Columns("C").ColumnWidth = 53
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 23
    ws.Columns("E").ColumnWidth = 33
    ws.Columns("G").ColumnWidth = 15

    'Truncate Title and then Resize'
    'Find the last row of data in column C
    lastRoww = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    'Loop through each cell in column F and truncate the value if it exceeds 150 characters
    For i = 1 To lastRoww
        If Len(ws.Cells(i, "C").Value) > 150 Then
            ws.Cells(i, "C").Value = Left(ws.Cells(i, "C").Value, 150)
        End If
    Next i

    'Resize F'
    ws.Columns("C").AutoFit

    If ws.Columns("C").ColumnWidth > 50 Then
        ws.Columns("C").ColumnWidth = 50
    End If

    ' Format almost the whole sheet as Arial, 14pt font
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 14
    ws.Cells(1, 1).Font.Name = "Wingdings"
    ws.Columns("C").Font.Size = 12

    'Make it landscape'
    ws.PageSetup.Orientation = xlLandscape

    ' Format the sheet as a table with alternating row colors
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes).TableStyle = "TableStyleLight15"
    Range("A1:G1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
     End With
    With Selection.Font
        .Size = 14
       
    End With
   
    ' Check if a filter is currently applied
    If ws.AutoFilterMode Then
        ' Remove the filter
        ws.AutoFilterMode = False
    End If

    'Middle Align Entire Sheet'
    ws.Cells.VerticalAlignment = xlCenter
    ws.Columns.HorizontalAlignment = xlCenter
    ws.Columns("C:C").HorizontalAlignment = xlLeft
   
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
   
    ' Delete any columns not in columns A-G
    ws.Range("H:BB").EntireColumn.Delete

    'Autofit Row Height'
    ws.Rows.AutoFit

    ' Set all 4 margins to 0.25"
    With ws.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
    End With
   
    ' Set the header and footer to 0"
    With ws.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    End With

    'Autofit rows again'
    ws.Rows.AutoFit
    Rows("1:1").RowHeight = 40.75

End Sub
