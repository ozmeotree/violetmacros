Sub PickFromShelfBorrowing()

'Part 1: Set up the Sheet'
    
    'Set the active sheet as a variable for easier reference
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim cell As Range

    ' Check if a filter is currently applied
    If ws.AutoFilterMode Then
        ' Remove the filter
        ws.AutoFilterMode = False
    End If

    ' Add a new column A with the header 'ü' and 5px wide
    ws.Columns("A").Insert
    ws.Cells(1, 1).Value = "ü"
    ws.Columns("A").ColumnWidth = 4

    ' Add a new column B with the header '#' and 5px wide, and number the rows (but as values)
    ws.Columns("B").Insert
    ws.Cells(1, 2).Value = "#"
    ws.Columns("B").ColumnWidth = 4

'Part 2: Find all of the columns we're going to use and put them in the right order using X = X+1 to keep the columns in the ordered order

    'Create variable X as integer'
    Dim X As Long
    X = 3

    ' Find and move the "Call Number" column'
    Dim CallCol As Long
    CallCol = ws.Rows(1).Find("Call Number").Column
    
    If CallCol <> X Then
        ws.Columns(CallCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    ' Find and move "Location" column'
    Dim LocationCol As Long
    LocationCol = ws.Rows(1).Find("Location").Column
    
    If LocationCol <> X Then
        ws.Columns(LocationCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    ' Find and move the "Title" column'
    Dim TitleCol As Long
    TitleCol = ws.Rows(1).Find("Title").Column
    
    If TitleCol <> X Then
        ws.Columns(TitleCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    ' Find and move the "Barcode" column'
    Dim BarcodeCol As Long
    BarcodeCol = ws.Rows(1).Find("Barcode").Column
    
    If BarcodeCol <> X Then
        ws.Columns(BarcodeCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Find and move the "Pickup Location" column'
    Dim PickupCol As Long
    PickupCol = ws.Rows(1).Find("Pickup Location").Column

    If PickupCol <> X Then
        ws.Columns(PickupCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Find and move the "Requester User Group" column'
    Dim UserCol As Long
    UserCol = ws.Rows(1).Find("Requester User Group").Column

    If UserCol <> X Then
        ws.Columns(UserCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Find and move the "Request Note" column'
    Dim RNCol As Long
    RNCol = ws.Rows(1).Find("Request Note").Column
    
    If RNCol <> X Then
        ws.Columns(RNCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Find and move the "Volume" column'
    Dim VolCol As Long
    VolCol = ws.Rows(1).Find("Volume").Column
    
    If VolCol <> X Then
        ws.Columns(VolCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Find and move the "Issue" column'
    Dim IssCol As Long
    IssCol = ws.Rows(1).Find("Issue").Column
    
    If IssCol <> X Then
        ws.Columns(IssCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Find and move the "Pages" column'
    Dim PageCol As Long
    PageCol = ws.Rows(1).Find("Pages").Column

    If PageCol <> X Then
        ws.Columns(PageCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Find and move the "Request Type" column'
    Dim TypeCol As Long
    TypeCol = ws.Rows(1).Find("Request Type").Column
    
    If TypeCol <> X Then
        ws.Columns(TypeCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Find and move the "Resource Sharing Volume" column'
    Dim RSVolumeCol As Long
    RSVolumeCol = ws.Rows(1).Find("Resource Sharing Volume").Column

    If RSVolumeCol <> X Then
        ws.Columns(RSVolumeCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    ' Find and move "Description" column'
    Dim DescCol As Long
    DescCol = ws.Rows(1).Find("Description").Column
    
    If DescCol <> X Then
        ws.Columns(DescCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

'Part 3: Assign Variables to all of the new column locations'

    Dim CallNumber As Long
    CallNumber = ws.Rows(1).Find("Call Number").Column

    Dim Location As Long
    Location = ws.Rows(1).Find("Location").Column

    Dim Title As Long
    Title = ws.Rows(1).Find("Title").Column

    Dim Barcode As Long
    Barcode = ws.Rows(1).Find("Barcode").Column

    Dim Pickup As Long
    Pickup = ws.Rows(1).Find("Pickup Location").Column
    ws.Cells(1, Pickup).Value = "Hold Shelf"

    Dim User As Long
    User = ws.Rows(1).Find("Requester User Group").Column
    ws.Cells(1, User).Value = "User"

    Dim RequestNote As Long
    RequestNote = ws.Rows(1).Find("Request Note").Column
    
    Dim Volume As Long
    Volume = ws.Rows(1).Find("Volume").Column

    Dim Issue As Long
    Issue = ws.Rows(1).Find("Issue").Column

    Dim Pages As Long
    Pages = ws.Rows(1).Find("Pages").Column

    Dim RType As Long
    RType = ws.Rows(1).Find("Request Type").Column
    ws.Cells(1, RType).Value = "Type"

    Dim RSVolume As Long
    RSVolume = ws.Rows(1).Find("Resource Sharing Volume").Column

    Dim Description As Long
    Description = ws.Rows(1).Find("Description").Column
    ws.Cells(1, Description).Value = "Description & Notes"

    'lastRow'
    lastRow = ws.Cells(ws.Rows.Count, Title).End(xlUp).Row

'Part 4: Alter Column Values'

    'Add floor number to location'
    'vbCrLf does not work on macs; there's another word for it'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Location)
        If InStr(1, LCase(cell), "(fa)") > 0 Then
            cell.Value = "Fine Arts"
        ElseIf InStr(1, LCase(cell), "(fm2)") > 0 Then
            cell.Value = "Fine Arts"

        ElseIf InStr(1, LCase(cell), "reserves") > 0 Then
            cell.Value = "Floor 1 Reserves"
        ElseIf InStr(1, LCase(cell), "manew") > 0 Then
            cell.Value = "Floor 1" & vbCrLf & "Bartle New Books"
        ElseIf InStr(1, LCase(cell), "(mov)") > 0 Then
            cell.Value = "Floor 2 Oversize"
        ElseIf InStr(1, LCase(cell), "(movth)") > 0 Then
            cell.Value = "Floor 2 Oversize Thesis"
        ElseIf InStr(1, LCase(cell), "oversize") > 0 Then
            cell.Value = "Floor 2 " & cell.Value
        ElseIf InStr(1, LCase(cell), "reference") > 0 Then
            cell.Value = "Floor 1" & vbCrLf & cell.Value
        ElseIf InStr(1, LCase(cell), "mezzanine") > 0 Then
            cell.Value = "Floor 2" & vbCrLf & cell.Value
        ElseIf InStr(1, LCase(cell), "(mmdvd)") > 0 Then
            cell.Value = "Floor 1 DVD"
        ElseIf InStr(1, LCase(cell), "dvd") > 0 Then
            cell.Value = "Floor 1 " & cell.Value
        ElseIf InStr(1, LCase(cell), "mwang") > 0 Then
            cell.Value = "Floor 4 Stacks"
        ElseIf InStr(1, LCase(cell), "(main)") > 0 Then
            If Left(ws.Cells(i, CallNumber).Value, 1) = "A" Then
                cell.Value = "Floor 1 Reference"
            Else: cell.Value = "Floor 4 Stacks"
            End If
        ElseIf InStr(1, LCase(cell), "(mgdoc)") > 0 Then
            cell.Value = "Floor 1" & vbCrLf & "Gov Docs"
        ElseIf InStr(1, LCase(cell), "gov") > 0 Then
            cell.Value = "Floor 1" & vbCrLf & cell.Value
        ElseIf InStr(1, LCase(cell), "arabic") > 0 Then
            cell.Value = "Floor 2" & vbCrLf & cell.Value
        ElseIf InStr(1, LCase(cell), "hebrew") > 0 Then
            cell.Value = "Floor 2" & vbCrLf & cell.Value
        ElseIf InStr(1, LCase(cell), "holocaust") > 0 Then
            cell.Value = "Floor 2" & vbCrLf & cell.Value
        ElseIf InStr(1, LCase(cell), "persia") > 0 Then
            cell.Value = "Floor 2" & vbCrLf & cell.Value
        ElseIf InStr(1, LCase(cell), "loft") > 0 Then
            cell.Value = "Floor 2" & vbCrLf & cell.Value
        ElseIf InStr(1, LCase(cell), "leisure") > 0 Then
            cell.Value = "Floor 2 Loft"
        ElseIf InStr(1, LCase(cell), "success") > 0 Then
            cell.Value = "Floor 2 Success Shelf"
        ElseIf InStr(1, LCase(cell), "loops") > 0 Then
            cell.Value = "Ground Floor" & vbCrLf & cell.Value
        End If
    Next i

    'Sort the data by location'
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(Location), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Columns(CallNumber), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A:BB")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Number B'
    ws.Range("B2:B" & lastRow).FormulaArray = "=ROW() - 1"
    ws.Columns("C").Insert
    ws.Range("B:B").Copy
    ws.Range("C:C").PasteSpecial xlPasteValues
    ws.Range("B:B").EntireColumn.Delete

    'Loop through each cell in Title and truncate the value if it exceeds 100 characters
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Title)
        If Len(cell.Value) > 75 Then
            cell.Value = Left(cell.Value, 75)
        End If
    Next i

    'Send error if there are too many barcodes'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Barcode)
        cell.NumberFormat = "0"
        cell.Value = Replace(cell.Value, ",", vbCrLf)
        If Len(cell.Value) > 75 Then
            cell.Value = "Too many barcodes; please check request"
        End If
    Next i
    
    'Loop through each cell in Description and truncate the value if it exceeds 85 characters, add note'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Description)
        If Left(cell.Value, 21) = "Chapter/Article Title" Then
            cell.Value = "Item(s) may not scan-in properly; check request if so."
        ElseIf Len(cell.Value) > 50 Then
            cell.Value = Left(cell.Value, 50)
        End If
    Next i

    'Remove "Title:" from the Request Note'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, RequestNote)
        If Left(cell.Value, 6) = "Title:" Then
            cell.Value = Mid(cell.Value, 7)
        End If
    Next i

    'If Request Note doesn't match the first 7 characters of Title, Concatenate'
    For i = 2 To lastRow
        If Left(LCase(ws.Cells(i, RequestNote).Value), 7) <> Left(LCase(ws.Cells(i, Title).Value), 7) Then
            ws.Cells(i, Description).Value = ws.Cells(i, Description).Value & ws.Cells(i, RequestNote).Value
        End If
    Next i

    'Loop through each cell in Description and truncate the value if it exceeds 125 characters this time
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Description)
        If Len(cell.Value) > 50 Then
            cell.Value = Left(cell.Value, 50)
        End If
    Next i

    'Append Resource Sharing Volume to Description'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Description)
        If InStr(1, LCase(cell), LCase(ws.Cells(i, RSVolume))) = 0 And Len(ws.Cells(i, RSVolume)) > 0 Then
            cell.Value = cell & " " & ws.Cells(i, RSVolume)
        End If
    Next i

    'Append Volume to Description'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Description)
        If InStr(1, LCase(cell), LCase(ws.Cells(i, Volume))) = 0 And Len(ws.Cells(i, Volume)) > 0 Then
            cell.Value = cell & " v." & ws.Cells(i, Volume)
        End If
    Next i

    'Append Issue to Description'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Description)
        If InStr(1, LCase(cell), LCase(ws.Cells(i, Issue))) = 0 And Len(ws.Cells(i, Issue)) > 0 Then
            cell.Value = cell & " i." & ws.Cells(i, Issue)
        End If
    Next i

    'Append Pages to Description'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Description)
        If InStr(1, LCase(cell), LCase(ws.Cells(i, Pages))) = 0 And Len(ws.Cells(i, Pages)) > 0 Then
            cell.Value = cell & " p." & ws.Cells(i, Pages)
        ElseIf Len(ws.Cells(i, Pages)) > 0 Then
            cell.Value = cell.Value & " p." & ws.Cells(i, Pages)
        End If
    Next i

    'Loop through each cell in Request Type and truncate the value if it exceeds 13 characters
    For i = 2 To lastRow
        Set cell = ws.Cells(i, RType)
        If Len(cell.Value) > 13 Then
            cell.Value = Left(cell.Value, 13)
        End If
    Next i

    'Change Usergroup values to codes'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, User)
        If cell.Value = "Undergrad" Then
            cell.Value = "UG"
        ElseIf cell.Value = "Grad student" Then
            cell.Value = "GR"
        ElseIf cell.Value = "Faculty" Then
            cell.Value = "FAC"
        ElseIf cell.Value = "Faculty/Staff" Then
            cell.Value = "FAC"
        ElseIf cell.Value = "Preservation Department" Then
            cell.Value = "Presv"
        ElseIf cell.Value = "Local Borrower" Then
            cell.Value = "Cancel if NOS"
        ElseIf cell.Value = "Alumni" Then
            cell.Value = "Cancel if NOS"
        ElseIf cell.Value = "Retiree" Then
            cell.Value = "Cancel if NOS"
        ElseIf cell.Value = "Volunteer" Then
            cell.Value = "Cancel if NOS"
        ElseIf cell.Value = "University Programs" Then
            cell.Value = "Cancel if NOS"
        End If
    Next i

    'Change Pickup Location values to codes'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Pickup)
        If cell.Value = "Preservation Department" Then
            cell.Value = "Presv"
        ElseIf cell.Value = "Downtown Center" Then
            cell.Value = "UDC"
        ElseIf cell.Value = "Science Library" Then
            cell.Value = "Sci"
        ElseIf cell.Value = "Bartle Library" Then
            cell.Value = "Bartle"
        End If
    Next i

    'Combine usergroup & pickup location columns'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Pickup)
        If ws.Cells(i, User) = "Cancel if NOS" Then
            cell.Value = cell.Value & ", " & ws.Cells(i, User)
        End If
    Next i

    'Combine Call Number & Description columns'
    For i = 2 To lastRow
        Set cell = ws.Cells(i, Description)
        If Len(cell.Value) > 1 Then
            ws.Cells(i, CallNumber).Value = ws.Cells(i, CallNumber).Value & vbCrLf & cell.Value
        End If
    Next i

'Part 5: Delete Extraneous Data'

    'Delete any columns not in columns A-H'
    ws.Range("H:BB").EntireColumn.Delete

'Part 6: Assign Column Widths'

    'Middle Align Entire Sheet, Center columns'
    ws.Cells.VerticalAlignment = xlCenter
    ws.Columns("B:B").HorizontalAlignment = xlCenter
    ws.Columns(Barcode).HorizontalAlignment = xlCenter
    ws.Columns(Pickup).HorizontalAlignment = xlCenter

    'Format Column Widths'
    ws.Columns(CallNumber).ColumnWidth = 19
    ws.Columns(Location).ColumnWidth = 17.25
    ws.Columns(Barcode).ColumnWidth = 22
    'ws.Columns(User).ColumnWidth = 7.57
    'ws.Columns(Description).ColumnWidth = 21
    ws.Columns(Title).ColumnWidth = 30.63
    'ws.Columns(RType).ColumnWidth = 7.86
    ws.Columns("A").ColumnWidth = 4

    'Tell Pickup/User not to wrap'
        'Columns(Pickup).Select
    'With Selection
        '.WrapText = False
    'End With

'Part 7: Format Table'

    ' Format almost the whole sheet as Cascadia Code, 12pt font
    ws.Cells.Font.Name = "Cascadia Code"
    ws.Cells.Font.Size = 12
    ws.Cells(1, 1).Font.Name = "Wingdings"
    ws.Columns(Pickup).Font.Size = 10
    'ws.Columns(Title).Font.Size = 12
    'ws.Columns(RType).Font.Size = 12
    ws.Columns(CallNumber).Font.Bold = True

    'Format the sheet as a table with alternating row colors
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes).TableStyle = "TableStyleLight18"
    ws.Range("A1:G1").Interior.Color = RGB(255, 255, 255)
    Range("A1:G1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    'With Selection.Font
        '.Size = 14
    'End With

    ' Check if a filter is currently applied
    If ws.AutoFilterMode Then
        ' Remove the filter
        ws.AutoFilterMode = False
    End If

    'Autofit Row Height'
    ws.Rows.AutoFit
    Rows("1:1").RowHeight = 40.75

    'Highlight rows based on common locations'
    'Range("A2:H" & lastRow).Select
    'Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        '"=$E2=""Floor 4 Stacks"""
        'With Selection.FormatConditions(1).Interior
        '.PatternColorIndex = xlAutomatic
        '.ThemeColor = xlThemeColorAccent2
        '.TintAndShade = 0.8
        'End With
    'Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        '"=$E2=""Fine Arts"""
    'Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    'With Selection.FormatConditions(1).Interior
        '.PatternColorIndex = xlAutomatic
        '.ThemeColor = xlThemeColorAccent6
        '.TintAndShade = 0.8
    'End With

'Part 8: Format Sheet for Printing'

    'Make it landscape'
    ws.PageSetup.Orientation = xlLandscape

    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
    End With

    ' Set all 4 margins to 0.25"
    With ws.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.33)
        .BottomMargin = Application.InchesToPoints(0.25)
        .FooterMargin = 0
    End With
    
    ' Set the header and footer to 0"
    With ws.PageSetup
        .LeftHeader = "&""Franklin Gothic Book,Regular""&10 Pick Up Requested Resources | Printed on &D"
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .HeaderMargin = Application.InchesToPoints(0.15)
        .DifferentFirstPageHeaderFooter = False
        '.FirstPage.LeftHeader.Text = "&""Franklin Gothic Book,Regular""&10 Pick Up Requested Resources | Printed on &D"
    End With

    'Format Column Widths'
    ws.Columns(CallNumber).ColumnWidth = 28.75
    ws.Columns(Location).ColumnWidth = 17.25
    ws.Columns(Barcode).ColumnWidth = 21.13
    'ws.Columns(User).ColumnWidth = 7.57
    'ws.Columns(Description).ColumnWidth = 21
    ws.Columns(Title).ColumnWidth = 48.5
    'ws.Columns(RType).ColumnWidth = 7.86
    ws.Columns("A").ColumnWidth = 4
    ws.Columns(Pickup).ColumnWidth = 8.63
    ws.Columns("A").ColumnWidth = 4

    'Autofit rows again'
    ws.Rows.AutoFit
    Rows("1:1").RowHeight = 40.75

    Range("A1:A1").Select

End Sub
