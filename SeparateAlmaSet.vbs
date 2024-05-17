Sub ReservesPullList()

'This macro formats the export of an Alma set into multiple pull lists depending on the title's location. We ended up going a different direction so this macro works but is not a finalized version.'
'It depends on a lot of library-specific configurations so if you want to use it, you'll have to edit it heavily'

    'To do: change "Oversize [*]" to "Oversize", import instructor/course from bookstore list '

'Part 1: Set up the Sheet'
    
    'Prepare the workbooks and sheets as objects for easier reference

    Application.CutCopyMode = False

    'Create Physical and Electronic sheets'
    ActiveSheet.Name = "Physical"
    Sheets.Add After:=ActiveSheet
    Sheets.Add After:=ActiveSheet
    Sheets.Add After:=ActiveSheet
    Worksheets(2).Name = "Science"
    Worksheets(3).Name = "UDC"
    Worksheets(4).Name = "Electronic"

    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim wsS As Worksheet
    Dim wsD As Worksheet
    Dim wkb As Workbook
    Set wkb = Workbooks(ActiveWorkbook.Name)
    Set ws = wkb.Sheets("Physical")
    Set ws2 = wkb.Sheets("Electronic")
    Set wsS = wkb.Sheets("Science")
    Set wsD = wkb.Sheets("UDC")
    Dim cell As Range
    'After looking into variable types more, all of my Longs could be Integer. But, I've kept them as Long for now to keep it consistent.'
    'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary'
    Dim lastrow As Long
    Dim lastRow2 As Long
    Dim lastRowD As Long
    Dim lastRowS As Long

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

'Part 2: Find all of the columns we're going to use and put them in the right order using X = X+1

    'Create variable X as integer'
    Dim X As Long
    X = 3

    ' Find and move the "Title" column'
    Dim TitleCol As Long
    TitleCol = ws.Rows(1).Find("Title").Column
    
    If TitleCol <> X Then
        ws.Columns(TitleCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    'Add Author column'
    ws.Columns(X).Insert
    ws.Cells(1, X).Value = "Author"
    X = X + 1

    'Find and move the "Edition" column'
    Dim EdCol As Long
    EdCol = ws.Rows(1).Find("Edition").Column

    If EdCol <> X Then
        ws.Columns(EdCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    ' Find and move the "Availability" column'
    Dim AvaCol As Long
    AvaCol = ws.Rows(1).Find("Availability").Column
    
    If AvaCol <> X Then
        ws.Columns(AvaCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    ' Find and move the "Type" column'
    Dim TypeCol As Long
    TypeCol = ws.Rows(1).Find("Type").Column
    
    If TypeCol <> X Then
        ws.Columns(TypeCol).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

'Part 3: Delete everything else'

    'Delete any columns not in columns A-G'
    ws.Range("H:BB").EntireColumn.Delete

'Part 4: Assign Variables to all of the new column locations'

    Dim Title As Long
    Title = ws.Rows(1).Find("Title").Column

    Dim Author As Long
    Author = ws.Rows(1).Find("Author").Column

    Dim Edition As Long
    Edition = ws.Rows(1).Find("Edition").Column

    Dim Availability As Long
    Availability = ws.Rows(1).Find("Availability").Column
    ws.Cells(1, Availability).Value = "Location"

    Dim RType As Long
    RType = ws.Rows(1).Find("Type").Column
    ws.Cells(1, RType).Value = "Type"

    'lastRow'
    lastrow = ws.Cells(ws.Rows.Count, Title).End(xlUp).Row

'Part 5: Alter Column Values'

    'Don't name your variables like I do lol'
    Dim InStrr As Long
    Dim Length As Long
    Dim Lengthh As Long

    'Remove extraneous info from Type'
    Dim Loc1 As Long
    Dim Loc2 As Long

    For i = 2 To lastrow
        Set cell = ws.Cells(i, RType)
        Loc1 = InStr(1, cell, "{")
        cell.Value = Mid(cell, (Loc1 + 1))
        Loc2 = InStr(1, cell, "}")
        cell.Value = Left(cell, (Loc2 - 1))
    Next i

    'Replace Ordinal Numbers with Numerals'
    For i = 2 To lastrow
        Set cell = ws.Cells(i, Edition)
            cell.Value = Replace(cell.Value, ".", "")
            cell.Value = Replace(cell, "edition", "")
            cell.Value = Replace(cell, " ed", "")
            cell.Value = Replace(cell, "[", "")
            cell.Value = Replace(cell, "]", "")
            cell.Value = Replace(cell, "First", "1st")
            cell.Value = Replace(cell, "Second", "2nd")
            cell.Value = Replace(cell, "Third", "3rd")
            cell.Value = Replace(cell, "Fourth", "4th")
            cell.Value = Replace(cell, "Fifth", "5th")
            cell.Value = Replace(cell, "Sixth", "6th")
            cell.Value = Replace(cell, "Seventh", "7th")
            cell.Value = Replace(cell, "Eighth", "8th")
            cell.Value = Replace(cell, "Ninth", "9th")
            cell.Value = Replace(cell, "Tenth", "10th")
            cell.Value = Replace(cell, "Eleventh", "11th")
    Next i

    'Trim Location'
    For i = 2 To lastrow
        ws.Cells(i, Availability).Value = Trim(ws.Cells(i, Availability))
    Next i

    'Figuring out what character that weird whitespace was was way more difficult than it had to be'
    'Remove that weird whitespace'
    For i = 2 To lastrow
        Set cell = ws.Cells(i, Availability)
        If InStr(1, cell, ChrW(8195)) > 0 Then
            cell.Value = Replace(cell.Value, ChrW(8195), " ")
        End If
    Next i

    'Remove Availability:, etc From Location, remove line breaks at the end of cells' 'I couldn't get oversize to fully remove when I had the whole thing in there?'
    For i = 2 To lastrow
        Set cell = ws.Cells(i, Availability)
        cell = Replace(cell, "Availability:", "")
        cell = Replace(cell, "Electronic version at ", "")
        cell = Replace(cell, "Physical version temporarily at ", "")
        cell = Replace(cell, "Physical version at ", "")
        Length = Len(cell)
        Lengthh = Len(cell) - 1
        If Len(cell) > 1 Then
            If InStr(Length, cell, Chr(10)) = Length Then
                cell.Value = Left(cell.Value, (Lengthh))
            End If
            If InStr(Length, cell, Chr(13)) = Length Then
                cell.Value = Left(cell.Value, Lengthh)
            End If
            If InStr(Length, cell, vbCrLf) = Length Then
                cell.Value = Left(cell.Value, Lengthh)
            End If
            If InStr(1, cell, "  ") > 0 Then
            cell.Value = Replace(cell.Value, "  ", " ", 1, -1)
            End If
        End If
        If InStr(1, cell, "BINMA Bartle Library: BIN_LND ") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: BIN_LND ", "")
        End If
        If InStr(1, cell, "BINMA Bartle Library: FA FINE ART") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: FA FINE ART", "Fine Arts")
        End If
        If InStr(1, cell, "BINMA Bartle Library: MV3HR Bartle Reserves 3HR") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: MV3HR Bartle Reserves 3HR", "Reserves")
        End If
        Length = Len(cell)
        If InStr(1, cell, "Reserves (1 copy, 1 available)") = 1 Then
            cell.Value = "Reserves {1/1 available}"
        End If
        If InStr(1, cell, "BINMA Bartle Library: MAIN Bartle Stacks") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: MAIN Bartle Stacks", "Bartle Stacks")
        End If
        If InStr(1, cell, "BINDN Downtown Center: DOWN Stacks") > 0 Then
            cell.Value = Replace(cell.Value, "BINDN Downtown Center: DOWN Stacks", "UDC Stacks")
        End If
        If InStr(1, cell, "BINST Library Annex: KTRAY KTRAY") > 0 Then
            cell.Value = Replace(cell.Value, "BINST Library Annex: KTRAY KTRAY", "CMF KTRAY")
        End If
        If InStr(1, cell, "BINST Library Annex: KHOLD KHOLD") > 0 Then
            cell.Value = Replace(cell.Value, "BINST Library Annex: KHOLD KHOLD", "CMF KHOLD")
        End If
        If InStr(1, cell, "BINST Library Annex: RTRAY RTRAY") > 0 Then
            cell.Value = Replace(cell.Value, "BINST Library Annex: RTRAY RTRAY", "CMF RTRAY")
        End If
        If InStr(1, cell, "BINST Library Annex: ") > 0 Then
            cell.Value = Replace(cell.Value, "BINST Library Annex: ", "CMF: ")
        End If
        If InStr(1, cell, "BINMA Bartle Library: BIN_BOR ") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: BIN_BOR ", "")
        End If
        If InStr(1, cell, "BINSP Special Collections: RARC RARC") > 0 Then
            cell.Value = Replace(cell.Value, "BINSP Special Collections: RARC RARC", "Sp Col")
        End If
        If InStr(1, cell, "BINSP Special Collections: RLOCL RLOCL") > 0 Then
            cell.Value = Replace(cell.Value, "BINSP Special Collections: RLOCL RLOCL", "Sp Col")
        End If
        If InStr(1, cell, "BINSP Special Collections: RREF RREF") > 0 Then
            cell.Value = Replace(cell.Value, "BINSP Special Collections: RREF RREF", "Sp Col")
        End If
        If InStr(1, cell, "BINSP Special Collections: RARE SPEC COL") > 0 Then
            cell.Value = Replace(cell.Value, "BINSP Special Collections: RARE SPEC COL", "Sp Col")
        End If
        If InStr(1, cell, "BINSP Special Collections: RPOE RPOE") > 0 Then
            cell.Value = Replace(cell.Value, "BINSP Special Collections: RPOE RPOE", "Sp Col")
        End If
        If InStr(1, cell, "BINMA Bartle Library: LEISURE Loft: Popular Reading") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: LEISURE Loft: Popular Reading", "Loft")
        End If
        If InStr(1, cell, "BINMA Bartle Library: MCAU ") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: MCAU ", "")
        End If
        If InStr(1, cell, "BINMA Bartle Library: MOV Stacks, ") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: MOV Stacks, ", "")
        End If
        If InStr(1, cell, "BINMA Bartle Library: MRDSK Reference (Reference Desk)") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: MRDSK Reference (Reference Desk)", "Reference")
        End If
        If InStr(1, cell, "BINMA Bartle Library: MREF Reference") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: MREF Reference", "Reference")
        End If
        If InStr(1, cell, "BINSC Science Library: SGRD Science Ground") > 0 Then
            cell.Value = Replace(cell.Value, "BINSC Science Library: SGRD Science Ground", "Sci Ground")
        End If
        If InStr(1, cell, "BINSC Science Library: SCI SCIENCE") > 0 Then
            cell.Value = Replace(cell.Value, "BINSC Science Library: SCI SCIENCE", "Sci Stacks")
        End If
        If InStr(1, cell, "BINSC Science Library: SUSTAIN Sustainability Collection") > 0 Then
            cell.Value = Replace(cell.Value, "BINSC Science Library: SUSTAIN Sustainability Collection", "Sustainability")
        End If
        If InStr(1, cell, "BINDN Downtown Center: DOWNR Reference") > 0 Then
            cell.Value = Replace(cell.Value, "BINDN Downtown Center: DOWNR Reference", "UDC Reference")
        End If
        If InStr(1, cell, "BINSC Science Library: SOV Oversize (2nd floor)") > 0 Then
            cell.Value = Replace(cell.Value, "BINSC Science Library: SOV Oversize (2nd floor)", "Sci Oversize")
        End If
        If InStr(1, cell, "BINDN Downtown Center: DV3HR Reserves - 3 hour") > 0 Then
            cell.Value = Replace(cell.Value, "BINDN Downtown Center: DV3HR Reserves - 3 hour", "UDC Reserves")
        End If
        If InStr(1, cell, "BINMA Bartle Library: MORD ") > 0 Then
            cell.Value = Replace(cell.Value, "BINMA Bartle Library: MORD ", "")
        End If
    Next i

    'Move Author information from Title column to Author Column'

    'Find " / " in Title, Value of Author = Right, Value of Title = Left, truncate Author, flip first & last if there's one space and no commas'
    'We sort our Course Reserve books by Author, so having the last name first was a big help to our course reserves staff'
    For i = 2 To lastrow
        Set cell = ws.Cells(i, Title)
        If InStr(1, cell, " / ") > 0 Then
            InStrr = InStr(1, cell, " / ")
            Length = Len(cell)
            ws.Cells(i, Author).Value = Right(cell.Value, (Length - InStrr - 2))
                ws.Cells(i, Author).Value = Replace(ws.Cells(i, Author).Value, ".", "")
                ws.Cells(i, Author).Value = Replace(ws.Cells(i, Author).Value, " ; ", "; ")
                If InStr(1, ws.Cells(i, Author), "by ") = 1 Then
                    ws.Cells(i, Author).Value = Replace(ws.Cells(i, Author).Value, "by ", "", 1, 1)
                End If
                If (Len(Replace(ws.Cells(i, Author).Value, " ", "")) = (Len(ws.Cells(i, Author)) - 1)) And (InStr(1, ws.Cells(i, Author), ",") = 0) Then
                    Loc1 = InStr(1, ws.Cells(i, Author), " ")
                    Lengthh = Len(ws.Cells(i, Author))
                    ws.Cells(i, Author).Value = Right(ws.Cells(i, Author).Value, (Lengthh - Loc1)) & ", " & Left(ws.Cells(i, Author).Value, (Loc1 - 1))
                End If
                If Len(ws.Cells(i, Author)) > 40 Then
                    ws.Cells(i, Author).Value = Left(ws.Cells(i, Author).Value, 40)
                End If
            cell.Value = Left(cell.Value, InStrr)
        End If
        Length = Len(cell)
        If InStr((Length - 1), cell, " /") > 0 Then
            cell.Value = Left(cell.Value, Length - 2)
        End If
    Next i

'Part 6: Move Electronc, Science, and UDC items to a new sheet'

    X = 2

    For i = 2 To lastrow
        If InStr(1, ws.Cells(i, RType), "Electronic") > 0 Then
            ws.Rows(i).EntireRow.Cut ws2.Rows(X)
            Application.CutCopyMode = False
            X = X + 1
        End If
    Next i

    ws.Columns(RType).EntireColumn.Delete
    ws2.Columns(RType).EntireColumn.Delete

    X = 2
    Dim Y As Long
    Y = 2

    For i = 2 To lastrow
        Set cell = ws.Cells(i, Availability)
        If (InStr(1, cell, "Sci ") > 0 Or InStr(1, cell, "Sustainability") > 0) And InStr(1, cell, "Bartle") = 0 And InStr(1, cell, "Reserves") = 0 Then
            ws.Rows(i).EntireRow.Cut wsS.Rows(X)
            Application.CutCopyMode = False
            X = X + 1
        End If
        If InStr(1, cell, "UDC") > 0 And InStr(1, cell, "Bartle") = 0 Then
            ws.Rows(i).EntireRow.Cut wsD.Rows(Y)
            Application.CutCopyMode = False
            Y = Y + 1
        End If
    Next i

'Part 7: Format headers now that everything is separated out'

    'Copy headers'
    ws.Rows(1).EntireRow.Copy ws2.Rows(1)
    ws.Rows(1).EntireRow.Copy wsS.Rows(1)
    ws.Rows(1).EntireRow.Copy wsD.Rows(1)

'Part 8: Sort the Data'

    'Sort the data by location'
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(Availability), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Columns(Title), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A:BB")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort the data by location'
    With ws2.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws2.Columns(Title), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws2.Columns(Availability), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A:BB")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort the data by location'
    With wsS.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsS.Columns(Availability), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=wsS.Columns(Title), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A:BB")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Sort the data by location'
    With wsD.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsD.Columns(Availability), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=wsD.Columns(Title), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A:BB")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'lastRow etc'
    lastrow = ws.Cells(ws.Rows.Count, Title).End(xlUp).Row

    lastRow2 = ws2.Cells(ws2.Rows.Count, Title).End(xlUp).Row

    lastRowS = wsS.Cells(wsS.Rows.Count, Title).End(xlUp).Row

    lastRowD = wsD.Cells(wsD.Rows.Count, Title).End(xlUp).Row

    '(1 copy, 1 available) to {1/1 available} - will only work with 2 single digit numbers sorry regex.replace was not working for me'
    'I fought with Excel for a long time but could not get the regular expression code to work how I thought it was supposed to'
    'This is the regular expression I was using, if you can get it to work'
    '    regexPattern = "\((\d+\s)(?:copy|copies)\,\s(\d+)\s(?:available)\)"
    For i = 2 To lastrow
        Set cell = ws.Cells(i, Availability)
        While InStr(1, cell, "(") > 0 And InStr(1, cell, "available)") > 0
            Locp = InStr(1, cell, "(")
            Locpp = InStr(1, cell, ")")
            Length = Len(cell)
            Phrase = Mid(cell, Locp, (Locpp - Locp + 1))
            Cop = Mid(cell, (Locp + 1), 1)
            Ava = Mid(cell, (Locpp - 11), 1)
            Phrase2 = "{" & Ava & "/" & Cop & " available}"
            If InStr(1, Phrase, "cop") > 0 Then
                cell.Value = Replace(cell, Phrase, Phrase2, 1, -1)
            Else
                cell.Value = Replace(cell, "(", "[", 1, 1)
                cell.Value = Replace(cell, ")", "]", 1, 1)
            End If
        Wend
    Next i

    For i = 2 To lastRowS
        Set cell = wsS.Cells(i, Availability)
        While InStr(1, cell, "(") > 0 And InStr(1, cell, "available)") > 0
            Locp = InStr(1, cell, "(")
            Locpp = InStr(1, cell, ")")
            Length = Len(cell)
            Phrase = Mid(cell, Locp, (Locpp - Locp + 1))
            Cop = Mid(cell, (Locp + 1), 1)
            Ava = Mid(cell, (Locpp - 11), 1)
            Phrase2 = "{" & Ava & "/" & Cop & " available}"
            If InStr(1, Phrase, "cop") > 0 Then
                cell.Value = Replace(cell, Phrase, Phrase2, 1, -1)
            Else
                cell.Value = Replace(cell, "(", "[", 1, 1)
                cell.Value = Replace(cell, ")", "]", 1, 1)
            End If
        Wend
    Next i

    For i = 2 To lastRowD
        Set cell = wsD.Cells(i, Availability)
        While InStr(1, cell, "(") > 0 And InStr(1, cell, "available)") > 0
            Locp = InStr(1, cell, "(")
            Locpp = InStr(1, cell, ")")
            Length = Len(cell)
            Phrase = Mid(cell, Locp, (Locpp - Locp + 1))
            Cop = Mid(cell, (Locp + 1), 1)
            Ava = Mid(cell, (Locpp - 11), 1)
            Phrase2 = "{" & Ava & "/" & Cop & " available}"
            If InStr(1, Phrase, "cop") > 0 Then
                cell.Value = Replace(cell, Phrase, Phrase2, 1, -1)
            Else
                cell.Value = Replace(cell, "(", "[", 1, 1)
                cell.Value = Replace(cell, ")", "]", 1, 1)
            End If
        Wend
    Next i

    'Electronic'
    For i = 2 To lastRow2
        Set cell = (ws2.Cells(i, Availability))
        If InStr(1, cell, "  ") > 0 Then
            cell.Value = Replace(cell.Value, "  ", " ")
        End If
    Next i


'Part 9: Format the sheets'
    'Column Widths'
    'Column widths are weird in excel. If you are viewing the print layout it will show you in inches instead of points, which these are, I think?'
    'You can provide the measurement in inches but you have to use an 'inches to points' thing
    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 5
    ws.Columns(Title).ColumnWidth = 42
    ws.Columns(Author).ColumnWidth = 20
    ws.Columns(Edition).ColumnWidth = 9
    ws.Columns(Availability).ColumnWidth = 50

    ws2.Columns(1).ColumnWidth = 4
    ws2.Columns(2).ColumnWidth = 5
    ws2.Columns(Title).ColumnWidth = 42
    ws2.Columns(Author).ColumnWidth = 20
    ws2.Columns(Edition).ColumnWidth = 9
    ws2.Columns(Availability).ColumnWidth = 50

    wsS.Columns(1).ColumnWidth = 4
    wsS.Columns(2).ColumnWidth = 5
    wsS.Columns(Title).ColumnWidth = 42
    wsS.Columns(Author).ColumnWidth = 20
    wsS.Columns(Edition).ColumnWidth = 9
    wsS.Columns(Availability).ColumnWidth = 50

    wsD.Columns(1).ColumnWidth = 4
    wsD.Columns(2).ColumnWidth = 5
    wsD.Columns(Title).ColumnWidth = 42
    wsD.Columns(Author).ColumnWidth = 20
    wsD.Columns(Edition).ColumnWidth = 9
    wsD.Columns(Availability).ColumnWidth = 50

    'Number B'

    ws.Range("B2:B" & lastrow).FormulaArray = "=ROW() - 1"
    ws.Columns("C").Insert
    ws.Range("B:B").Copy
    ws.Range("C:C").PasteSpecial xlPasteValues
    ws.Range("B:B").EntireColumn.Delete

    ws2.Range("B2:B" & lastRow2).FormulaArray = "=ROW() - 1"
    ws2.Columns("C").Insert
    ws2.Range("B:B").Copy
    ws2.Range("C:C").PasteSpecial xlPasteValues
    ws2.Range("B:B").EntireColumn.Delete

    wsS.Range("B2:B" & lastRowS).FormulaArray = "=ROW() - 1"
    wsS.Columns("C").Insert
    wsS.Range("B:B").Copy
    wsS.Range("C:C").PasteSpecial xlPasteValues
    wsS.Range("B:B").EntireColumn.Delete

    wsD.Range("B2:B" & lastRowD).FormulaArray = "=ROW() - 1"
    wsD.Columns("C").Insert
    wsD.Range("B:B").Copy
    wsD.Range("C:C").PasteSpecial xlPasteValues
    wsD.Range("B:B").EntireColumn.Delete

    'Align Entire Sheet'
    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.HorizontalAlignment = xlLeft
    ws.Columns("B:B").HorizontalAlignment = xlCenter
    ws.Columns(Edition).HorizontalAlignment = xlCenter

    ws2.Cells.VerticalAlignment = xlCenter
    ws2.Cells.HorizontalAlignment = xlLeft
    ws2.Columns("B:B").HorizontalAlignment = xlCenter
    ws2.Columns(Edition).HorizontalAlignment = xlCenter

    wsS.Cells.VerticalAlignment = xlCenter
    wsS.Cells.HorizontalAlignment = xlLeft
    wsS.Columns("B:B").HorizontalAlignment = xlCenter
    wsS.Columns(Edition).HorizontalAlignment = xlCenter
    
    wsD.Cells.VerticalAlignment = xlCenter
    wsD.Cells.HorizontalAlignment = xlLeft
    wsD.Columns("B:B").HorizontalAlignment = xlCenter
    wsD.Columns(Edition).HorizontalAlignment = xlCenter
    

    'Tell Edition not to wrap if it's long
    ws.Columns(Edition).WrapText = True
    ws2.Columns(Edition).WrapText = True
    wsS.Columns(Edition).WrapText = True
    wsD.Columns(Edition).WrapText = True

    For i = 2 To lastrow
        Set cell = ws.Cells(i, Edition)
        If Len(cell) > 10 Then
            cell.HorizontalAlignment = xlLeft
            cell.WrapText = False
        End If
    Next i

    For i = 2 To lastRow2
        Set cell = ws2.Cells(i, Edition)
        If Len(cell) > 10 Then
            cell.HorizontalAlignment = xlLeft
            cell.WrapText = False
        End If
    Next i

    For i = 2 To lastRowS
        Set cell = wsS.Cells(i, Edition)
        If Len(cell) > 10 Then
            cell.HorizontalAlignment = xlLeft
            cell.WrapText = False
        End If
    Next i

    For i = 2 To lastRowD
        Set cell = wsD.Cells(i, Edition)
        If Len(cell) > 10 Then
            cell.HorizontalAlignment = xlLeft
            cell.WrapText = False
        End If
    Next i

    'Format font'
    ws.Cells.Font.Name = "Franklin Gothic Book"
    ws.Cells.Font.Size = 11
    ws.Cells(1, 1).Font.Name = "Wingdings"

    ws2.Cells.Font.Name = "Franklin Gothic Book"
    ws2.Cells.Font.Size = 11
    ws2.Cells(1, 1).Font.Name = "Wingdings"

    wsS.Cells.Font.Name = "Franklin Gothic Book"
    wsS.Cells.Font.Size = 11
    wsS.Cells(1, 1).Font.Name = "Wingdings"

    wsD.Cells.Font.Name = "Franklin Gothic Book"
    wsD.Cells.Font.Size = 11
    wsD.Cells(1, 1).Font.Name = "Wingdings"

    'Trim Location (Again)'
    For i = 2 To lastrow
        ws.Cells(i, Availability).Value = Trim(ws.Cells(i, Availability))
    Next i

    For i = 2 To lastRow2
        ws2.Cells(i, Availability).Value = Trim(ws2.Cells(i, Availability))
    Next i

    For i = 2 To lastRowS
        wsS.Cells(i, Availability).Value = Trim(wsS.Cells(i, Availability))
    Next i

    For i = 2 To lastRowD
        wsD.Cells(i, Availability).Value = Trim(wsD.Cells(i, Availability))
    Next i

    'Autofit Row Height'
    ws.Rows.AutoFit
    ws2.Rows.AutoFit
    wsS.Rows.AutoFit
    wsD.Rows.AutoFit

    ' Check if a filter is currently applied
    If ws.AutoFilterMode Then
        ' Remove the filter
        ws.AutoFilterMode = False
    End If

    If ws2.AutoFilterMode Then
        ' Remove the filter
        ws2.AutoFilterMode = False
    End If

    If wsS.AutoFilterMode Then
        ' Remove the filter
        wsS.AutoFilterMode = False
    End If

    If wsD.AutoFilterMode Then
        ' Remove the filter
        wsD.AutoFilterMode = False
    End If

    'Format the sheet as a table with alternating row colors
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes).TableStyle = "TableStyleLight16"
    ws.Range("A1:F1").Interior.Color = RGB(255, 255, 255)
    ws.Activate
    ws.Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Size = 14
    End With

    'Format the sheet as a table with alternating row colors
    ws2.ListObjects.Add(xlSrcRange, ws2.Range("A1").CurrentRegion, , xlYes).TableStyle = "TableStyleLight15"
    ws2.Range("A1:F1").Interior.Color = RGB(255, 255, 255)
    ws2.Activate
    ws2.Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Size = 14
    End With

    'Format the sheet as a table with alternating row colors
    wsS.ListObjects.Add(xlSrcRange, wsS.Range("A1").CurrentRegion, , xlYes).TableStyle = "TableStyleLight19"
    wsS.Range("A1:F1").Interior.Color = RGB(255, 255, 255)
    wsS.Activate
    wsS.Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Size = 14
    End With

    'Format the sheet as a table with alternating row colors
    wsD.ListObjects.Add(xlSrcRange, wsD.Range("A1").CurrentRegion, , xlYes).TableStyle = "TableStyleLight21"
    wsD.Range("A1:F1").Interior.Color = RGB(255, 255, 255)
    wsD.Activate
    wsD.Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Size = 14
    End With

    ' Check if a filter is currently applied (again)
    If ws.AutoFilterMode Then
        ' Remove the filter
        ws.AutoFilterMode = False
    End If

    If ws2.AutoFilterMode Then
        ' Remove the filter
        ws2.AutoFilterMode = False
    End If

    If wsS.AutoFilterMode Then
        ' Remove the filter
        wsS.AutoFilterMode = False
    End If

    If wsD.AutoFilterMode Then
        ' Remove the filter
        wsD.AutoFilterMode = False
    End If

    'Autofit Row Height (again)'
    ws.Rows.AutoFit
    ws2.Rows.AutoFit
    wsS.Rows.AutoFit
    wsD.Rows.AutoFit

'Part 10: Format Sheet for Printing'

    'Make it landscape'
    ws.PageSetup.Orientation = xlLandscape
    ws2.PageSetup.Orientation = xlLandscape
    wsS.PageSetup.Orientation = xlLandscape
    wsD.PageSetup.Orientation = xlLandscape

    'Print all columns on 1 page and freeze header row'
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
    End With

    With ws2.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
    End With

    With wsS.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
    End With

    With wsD.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
    End With

    ' Set all 4 margins to 0.25 and add page number footer"
    'This is the inches to points thing I mentioned earlier'
    With ws.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
        .RightFooter = "&""Franklin Gothic Book,Regular""&10&A - Page &P of &N"
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
    End With

    With ws2.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
        .RightFooter = "&""Franklin Gothic Book,Regular""&10&A - Page &P of &N"
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
    End With

    With wsS.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
        .RightFooter = "&""Franklin Gothic Book,Regular""&10&A - Page &P of &N"
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
    End With

    With wsD.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
        .RightFooter = "&""Franklin Gothic Book,Regular""&10&A - Page &P of &N"
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
    End With

    'Autofit Row Height (again)'
    'I was having issues with autofit not working so I just put it in multiple times'
    ws.Rows.AutoFit
    ws2.Rows.AutoFit
    wsS.Rows.AutoFit
    wsD.Rows.AutoFit

    'If you don't put a select at the end it will stay on the last thing it selected, which can sometimes look weird
    ws.Activate
    ws.Range("A1:A1").Select
    'End Select

End Sub
