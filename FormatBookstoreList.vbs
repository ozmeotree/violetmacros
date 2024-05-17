Sub FormatBookstoreList()

'This macro formats our textbook vendor's export of all textbooks and courses for next semester.
'There are a lot of columns that could contain important information, but are not helpful at-a-glance, so this macro hides them.
'After formating this into a readable state, I run Princeton's Excel Alma Lookup tool (https://github.com/pulibrary/ExcelAlmaLookup/#readme) to see what textbooks we have in our catalog (and where they are)


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

    'Rename sheet with B5 (semester)'
    Worksheets(1).Name = ws.Cells(5, 2).Value

    'Delete rows that don't have data in column F'
    Do While ws.Cells(1, 6).Value = ""
        ws.Rows(1).Delete
    Loop

'Part 2: Find all of the columns we're going to use and put them in the right order using X = X+1'
'School, Department, Course Code, Course Title, SKU, Digital Available, Note for Bookstore, Price at Time of Adoption, Submitter, Submitter Email, Date Submitted'

    'Create variable X as integer'
    Dim X As Long
    X = 3

    'Find and move the "School" column'
    Dim Col As Long
    Col = ws.Rows(1).Find("School").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Department").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Course Code").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Course Title").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    If InStr(1, ws.Cells(1, X).Value, "Course Title") > 0 Then
        ws.Cells(1, X).Value = "Course Name"
    End If

    'The book column Title and Course Title were confusing each other so I had it print a MsgBox so I could see which column it was picking up here
    'Dim aaa As String
    'aaa = ws.Cells(1, X).Value
    'MsgBox aaa, vbOKOnly
    'X = X + 1


    Col = ws.Rows(1).Find("SKU").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Digital Available").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Price").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Note for Bookstore").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Submitter").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Submitter Email").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Date Submitted").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Date Submitted").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Dim DeleteThese As Long
    DeleteThese = X

    Col = ws.Rows(1).Find("Section").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Instructor").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("ISBN").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1


    'Now that I've renamed Course Title this line is probably unnecessary but why change something that's working?
    Col = ws.Rows(1).Find(What:="Title", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    If ws.Cells(1, X).Value = "Title" Then
        ws.Cells(1, X).Value = "Title"
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Author").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Edition").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

    Col = ws.Rows(1).Find("Type").Column
    
    If Col <> X Then
        ws.Columns(Col).Cut
        ws.Columns(X).Insert Shift:=xlToRight
    End If
    X = X + 1

'Part 3: Assign Variables to all of the new column locations'

    Dim ISBN As Long
    ISBN = ws.Rows(1).Find("ISBN").Column

    Dim SKU As Long
    SKU = ws.Rows(1).Find("SKU").Column

    Dim Title As Long
    Title = ws.Rows(1).Find("Title").Column

    Dim NUM As Long
    NUM = ws.Rows(1).Find("Not Using Materials").Column

    Dim Author As Long
    Author = ws.Rows(1).Find("Author").Column

    Dim Section As Long
    Section = ws.Rows(1).Find("Section").Column

    Dim Edition As Long
    Edition = ws.Rows(1).Find("Edition").Column

    Dim RType As Long
    RType = ws.Rows(1).Find("Type").Column

    Dim Instructor As Long
    Instructor = ws.Rows(1).Find("Instructor").Column

'Part 4: Delete a bunch of stuff that we are never going to have in our Alma'

    'lastRow'
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    'Use 1 for loop to delete the row if it contains any of the following:'
    'Deleting the row would change the row number, and multiple "next i" statements are invalid, so we are going to clear contents instead'
    For i = 2 To lastRow
        If ws.Cells(i, Edition).Value = "Not Available" Then
            ws.Cells(i, Edition).Value = "N/A"
        End If
        If ws.Cells(i, RType).Value = "RECOMMENDED" Then
            ws.Cells(i, RType).Value = "Recom."
        ElseIf ws.Cells(i, RType).Value = "REQUIRED" Then
            ws.Cells(i, RType).Value = "Required"
        ElseIf ws.Cells(i, RType).Value = "Not Available" Then
            ws.Cells(i, RType).Value = "N/A"
        End If
        If Left(ws.Cells(i, ISBN).Value, "3") = "281" Then
            ws.Rows(i).ClearContents
        ElseIf ws.Cells(i, NUM).Value = "Yes" Then
            ws.Rows(i).ClearContents
        ElseIf InStr(1, ws.Cells(i, Title), "FROM INQUIRY") > 0 Then
            ws.Rows(i).ClearContents
        ElseIf InStr(1, ws.Cells(i, Title), "CENGAGE UNLIMITED") > 0 Then
            ws.Rows(i).ClearContents
        ElseIf InStr(1, LCase(ws.Cells(i, Title)), "binghamton writes") > 0 Then
            ws.Rows(i).ClearContents
        ElseIf InStr(1, LCase(ws.Cells(i, Title)), "student lab notebook") > 0 Then
            ws.Rows(i).ClearContents
        ElseIf InStr(1, LCase(ws.Cells(i, Author)), "iclicker") > 0 Then
            ws.Rows(i).ClearContents
        ElseIf InStr(1, LCase(ws.Cells(i, Author)), "i-clicker") > 0 Then
            ws.Rows(i).ClearContents
        ElseIf Left(ws.Cells(i, Title).Value, "19") = "DAVVERO 1: 12-MONTH" Then
            ws.Rows(i).ClearContents
        ElseIf Left(ws.Cells(i, Title).Value, "18") = "DAVVERO 1: 6-MONTH" Then
            ws.Rows(i).ClearContents
        Else
        End If
    Next i

    'Sort the data so that it is usable'

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(Author), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Columns(Section), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Columns(ISBN), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A:BB")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Part 5: Clean up the data'

    'Delete Columns: Store, Not Using Materials, Campus, Publisher, Condition, OER, OER+'
        If ws.Cells(1, 1).Value = "Store" Then
        ws.Columns(1).EntireColumn.Delete
        End If

        Col = ws.Rows(1).Find("Not Using Materials").Column
        ws.Columns(Col).EntireColumn.Delete

        Col = ws.Rows(1).Find("Campus").Column
        ws.Columns(Col).EntireColumn.Delete

        Col = ws.Rows(1).Find("Publisher").Column
        ws.Columns(Col).EntireColumn.Delete

        Col = ws.Rows(1).Find("Condition").Column
        ws.Columns(Col).EntireColumn.Delete

        Col = ws.Rows(1).Find("OER").Column
        ws.Columns(Col).EntireColumn.Delete
        
        Col = ws.Rows(1).Find("OER+").Column
        ws.Columns(Col).EntireColumn.Delete

    'Fix ISBN and SKU'
    ISBN = ws.Rows(1).Find("ISBN").Column
    Columns(ISBN).NumberFormat = "0"
    SKU = ws.Rows(1).Find("SKU").Column
    Columns(SKU).NumberFormat = "0"

    'Hide columns we want to keep but not see (School, Department, Course Code, Course Title, SKU, Digital Available, Note for Bookstore, Price at Time of Adoption, Submitter, Submitter Email, Date Submitted)'
    'Selection.EntireColumn.Hidden = True'

    Dim School As Long
    School = ws.Rows(1).Find("School").Column
    ws.Columns(School).Select
    Selection.EntireColumn.Hidden = True

    Dim Dept As Long
    Dept = ws.Rows(1).Find("Department").Column
    ws.Columns(Dept).Select
    Selection.EntireColumn.Hidden = True

    Dim CourseCode As Long
    CourseCode = ws.Rows(1).Find("Course Code").Column
    ws.Columns(CourseCode).Select
    Selection.EntireColumn.Hidden = True

    Dim CourseN As Long
    CourseN = ws.Rows(1).Find("Course Name").Column
    ws.Columns(CourseN).Select
    Selection.EntireColumn.Hidden = True

    ws.Columns(SKU).Select
    Selection.EntireColumn.Hidden = True

    Dim Digital As Long
    Digital = ws.Rows(1).Find("Digital Available").Column
    ws.Columns(Digital).Select
    Selection.EntireColumn.Hidden = True

    Dim NoteB As Long
    NoteB = ws.Rows(1).Find("Note for Bookstore").Column
    ws.Columns(NoteB).Select
    Selection.EntireColumn.Hidden = True

    Dim Price As Long
    Price = ws.Rows(1).Find("Price at Time of Adoption").Column
    ws.Columns(Price).Select
    Selection.EntireColumn.Hidden = True

    Dim Submitter As Long
    Submitter = ws.Rows(1).Find("Submitter").Column
    ws.Columns(Submitter).Select
    Selection.EntireColumn.Hidden = True

    Dim SubmitterE As Long
    SubmitterE = ws.Rows(1).Find("Submitter Email").Column
    ws.Columns(SubmitterE).Select
    Selection.EntireColumn.Hidden = True

    Dim DateS As Long
    DateS = ws.Rows(1).Find("Date Submitted").Column
    ws.Columns(DateS).Select
    Selection.EntireColumn.Hidden = True

        'Selection.EntireColumn.Hidden = True'

'Part 6: Format the sheet for viewing and printing'
    ISBN = ws.Rows(1).Find("ISBN").Column
    Title = ws.Rows(1).Find("Title").Column
    Author = ws.Rows(1).Find("Author").Column
    Section = ws.Rows(1).Find("Section").Column
    Edition = ws.Rows(1).Find("Edition").Column
    RType = ws.Rows(1).Find("Type").Column
    Instructor = ws.Rows(1).Find("Instructor").Column

    'Make it a table'
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes).TableStyle = "TableStyleLight18"

    'Autofit stuff'
    ws.Columns(Title).ColumnWidth = 42
    ws.Columns(Author).ColumnWidth = 28
    ws.Columns(Edition).ColumnWidth = 8
    ws.Columns(RType).ColumnWidth = 7
    ws.Columns(ISBN).ColumnWidth = 13.2
    ws.Columns(Instructor).ColumnWidth = 17
    ws.Columns(Section).ColumnWidth = 12.86
    'Font size'
    'Center, left, right stuff'
    ws.Columns(Edition).HorizontalAlignment = xlCenter

    'This sheet is 5,000 items long at this point (down from 7,000!) so I do not format it for printing'

End Sub