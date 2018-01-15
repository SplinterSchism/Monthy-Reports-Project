Attribute VB_Name = "Module11"
Sub TexttoColumns()
Attribute TexttoColumns.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TexttoColumns Macro
'

'
    Columns("A:A").Select
    Selection.TexttoColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(4, 1), Array(9, 1), Array(30, 1), Array(61, 1), _
        Array(64, 1), Array(82, 1), Array(100, 1), Array(120, 1), Array(134, 1), Array(150, 1), _
        Array(166, 1), Array(183, 1)), TrailingMinusNumbers:=True
    ActiveWindow.SmallScroll Down:=-3
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Range("A7").Select
    Columns("I:I").ColumnWidth = 13.33
    Columns("H:H").ColumnWidth = 18.22
    Columns("G:G").ColumnWidth = 15
    Columns("F:F").ColumnWidth = 19.56
    Columns("E:E").ColumnWidth = 17.89
    Columns("D:D").ColumnWidth = 14.78
    Columns("C:C").ColumnWidth = 14.67
    Columns("B:B").ColumnWidth = 31.67
    Columns("A:A").ColumnWidth = 14
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=-33
End Sub

Sub OrderData()
Attribute OrderData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' OrderData Macro
'

'
    Rows("9:9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Rows("7:7").Select
    Selection.Font.Bold = True
    Rows("8:8").Select
    ActiveWindow.FreezePanes = True 'Freeze top pane
    ActiveWindow.SmallScroll Down:=-6
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add _
        Key:=Range("A8:A999999"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A8:K999999")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=-3
End Sub

Sub DeleteJunk()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim LastrowD As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = 8
        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")
                If .Value2 = "Part Number" Then         'Delete rows with "Part Number"
                .EntireRow.Delete
                ElseIf .Value2 Like "=*" Then           'Delete ===== spacer rows
                .EntireRow.Delete
                ElseIf .Value2 Like "*@ REPORT" Then    'Delete "JMEEHAN @ REPORT" rows
                .EntireRow.Delete
                ElseIf .Value2 Like "IV*" Then          'Delete IV55R report row
                .EntireRow.Delete
                ElseIf .Value2 Like "S0*" Then          'Delete S0XXXX parts row
                .EntireRow.Delete
                ElseIf .Value2 Like "P0*" Then          'Delete P0XXXX parts row
                .EntireRow.Delete
                End If
            End With
        Next Lrow
        
        Firstrow = 8
        Lastrow = .Cells(.Rows.Count, "D").End(xlUp).Row

        For Lrow = Lastrow To Firstrow Step -1

            With .Cells(Lrow, "D")
                If .Value Like "For*" Then              'Delete empty lines
                .EntireRow.Delete
                ElseIf .Value Like "-*" Then            'Delete empty lines
                .EntireRow.Delete
                ElseIf .Value Like "Major*" Then        'Delete Time
                .EntireRow.Delete
                ElseIf .Value2 Like "Plant:*" Then      'Delete Dates
                .EntireRow.Delete
                End If
            End With

        Next Lrow

    End With

    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = False
        .Calculation = CalcMode
    End With

End Sub

Sub SortByPart()
'
' SortByPart Macro
'

'
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim partRange As Range
    Dim currentCell As Range
    
    With Application
        .ScreenUpdating = False
    End With
    
    Rows("8:8").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("A8:A2311" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A8:K2311")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A8").Select
    
    Dim cell As Range
    Dim NewRange As Range
    Dim MyCount As Long
    MyCount = 1
    For Each cell In Range("A8:A99999")
        If cell.Value Like "8*" Then
        If MyCount = 1 Then Set NewRange = cell
            Set NewRange = Application.Union(NewRange, cell)
            MyCount = MyCount + 1
        End If
    Next cell
    NewRange.EntireRow.Activate
    Selection.cut
    Rows("8:8").Select
    Selection.Insert Shift:=xlDown

End Sub


