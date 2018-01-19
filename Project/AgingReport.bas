Attribute VB_Name = "AgingReport"
Sub AgingReport()
    Call TextToColumns
    Call OrderData
    Call DeleteJunk
    Call SortByPart
    Call HeaderTotals
    Call AddSheets
    Call MoveRM
    Call MoveInserts
    Call MoveCinserts
    Call MoveMS
    Call HeaderSetup
    Call InsertOrder
End Sub


Private Sub TextToColumns()
Attribute TextToColumns.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TexttoColumns Macro
'

'
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
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

Private Sub OrderData()
Attribute OrderData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' OrderData Macro
'

'
    Columns("A:A").Select
    Rows("9:9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Rows("7:7").Select
    Selection.Font.Bold = True
    Rows("8:8").Select
    ActiveWindow.FreezePanes = True 'Freeze top header pane
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
End Sub

Private Sub DeleteJunk()
    Dim Firstrow As Long
    Dim LastRow As Long
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
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = LastRow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")
                If .Value2 = "Part Number" Then         'Delete rows with "Part Number"
                .EntireRow.Delete
                ElseIf .Value2 Like "=*" Then           'Delete ===== spacer rows
                .EntireRow.Delete
                ElseIf .Value2 Like "*@*" Then    'Delete "JMEEHAN @ REPORT" rows
                .EntireRow.Delete
                ElseIf .Value2 Like "IV*" Then          'Delete IV55R report rows
                .EntireRow.Delete
                ElseIf .Value2 Like "S0*" Then          'Delete S0XXXX parts rows
                .EntireRow.Delete
                ElseIf .Value2 Like "P0*" Then          'Delete P0XXXX parts rows
                .EntireRow.Delete
                ElseIf .Value2 Like "R2*" Then          'Delete Resale parts rows
                .EntireRow.Delete
                End If
            End With
        Next Lrow
        
        On Error Resume Next
        Range("A11:A99999").SpecialCells(xlCellTypeBlanks).EntireRow.Delete 'Remove blank rows
        
        Firstrow = 8
        LastRow = .Cells(.Rows.Count, "D").End(xlUp).Row

        For Lrow = LastRow To Firstrow Step -1

            With .Cells(Lrow, "D")
                If .Value Like "Major*" Then        'Delete Time
                .EntireRow.Delete
                ElseIf .Value2 Like "Plant:*" Then      'Delete Dates
                .EntireRow.Delete
                End If
            End With

        Next Lrow
        
        Firstrow = 8
        LastRow = .Cells(.Rows.Count, "B").End(xlUp).Row

        For Lrow = LastRow To Firstrow Step -1

            With .Cells(Lrow, "B")
                If .Value Like "Report*" Then              'Delete Totals at the bottom
                .EntireRow.Delete
                ElseIf .Value Like "CARTON*" Then              'Delete Cartons parts
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

Private Sub SortByPart()
'
' SortByPart Macro
'

'
    Dim Firstrow As Long
    Dim LastRow As Long
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

Private Sub HeaderTotals()
'
' HeaderTotals Macro
'

'
    With Application
        .ScreenUpdating = False
    End With
    
    Range("C1:J5").Select
    Range("J5").Activate
    Selection.ClearContents
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=SUM"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2306]C)"
    Range("D6").Select

    Range("E5").Select
    ActiveWindow.SmallScroll Down:=-9
    ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2306]C)"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2306]C)"
    Range("F5").Select
    Selection.ClearContents
    Range("E5").Select
    Selection.ClearContents
    Range("D5").Select
    Range("D5").Select
    Selection.cut Destination:=Range("E5")
    Range("E5").Select
    Selection.cut Destination:=Range("D5")
    Range("D5").Select
    Selection.AutoFill Destination:=Range("D5:E5"), Type:=xlFillDefault
    Range("D5:E5").Select
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2306]C)"
    Range("E5").Select
    Selection.AutoFill Destination:=Range("E5:F5"), Type:=xlFillDefault
    Range("E5:F5").Select
    Selection.AutoFill Destination:=Range("E5:I5"), Type:=xlFillDefault
    Range("E5:I5").Select
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2306]C)"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2306]C)"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2306]C)"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2306]C)"
    Range("D4").Select
    
    Cells.Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
    Columns("A:A").Select
    Selection.NumberFormat = "General"
    Range("D7").Select
    Selection.NumberFormat = "@"
    ActiveCell.FormulaR1C1 = "1 - 30"
    Range("D8").Select
End Sub

Private Sub AddSheets()

    With Application
        .ScreenUpdating = False
    End With
    
    Dim sheetNames As Variant
    sheetNames = Array("MS", "Inserts", "RM")
    ActiveSheet.Name = "FG" 'Rename first sheet to FG
    'Step 1: Tell Excel what to do if error
        On Error GoTo MyError
    'Step 2:  Add a sheet and name it
    For Each i In sheetNames
        Sheets.Add After:=Sheets(ActiveWorkbook.Sheets.Count)
        ActiveSheet.Name = i
    Next i
    Sheets("FG").Select
    Range("A8").Select
    
        Exit Sub
    'Step 3: If here, an error happened; tell the user
MyError:
        MsgBox "There is already a sheet called that."
        

End Sub

Private Sub MoveMS()
    With Application
        .ScreenUpdating = False
    End With
    
    Dim cell As Range
    Dim NewRange As Range
    Dim MyCount As Long
    MyCount = 1
    For Each cell In Range("A8:A99999")
        If cell.Value Like "7*" Then
        If MyCount = 1 Then Set NewRange = cell
            Set NewRange = Application.Union(NewRange, cell)
            MyCount = MyCount + 1
        End If
    Next cell
    NewRange.EntireRow.Activate
    Selection.cut
    Sheets("MS").Select
    Rows("8:8").Select
    Selection.Insert Shift:=xlDown
    Sheets("FG").Select
End Sub

Private Sub MoveRM()
    With Application
        .ScreenUpdating = False
    End With
    
    Dim cell As Range
    Dim NewRange As Range
    Dim MyCount As Long
    MyCount = 1
    For Each cell In Range("A8:A99999")
        If cell.Value Like "RM*" Then
        If MyCount = 1 Then Set NewRange = cell
            Set NewRange = Application.Union(NewRange, cell)
            MyCount = MyCount + 1
        End If
    Next cell
    NewRange.EntireRow.Activate
    Selection.cut
    Sheets("RM").Select
    Rows("8:8").Select
    Selection.Insert Shift:=xlDown
    Sheets("FG").Select
End Sub

Private Sub MoveInserts()
    With Application
        .ScreenUpdating = False
    End With
    
    Dim cell As Range
    Dim NewRange As Range
    Dim LastRow As String
    Dim MyCount As Long
    MyCount = 1
    
    For Each cell In Range("A8:A99999")
        If cell.Value Like "I*" Then
        If MyCount = 1 Then Set NewRange = cell
            Set NewRange = Application.Union(NewRange, cell)
            MyCount = MyCount + 1
        End If
    Next cell
    NewRange.EntireRow.Activate
    Selection.cut
    Sheets("Inserts").Select
    Rows("8:8").Select
    Selection.Insert Shift:=xlDown
    Sheets("FG").Select
    
End Sub



Private Sub MoveCinserts()
    With Application
        .ScreenUpdating = False
    End With
    
    Dim cell As Range
    Dim NewRangeC As Range
    Dim LastRow As String
    Dim MyCount As Long
    MyCount = 1
    LastRow = Sheets("Inserts").Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    For Each cell In Range("A8:A99999")
        If cell.Value Like "C*" Then
        If MyCount = 1 Then Set NewRangeC = cell
            Set NewRangeC = Application.Union(NewRangeC, cell)
            MyCount = MyCount + 1
        End If
    Next cell
    NewRangeC.EntireRow.Activate
    Selection.cut
    Sheets("Inserts").Select
    Range("A" & LastRow).Select
    Selection.Insert Shift:=xlDown
    Sheets("FG").Select

End Sub

Private Sub HeaderSetup()
    With Application
        .ScreenUpdating = False
    End With
    
    Dim sheetList As Variant
    sheetList = Array("MS", "Inserts", "RM")

    For Each i In sheetList
        Sheets("FG").Select
        Rows("1:7").Select
        Selection.Copy
        Sheets(i).Select
        Range("A1").Select
        ActiveSheet.Paste
        Rows("8:8").Select
        ActiveWindow.FreezePanes = True     'Freeze top header pane
        
        Range("D5").Select                  'Apply SUM formula
        ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2304]C)"
        Range("E5").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2304]C)"
        Range("F5").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2304]C)"
        Range("G5").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2304]C)"
        Range("H5").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2304]C)"
        Range("I5").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[3]C:R[2304]C)"
        
        Columns("I:I").ColumnWidth = 13.33  'Adjust column width
        Columns("H:H").ColumnWidth = 18.22
        Columns("G:G").ColumnWidth = 15
        Columns("F:F").ColumnWidth = 19.56
        Columns("E:E").ColumnWidth = 17.89
        Columns("D:D").ColumnWidth = 14.78
        Columns("C:C").ColumnWidth = 14.67
        Columns("B:B").ColumnWidth = 31.67
        Columns("A:A").ColumnWidth = 14
    Next i
    
    Sheets("FG").Select

End Sub

Private Sub InsertOrder()
'
' InsertOrder Macro
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Inserts").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Inserts").Sort.SortFields.Add Key:=Range("I8:I63") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Inserts").Sort.SortFields.Add Key:=Range("H8:H63") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Inserts").Sort.SortFields.Add Key:=Range("G8:G63") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Inserts").Sort.SortFields.Add Key:=Range("F8:F63") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Inserts").Sort.SortFields.Add Key:=Range("E8:E63") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Inserts").Sort.SortFields.Add Key:=Range("D8:D63") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Inserts").Sort
        .SetRange Range("A8:K63")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Inserts").Select
    Range("A8").Select
    Range("F8:I45").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A8").Select
    Sheets("FG").Select
End Sub















