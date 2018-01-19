Attribute VB_Name = "CostEvaluation"
Sub CostEvaluation()
    Call TextToColumns
    Call OrderData
    Call DeleteJunk
    Call FormatAndReplace
End Sub

Private Sub TextToColumns()
Attribute TextToColumns.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.CutCopyMode = False
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(13, 1), Array(53, 1), Array(56, 1), Array(72, 1), _
        Array(80, 1), Array(89, 1), Array(102, 1), Array(118, 1), Array(129, 1), Array(142, 1), _
        Array(154, 1), Array(167, 1), Array(183, 1)), TrailingMinusNumbers:=True
    Selection.ColumnWidth = 20.78
    Selection.ColumnWidth = 13.89
    Columns("B:E").Select
    Selection.ColumnWidth = 7.89
    Columns("B:B").Select
    Selection.ColumnWidth = 29
    Selection.ColumnWidth = 46.22
    Selection.ColumnWidth = 40.22
    Columns("F:N").Select
    Selection.ColumnWidth = 12.11
    Selection.ColumnWidth = 9.78
    Columns("K:K").Select
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    ActiveWindow.FreezePanes = True
    Range("A11").Select
    Rows("10:10").Select
    Selection.Font.Bold = True
    Range("A11").Select
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("M10").Select
    ActiveCell.FormulaR1C1 = "Cost"
    Range("N10").Select
    ActiveCell.FormulaR1C1 = "Plant"
    Range("B12").Select
    ActiveWindow.ScrollColumn = 1
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "ost"
    Range("B9").Select
    ActiveCell.FormulaR1C1 = ""
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "Cost"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "By: JMEEHAN @ QPADEV002R"
    Range("A5").Select
End Sub
Private Sub OrderData()
Attribute OrderData.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A11:N99999").Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "A11:A99999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A11:N999999")
        .Header = xlGuess
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
        Firstrow = 11
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = LastRow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")
                If .Value2 = "Part #" Then         'Delete rows with "Part Number"
                .EntireRow.Delete
                ElseIf .Value2 Like "=*" Then           'Delete ===== spacer rows
                .EntireRow.Delete
                ElseIf .Value2 Like "By:*" Then    'Delete "JMEEHAN @ REPORT" rows
                .EntireRow.Delete
                ElseIf .Value2 Like "Report:*" Then          'Delete Report rows
                .EntireRow.Delete
                ElseIf .Value2 Like "Date:*" Then          'Delete Date report rows
                .EntireRow.Delete
                ElseIf .Value2 Like "S0*" Then          'Delete S0 parts rows
                .EntireRow.Delete
                ElseIf .Value2 Like "T*" Then          'Delete TXXXXX rows
                .EntireRow.Delete
                End If
            End With
        Next Lrow
        
        On Error Resume Next
        Range("A11:A99999").SpecialCells(xlCellTypeBlanks).EntireRow.Delete 'Remove blank rows

    End With

    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = False
        .Calculation = CalcMode
    End With

End Sub
Private Sub FormatAndReplace()
Attribute FormatAndReplace.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("J:K,M:M").Select
    Range("M1").Activate
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Range("K11").Select
    Columns("F:I").Select
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Range("A11").Select
    Columns("D:D").Select
    Range("A11").Select
    
    Cells.Select
    Selection.Replace What:="*- N/A -*", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A11").Select
End Sub
