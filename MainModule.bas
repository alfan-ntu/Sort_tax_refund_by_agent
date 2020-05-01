Attribute VB_Name = "MainModule"
'
' Core utilities of this application
' DO NOT DELETE THIS MODULE!
'
' Coding Date: 2020/4/25
'
Option Explicit
    Const summary_sheet_name_constant As String = "彙總"
    Const master_sheet As String = "退稅明細表"
'
' split the master worksheet by agents (column C)
' Macro: Ctrl-Shift-S
'
Sub split_master_by_agent()
Attribute split_master_by_agent.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim lr, target_ws_lr As Long
    Dim master_ws, target_ws As Worksheet
    Dim vcol, i, j As Integer
    Dim icol As Long
    Dim myarr As Variant
    Dim title As String
    Dim titlerow As Integer
    Dim target_range As Range
    Dim number_of_agent, total_record_count As Integer
    total_record_count = 0
    number_of_agent = 0

    'This macro splits data into multiple worksheets based on the variables on a column found in Excel.
    'An InputBox asks you which columns you'd like to filter by, and it just creates these worksheets.

    ' data column of agency
    
    Set master_ws = Sheets(master_sheet)
    master_ws.Activate
    Application.ScreenUpdating = False
    vcol = 3
    
    Set master_ws = ActiveSheet
    lr = master_ws.Cells(master_ws.Rows.Count, vcol).End(xlUp).Row
    title = "A1"
    titlerow = master_ws.Range(title).Cells(1).Row
    icol = master_ws.Columns.Count
    master_ws.Cells(1, icol) = "經銷商"
    For i = 2 To lr
        On Error Resume Next
        If master_ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(Trim(master_ws.Cells(i, vcol)), master_ws.Columns(icol), 0) = 0 Then
            master_ws.Cells(master_ws.Rows.Count, icol).End(xlUp).Offset(1) = master_ws.Cells(i, vcol)
        End If
    Next

    myarr = Application.WorksheetFunction.Transpose(master_ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    '
    ' create a new summary worksheet and copy the agent column to the summary worksheet @ Range("B2")
    '
    Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = summary_sheet_name_constant
    Set target_ws = Sheets(summary_sheet_name_constant)
    master_ws.Columns(icol).Copy (target_ws.Columns(2))
    Call SupportModule.middle_align(target_ws.Range("A1:C1"))
    Call SupportModule.middle_align(target_ws.Columns(1))
    target_ws.Range("A1").Value = "項次"
    target_ws.Range("C1").Value = "退稅件數"
    target_ws.Columns(2).AutoFit
    master_ws.Columns(icol).Clear
    '
    ' disable wrap-text, set auto-fit to ensure proper width of all columns in split sheets
    '
    master_ws.Activate
    Set target_range = master_ws.Range("A:I")
    Call set_cell_autofit(target_range)
    Set target_range = master_ws.Range("A1")
    target_range.Select

    '
    ' Traverse myarr(), create correspondent worksheet
    '
    number_of_agent = UBound(myarr) - 1

    For i = 2 To UBound(myarr)
        master_ws.Range(title).AutoFilter field:=vcol, Criteria1:=myarr(i) & ""
        If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
            Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = myarr(i) & ""
        Else
            Sheets(myarr(i) & "").Move after:=Worksheets(Worksheets.Count)
        End If
        '
        ' Copy the entire rows
        '
        master_ws.Range("A" & titlerow & ":A" & lr).EntireRow.Copy Sheets(myarr(i) & "").Range("A1")
        '
        ' delete unnecessary columns
        '
        Set target_ws = Sheets(myarr(i) & "")
        target_ws.Range("F:H").EntireColumn.Delete
        target_ws.Range("G:S").EntireColumn.Delete
        target_ws.Columns.AutoFit
        target_ws_lr = target_ws.Cells(target_ws.Rows.Count, 1).End(xlUp).Row
        '
        ' indexing the first column
        '
        For j = 2 To target_ws_lr
            target_ws.Range("A" & j).Value = j - 1
        Next
        
        '
        ' add new column for tax-refund amount
        '
        Set target_range = target_ws.Range("G2:G" & target_ws_lr)
        target_range.Value = "50,000"
        target_ws.Range("G1").Value = "金額"
        target_ws.Range("G1").HorizontalAlignment = xlCenter
        Set target_range = target_ws.Range("G1:G" & target_ws_lr)
        Call SupportModule.cell_border(target_range)
        
        Debug.Print myarr(i) & " record# " & target_ws_lr - 1 & " / total: " & 5000 * (target_ws_lr - 1)
        
        target_ws.Range("G" & target_ws_lr + 1).Value2 = 5000 * (target_ws_lr - 1)
        target_ws.Range("G" & target_ws_lr + 1).Select
        Selection.NumberFormatLocal = "#,##0"
        
        Set target_range = target_ws.Range("A" & 1 & ":G" & target_ws_lr + 1)
        Call SupportModule.cell_border(target_range)
        ' Call SupportModule.set_font_jhenghei(target_range)
        Call SupportModule.set_font(target_range, "新細明體", 12)
        Set target_range = target_ws.Cells
        Call SupportModule.unfill_cells(target_range)
        target_ws.Range("A1").Select
        '
        ' update info in summary worksheet 彙總表
        '
        Set target_ws = Sheets(summary_sheet_name_constant)
        target_ws.Activate
        target_ws.Cells(i, 2).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
            myarr(i) & "!A1", TextToDisplay:=myarr(i)
        Set target_range = target_ws.Cells(i, 3)
        target_range.Value = target_ws_lr - 1
        total_record_count = total_record_count + target_ws_lr - 1
        Set target_range = target_ws.Cells(i, 1)
        target_range.Value = i - 1
    Next
    
    Set target_ws = Sheets(summary_sheet_name_constant)
    target_ws.Activate
    Set target_range = target_ws.Cells(UBound(myarr) + 1, 2)
    target_range.Value = "總件數"
    Call SupportModule.top_cell_border(target_range)
    
    Set target_range = target_ws.Cells(UBound(myarr) + 1, 3)
    target_range.Value = total_record_count
    Call SupportModule.top_cell_border(target_range)
    
    Unload UserForm_Progress
        
    master_ws.AutoFilterMode = False
    master_ws.Activate
    Application.ScreenUpdating = True
End Sub
'
' Delete worksheets of names specified in the first column
' Macro: Ctrl-Shift-D
'
Sub delete_worksheets()
Attribute delete_worksheets.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' delete_work_sheet Macro
'
' Keyboard Shortcut: Ctrl+Shift+D
'
    Dim lr As Long
    Dim ws As Worksheet
    Dim vcol, i As Integer
    Dim icol As Long
    Dim myarr As Variant
    Dim title As String
    Dim titlerow As Integer

    'This macro splits data into multiple worksheets based on the variables on a column found in Excel.
    'An InputBox asks you which columns you'd like to filter by, and it just creates these worksheets.
    vcol = 3
    Set ws = Sheets(master_sheet)
    ws.Activate
    Application.ScreenUpdating = False
    
    lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row
    title = "A1"
    titlerow = ws.Range(title).Cells(1).Row
    icol = ws.Columns.Count
    ws.Cells(1, icol) = "Unique"
    For i = 2 To lr
        On Error Resume Next
        If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
            ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)

            Debug.Print ws.Cells(i, vcol)
        End If
    Next
    
    myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    ws.Columns(icol).Clear
    Application.DisplayAlerts = False
    Sheets(summary_sheet_name_constant).Delete
    Application.DisplayAlerts = True
    For i = 2 To UBound(myarr)
        ws.Range(title).AutoFilter field:=vcol, Criteria1:=myarr(i) & ""
        If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
            Debug.Print myarr(i) & " has been dedeleted"
            ' Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = myarr(i) & ""
        Else
            ' Sheets(myarr(i) & "").Move after:=Worksheets(Worksheets.Count)
            Application.DisplayAlerts = False
            Sheets(myarr(i) & "").Delete
            Application.DisplayAlerts = True
        End If
        'Sheets(myarr(i) & "").Columns.AutoFit
    Next

    ws.AutoFilterMode = False
    ws.Activate
    Application.ScreenUpdating = True
End Sub
'
' Create worksheets with names listed in the column#1
'
Sub create_worksheets()
    Dim lr As Long
    Dim ws As Worksheet
    Dim vcol, i As Integer
    Dim icol As Long
    Dim myarr As Variant
    Dim title As String
    Dim titlerow As Integer

    'This macro splits data into multiple worksheets based on the variables on a column found in Excel.
    'An InputBox asks you which columns you'd like to filter by, and it just creates these worksheets.

    Application.ScreenUpdating = False
    vcol = Application.InputBox(prompt:="Which column would you like to filter by?", title:="Filter column", Default:="2", Type:=1)
    Set ws = ActiveSheet
    lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row
    title = "A1"
    titlerow = ws.Range(title).Cells(1).Row
    icol = ws.Columns.Count
    ws.Cells(1, icol) = "Unique"
    For i = 2 To lr
        On Error Resume Next
        If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
            ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)

            Debug.Print ws.Cells(i, vcol)
        End If
    Next

    myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
    ws.Columns(icol).Clear


    For i = 2 To UBound(myarr)
        ws.Range(title).AutoFilter field:=vcol, Criteria1:=myarr(i) & ""
        If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
            Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = myarr(i) & ""
        Else
            Sheets(myarr(i) & "").Move after:=Worksheets(Worksheets.Count)
        End If
    Next

    ws.AutoFilterMode = False
    ws.Activate
    Application.ScreenUpdating = True

End Sub



Sub cell_border(border_range As Range)
'
' cell_border Macro
'
' Keyboard Shortcut: Ctrl+Shift+B
'
    border_range.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
'
' turns off auto wrapping and set cell width to auto_fit
'
Sub set_cell_autofit(cell_range As Range)
    cell_range.Select
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    cell_range.EntireColumn.AutoFit
End Sub
'
'
'
Sub unselect_range()
    Dim ws As Worksheet
    Dim tr As Range
    
    Set ws = ActiveSheet
    'Set tr = ws.Range("A1")
    'tr.Select
        
    Set tr = ws.Cells(ws.Rows.Count, 3).End(xlUp)
    tr.Select
End Sub
'
'
'
Sub draw_top_border_over_range(target_range As Range)
    Dim ws As Worksheet
'    Dim tr As Range
    
'    Set ws = Sheets("彙總")
'    Set tr = ws.Range("A14:C14")
'    tr.Select
'    Worksheets("彙總").Activate
'    Set ws = ActiveSheet
'    Set tr = ws.Range("A14:C14")
    target_range.Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

