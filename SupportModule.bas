Attribute VB_Name = "SupportModule"
'
' Support utilities of this application, including
'   1. cell_border(border_range As Range): set solid border to the assigned range
'   2. middle_align(target_range As Range): set cells in the assigned range to middle
'   3. top_cell_border(target_range As Range): set solid top border of the assigned range
'   4. set_font_jhenghei(target_range As Range): set the font in the assigned range to JhengHei(·L³n¥¿¶Â)
'
' ==== DO NOT DELETE THIS MODULE! ====
' This module is insensistive to the application context and can be carried to other Excel VBA application
'
' Coding Date: 2020/4/25
'
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


Sub middle_align(target_range As Range)
Attribute middle_align.VB_ProcData.VB_Invoke_Func = " \n14"
'
' middle_align Macro
'
'
    target_range.Select
    With Selection
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
'
'
'

Sub top_cell_border(target_range As Range)
'
' top_cell_border Macro
'
    target_range.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
'
'
'
Sub set_font_jhenghei(target_range As Range)
'
' set_font_macro Macro
'
'
    target_range.Select
    With Selection.Font
        .Name = "Microsoft JhengHei"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
'
'
'
Sub set_font(target_range As Range, font_name As String, font_size As Integer)
'
' set_font_macro Macro
'
'
    target_range.Select
    With Selection.Font
        .Name = font_name
        .Size = font_size
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

End Sub
'
'
'
Sub unfill_cells(target_range As Range)
    target_range.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
'
' Sub code() demonstrates the use of progress bar. It passes a normalized progress percentage to progress_bar()
'
Sub code()
'    Dim i As Integer, j As Integer, pctCompl As Single

'    Sheet1.Cells.Clear
    
'    For i = 1 To 100
'        For j = 1 To 50
'            Cells(i, 1).Value = j
'        Next j
'        pctCompl = i
'        progress_bar pctCompl
'    Next i
'    i = MsgBox("Program Completes", vbOKOnly)
'    Unload UserForm_Progress
    Call MainModule.split_master_by_agent
    
End Sub

'
' UserForm_Progress is a user form containing
' 1. a Text control whose caption displays the percentage completes
' 2. a ProgressLabel label control whose width correspondent w/ the percentage comples
'
Sub progress_bar(pctCompl As Single)

    UserForm_Progress.Text.Caption = pctCompl & "% Completed"
    UserForm_Progress.ProgressLabel.Width = pctCompl * 2
    Debug.Print "Progress: " & pctCompl
    DoEvents

End Sub

