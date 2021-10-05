Sub Macro13()
'
' Macro13 Macro
'
' Keyboard Shortcut: Ctrl+g
'
    Range("B6:AF6").Select
    Selection.Copy
    Range("B9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B7:AF7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Compute").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Sheets("Input").Select
    Range("B9:AF9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Compute").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Output").Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=True, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    Sheets("Compute").Select
    Range("G2:G32").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("B2:B32").Select
    Selection.ClearContents
    Sheets("Input").Select
    Range("B7:AF9").Select
    Selection.ClearContents
    Range("B7:AF7").Select
    With Selection.Interior
        .PatternColor = 0
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Sheets("Output").Select
    'Clear any existing filters
      On Error Resume Next
        Cells.ShowAllData
      On Error GoTo 0

      '1. Apply Filter
      Range("A1:I32").AutoFilter Field:=1, Criteria1:=""

      '2. Delete Rows
      Application.DisplayAlerts = False
        Range("A2:I32").SpecialCells(xlCellTypeVisible).Delete
      Application.DisplayAlerts = True

      '3. Clear Filter
      On Error Resume Next
        Cells.ShowAllData
      On Error GoTo 0

    Dim wbNew As Excel.Workbook
    Dim wsSource As Excel.Worksheet, wsTemp As Excel.Worksheet
    Dim name As String

        Set wsSource = ThisWorkbook.Worksheets(3)
        name = "Shifts"
        user = Environ("Username")
        desktop = "C:\Users\" & user & "\OneDrive - McKinsey & Company\Desktop\"
        Application.DisplayAlerts = False 'will overwrite existing files without asking
        Set wsTemp = ThisWorkbook.Worksheets(3)
        Set wbNew = ActiveWorkbook
        Set wsTemp = wbNew.Worksheets(3)
        wbNew.SaveAs desktop & name & ".csv", xlCSVUTF8 'new way
        wbNew.Close
        Application.DisplayAlerts = True
End Sub
