Attribute VB_Name = "Module8"
Sub Constraintz()
Attribute Constraintz.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Constraintz Macro
'
    ThisWorkbook.Worksheets("Solver_Blackbox").Activate
    
    Range("E40").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-31]C:R[-2]C)"
    Selection.AutoFill Destination:=Range("E40:AW40"), Type:=xlFillDefault
    Range("E40:AW40").Select
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("E41").Select
    ActiveCell.FormulaR1C1 = "'="
    Range("E41").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.AutoFill Destination:=Range("E41:AW41"), Type:=xlFillDefault
    Range("E41:AW41").Select
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("E42").Select
    ActiveCell.FormulaR1C1 = 1
    Selection.AutoFill Destination:=Range("E42:AW42"), Type:=xlFillDefault
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("D45").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(Raw_Data!R[-38]C[1]:R[-9]C[45], R[-36]C[1]:R[-7]C[45])"
    Range("E40").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("F40").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("E42").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=-45
    
    ConstraintPt2
    
    
End Sub
Sub ConstraintPt2()
Attribute ConstraintPt2.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("AY9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.AutoFill Destination:=Range("AY9:AY38")
    Range("AY9:AY38").Select
    ActiveWindow.SmallScroll Down:=0
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("AX9").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=3
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("AZ9").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-47]:RC[-3])"
    Range("AZ9").Select
    Selection.AutoFill Destination:=Range("AZ9:AZ38")
    Range("AZ9:AZ38").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("AY9").Select
    Selection.Copy
    Range("BA9").Select
    ActiveSheet.Paste
    Range("BA9").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("BA9:BA38")
    Range("BA9:BA38").Select
    Range("AY9").Select
    ActiveCell.FormulaR1C1 = ">="
    Range("AY9").Select
    Selection.AutoFill Destination:=Range("AY9:AY38")
    Range("AY9:AY38").Select
    Range("BA9").Select
    ActiveCell.FormulaR1C1 = ">="
    Range("BA9").Select
    Selection.AutoFill Destination:=Range("BA9:BA38")
    Range("BA9:BA38").Select
    Range("BB9").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("BB9").Select
    Selection.AutoFill Destination:=Range("BB9:BB38")
    Range("BB9:BB38").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
End Sub
