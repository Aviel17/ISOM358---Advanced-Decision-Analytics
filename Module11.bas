Attribute VB_Name = "Module11"
Sub Decision_Variables()
Attribute Decision_Variables.VB_Description = "Highlights all decision variables in Solver_Results."
Attribute Decision_Variables.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Decision_Variables Macro
' Highlights all decision variables in Solver_Results.
'
    ThisWorkbook.Sheets("Solver_Results").Activate
    
    Range("E9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("E9:AW9").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=6
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=-6
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.SmallScroll Down:=33
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
    ActiveWindow.SmallScroll Down:=0
    Range("C45").Select
    ActiveCell.FormulaR1C1 = "Schedule Utility Score"
    Range("AZ8").Select
    ActiveCell.FormulaR1C1 = "Actual Hours Assigned"
    Range("BB8").Select
    ActiveCell.FormulaR1C1 = "Minimum Slots to Work"
    Range("BB9").Select
    Columns("AX:AX").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    Columns("AZ:AZ").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 44
    Columns("BB:BB").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.SmallScroll Down:=39
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
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
    ActiveWindow.SmallScroll Down:=-33

    Range("D42").Select
    ActiveCell.FormulaR1C1 = "Max. amount of volunteers per slot "
    Range("D40").Select
    ActiveCell.FormulaR1C1 = "# of volunteers in slot"
    
    Columns("D:D").EntireColumn.AutoFit
    
End Sub
