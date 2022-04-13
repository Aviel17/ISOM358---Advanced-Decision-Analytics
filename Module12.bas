Attribute VB_Name = "Module12"
Sub Menu_Background()
'
' Background Macro
' Nice touch to Home Menu!
'
'    Range(Selection, Selection.End(xlToLeft)).Select
'    Range("A1:T15").Select
'    ActiveWindow.ScrollColumn = 3
'    ActiveWindow.ScrollColumn = 2
'    ActiveWindow.ScrollColumn = 1
'    Range("A1:T20").Select
'    ActiveWindow.ScrollColumn = 4
'    ActiveWindow.ScrollColumn = 3
'    ActiveWindow.ScrollColumn = 2
'    ActiveWindow.ScrollColumn = 1
    Range("A1:T21").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=-15
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
End Sub
