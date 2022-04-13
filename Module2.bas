Attribute VB_Name = "Module2"
Sub CleanupData()

' Cleanup_Data Macro
' Adjusts column widths and bolds column headers.

    Columns("B:B").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Columns("B:AX").EntireColumn.AutoFit
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
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Rows(1).Insert
    Rows(1).Insert
    Rows(1).Insert
    Rows(1).Insert
End Sub

