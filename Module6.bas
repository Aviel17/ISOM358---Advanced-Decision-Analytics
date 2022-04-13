Attribute VB_Name = "Module6"
Sub ConvertDataToOutput()
Attribute ConvertDataToOutput.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ConvertDataToOutput Macro
'

    Range("E9").Select
    ActiveCell.FormulaR1C1 = "=RANDBETWEEN(0,1)"
    Selection.AutoFill Destination:=Range("E9:AW9"), Type:=xlFillDefault
    Range("E9:AW9").Select
    Selection.AutoFill Destination:=Range("E9:AW38")
    Range("E9:AW38").Select
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    Columns("B:D").AutoFit
    
End Sub
