Attribute VB_Name = "Module10"
Sub BlackBox_Clone()
Attribute BlackBox_Clone.VB_Description = "Makes a new copy of the solver blackbox so that the maximized utility score and decision variables are untouched."
Attribute BlackBox_Clone.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BlackBox_Clone Macro
' Makes a new copy of the solver blackbox so that the maximized utility score and decision variables are untouched.
'
    Sheets("Solver_Blackbox").Select
    Sheets("Solver_Blackbox").Copy Before:=Sheets(3)
    Sheets("Solver_Blackbox (2)").Select
    Sheets("Solver_Blackbox (2)").Name = "Solver_Results"
    
End Sub
