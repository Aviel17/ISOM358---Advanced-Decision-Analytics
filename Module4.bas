Attribute VB_Name = "Module4"
Sub CreateButton()
    Dim btn As Button
    Application.ScreenUpdating = False
    ActiveSheet.Buttons.Delete
    Dim t As Range
    Set t = ActiveSheet.Range(Cells(2, 3), Cells(4, 3))
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    
    'Customization of button
    btn.Font.ColorIndex = 3
    btn.Font.Bold = True
    btn.Font.Size = 16
    btn.Font.Name = "Times New Roman"
    
    'What the button will do once you press it
    btn.OnAction = "SolverSetup"
    btn.Caption = "Get Best Schedule"
    btn.Name = "Btn"

    Range("A1").Select
    Application.ScreenUpdating = True
    
End Sub

Sub SolverSetup()
Attribute SolverSetup.VB_Description = "Copies all data and sets up constraints to run OpenSolver later on."
Attribute SolverSetup.VB_ProcData.VB_Invoke_Func = " \n14"

    'Copy all data
    Application.ScreenUpdating = False
    
    ThisWorkbook.Sheets.Add(After:=Sheets(2)).Name = "Solver_Blackbox"
    Sheets("Raw_Data").Select
    Range("B5:AX36").Copy
    
    ThisWorkbook.Worksheets("Solver_Blackbox").Range("B7").PasteSpecial xlPasteValues
    
    ThisWorkbook.Worksheets("Solver_Blackbox").Select
    
    'Calls the ConvertDataToOutput function in Module 6
    Application.Run "Module6.ConvertDataToOutput"
    
    'Calls the Constraintz function in Module 8
    Application.Run "Module8.Constraintz"
    
    'Actually run OpenSolver
    Application.Run "Module7.SolvingSolver"
    
    'Switch to Solver_Blackbox to start working on final scheduling
    ThisWorkbook.Sheets.Add(After:=Sheets(3)).Name = "Final_Schedule"
    ThisWorkbook.Worksheets("Solver_Blackbox").Activate
    
    'Create copy of Solver results for troubleshooting (and format as well)
    Application.Run "Module10.BlackBox_Clone"
    Application.Run "Module11.Decision_Variables"
    
    'Switch to the blackbox as a template for creating the final schedule
    ThisWorkbook.Sheets("Solver_Blackbox").Activate
    
    'Create final schedule
    Application.Run "Module9.AddNames"
    
    'Send out emails
    Application.Run "Module14.CreateEmailButton"
    
    Application.ScreenUpdating = True
    
End Sub
