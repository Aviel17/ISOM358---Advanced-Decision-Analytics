Attribute VB_Name = "Module7"
'TODO: Look at PP slides about API stuff on integrating OpenSolver with this method.
'Remember: the button on the Raw Data page should activate this method!
Sub SolvingSolver()

    ThisWorkbook.Worksheets("Solver_Blackbox").Activate

    'This part is the same as the first part of Auto_solve(): We use the
    'built-in solver to set up the optimization model. OpenSolver will automatically
    'copy this model over, so there is no need to set it up there
    Range(Cells(9, 5), Cells(38, 49)).Value = "" 'Just an alternate way of referencing E9:AW38
    SolverReset
    SolverAdd CellRef:="$E$40:$AW$40", Relation:=2, FormulaText:="$E$42:$AW$42" ' Only 1 person per slot
    SolverAdd CellRef:="$E$9:$AW$38", Relation:=5, FormulaText:="binary" ' DVs are binary
    SolverAdd CellRef:="$AZ$9:$AZ$38", Relation:=1, FormulaText:="$AX$9:$AX$38"
    SolverAdd CellRef:="$AZ$9:$AZ$38", Relation:=3, FormulaText:="$BB$9:$BB$38"
    SolverOk SetCell:="$D$45", MaxMinVal:=1, ValueOf:=0, _
    ByChange:=Range(Cells(9, 5), Cells(38, 49)), _
    Engine:=2, EngineDesc:="Simplex LP"
    SetTolerance (0)
    'Run OpenSolver. ResultFlag = 0 means that an optimal solution was found. Visit
    '
    '   opensolver.org/opensolver-api-reference/#RunOpenSolver
    '
    'and look under the section 'OpenSolverResult' for a description of all possible return values
    
    ResultFlag = Application.Run("OpenSolver.xlam!RunOpenSolver")
    If ResultFlag <> 0 Then
        Range("B5").Value = "Problem with running LP: Open Solver error value is " & ResultFlag
    End If
    
End Sub

'Public Sub SetTolerance(Tolerance As Double, Optional sheet As Worksheet)
'    Set Tolerance = 0
'End Sub
