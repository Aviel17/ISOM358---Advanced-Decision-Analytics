Attribute VB_Name = "Module5"
Sub DeleteAllWorksheetsExceptActive()
'Step 1:  Declare your variables
    Dim ws As Worksheet
'Step 2: Start looping through all worksheets
    For Each ws In ThisWorkbook.Worksheets
 
'Step 3: Check each worksheet name
    If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
 
'Step 4: Turn off warnings and delete
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    End If
   
'Step 5:  Loop to next worksheet
    Next ws
End Sub
