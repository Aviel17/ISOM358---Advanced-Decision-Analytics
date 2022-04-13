Attribute VB_Name = "Module1"
Sub Get_Data_From_File()

    'Delete all worksheets to ensure memory efficiency
    DeleteAllWorksheetsExceptActive
    Application.Run "Module12.Menu_Background"
    
    Dim FileToUse As Variant
    Dim OpenBook As Workbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    FileToUse = Application.GetOpenFilename(Title:="Browse for your Microsoft Form Data", FileFilter:="Excel Files (*.xlsx*), *xlsx*")

    If FileToUse <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToUse)
        OpenBook.Sheets(1).Range("F1:BB31").Copy
        ThisWorkbook.Sheets.Add(After:=Sheets(1)).Name = "Raw_Data"
        ThisWorkbook.Worksheets("Raw_Data").Range("B1").PasteSpecial xlPasteValues
        OpenBook.Close False
    
        'Calls the AddDay function in Module 3
        Application.Run "Module3.AddDays"
        'Calls the CleanupData function in Module 2
        Application.Run "Module2.CleanupData"
        'Create the button to move onto the next sheet (found in Module 4)
        Application.Run "Module4.CreateButton"
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    
    End If
    
End Sub


