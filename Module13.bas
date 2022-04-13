Attribute VB_Name = "Module13"
Sub Send_Email_With_Snapshot()

    'Method in charge of sending the snapshot of the schedule to all volunteers

    '''
    Dim OutApp As Object, OutMail As Object
    Dim sh As Worksheet
    Dim schedule As Range
    '''
    
    Set sh = ThisWorkbook.Sheets("Final_Schedule")
    
    'Dim temp As Integer
    'temp = 4
    
    'Dim last_row As Integer
    'last_row = Application.CountA(sh.Range("D:D"))
    
    
    Set schedule = Nothing
    On Error Resume Next
        Set schedule = Range("B2:D48").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'Uncomment to send emails to all 30 volunteers
    'For i = 4 To last_row
    
        On Error Resume Next
        With OutMail
            'Send to Aviel
            .To = sh.Range("H6").Value
            
            'Uncomment to allow program to send to all emails in database
            '.To = sh.Range("H4:H48").Find(Range("H" & i)).Value
            
            'CC to Cassie and Becca
            .CC = sh.Range("H4").Value
            '.BCC = ""
            .Subject = "VMIS Scheduling"
            
            'Uncomment to allow program to go through every name in the database
            '.Subject = "VMIS Scheduling for " & sh.Range("D4:D48").Find(Range("D" & i)).Value
            
            .HTMLBody = "<BODY style = font-size:12pt; font-family: Calibri>" & _
            "Hello! <p> Here is the new VMIS schedule for this semester: <p>" & RangetoHTML(schedule)
            .Send
        End With
        On Error GoTo 0
        
    'Next i
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    MsgBox "Thanks for using our Scheduling tool!" & vbNewLine & vbNewLine & _
    "Schedules were successfully sent to all VMIS Volunteers."



End Sub

Function RangetoHTML(schedule As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in[/COLOR]
    schedule.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
            .DrawingObjects.Visible = True
            .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file[/COLOR]
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML[/COLOR]
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB[/COLOR]
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function[/COLOR]
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
End Function
