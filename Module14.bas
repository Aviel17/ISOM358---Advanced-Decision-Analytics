Attribute VB_Name = "Module14"
Sub CreateEmailButton()

    Application.ScreenUpdating = False
    ActiveSheet.Buttons.Delete
    
    Dim btn As Button
    Dim t As Range
    Set t = ActiveSheet.Range(Cells(4, 11), Cells(7, 14))
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)

    'Customization of button
    btn.Font.ColorIndex = 3
    btn.Font.Bold = True
    btn.Font.Size = 16
    btn.Font.Name = "Times New Roman"

    'What the button will do once you press it
    btn.OnAction = "Send_Email_With_Snapshot"
    btn.Caption = "Send Out Emails"
    btn.Name = "Btn"

    Application.ScreenUpdating = True

End Sub
