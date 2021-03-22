Module Processes
    ' The code in here is run across all forms, it has been moved here to avoid repetition. 
    Public Sub Form_Setup(FormName) 'This sets the form style
        With FormName
            FormName.MaximizeBox = False 'disable user making the window a full screen
            FormName.MinimizeBox = False 'disable user minimizing 
            FormName.FormBorderStyle = FormBorderStyle.FixedSingle 'disable user changing window size
            FormName.Width = 1000 'sets width of window
            FormName.Height = 800 'sets height of window 
            FormName.StartPosition = FormStartPosition.Manual 'allows me to change the location of the window
            FormName.Location = New Point(0, 0) 'moves window to top left
            FormName.BackColor = Color.LightBlue 'sets background colour to blue
            FormName.Icon = New Icon("icon.ico") 'sets icon to custom icon
        End With
    End Sub
    Public Sub RunTimeContructor(Name, locationX, locationy, sizeX, sizeY, text, textalignment, fontname, fontsize, cursortype, BackColor) 'this creates a control on the form.
        With Name
            .Location = New Point(locationX, locationy)
            .Size = New Size(sizeX, sizeY)
            .Text = text
            .TextAlign = textalignment
            .font = New Font(CStr(fontname), fontsize)
            .Cursor = cursortype
            .BackColor = BackColor
        End With
    End Sub

End Module
