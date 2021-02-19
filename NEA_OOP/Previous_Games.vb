Public Class Previous_Games
    ' Control Declarations

    Private Title As New PictureBox
    Private CreditLabel As New Label

    ' Variable Declarations

    ' Main Code
    Public Sub Form_Setup()
        Me.Show()
        MaximizeBox = False 'disable user making the window a full screen
        MinimizeBox = False 'disable user minimizing 
        FormBorderStyle = FormBorderStyle.FixedSingle 'disable user changing window size
        Width = 1000 'sets width of window
        Height = 800 'sets height of window 
        StartPosition = FormStartPosition.Manual 'allows me to change the location of the window
        Location = New Point(0, 0) 'moves window to top left
        BackColor = Color.LightBlue 'sets background colour to blue
        Icon = New Icon("icon.ico") 'sets icon to custom icon
        ' Form Title
        Title.Location = New Point(242, 5)
        Title.Size = New Size(500, 50)
        Title.SizeMode = PictureBoxSizeMode.Zoom
        Title.Image = Image.FromFile("Title.png")
        Controls.Add(Title)
        Main_Menu.RunTimeContructor(CreditLabel, 0, 748, 984, 13, "James Cracknell - 191673", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 8.25, Cursors.Default, Color.Transparent)
        Controls.Add(CreditLabel)
    End Sub
End Class
'https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.datagridview?view=net-5.0