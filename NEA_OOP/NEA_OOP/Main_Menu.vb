
Public Class Main_Menu

    ' Control Declarations

    Private Title As New PictureBox
    Private CreditLabel As New Label
    Private PlayerVsPlayerButton As New Button
    Private PlayerVsComputerButton As New Button
    Private PreviousGamesButton As New Button
    Private GameTypeChoiceLabel As New Label
    Private Shadows BackgroundImage As New PictureBox
    Private QuitButton As New Button

    ' Main Code 
    Private Sub Main_Menu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        This_Form_Setup()
        Main_Menu_Setup()
    End Sub
    Public Sub This_Form_Setup()
        Processes.Form_Setup(Me) 'Calls the processes module to run the form formatting code.
        ' Form Title
        Title.Location = New Point(242, 5)
        Title.Size = New Size(500, 50)
        Title.SizeMode = PictureBoxSizeMode.Zoom
        Title.Image = Image.FromFile("Title.png")
        Controls.Add(Title)
        Processes.RunTimeContructor(CreditLabel, 0, 748, 984, 13, "James Cracknell - 191673", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 8.25, Cursors.Default, Color.Transparent)
        Controls.Add(CreditLabel)
    End Sub
    Public Sub Main_Menu_Setup()
        ' Background Image
        BackgroundImage = New PictureBox
        BackgroundImage.Location = New Point(242, 100)
        BackgroundImage.Size = New Size(500, 500)
        BackgroundImage.SizeMode = PictureBoxSizeMode.Zoom 'sets size mode to allow the whole image to be shown
        BackgroundImage.Image = Image.FromFile("BackgroundImage.png")
        Controls.Add(BackgroundImage)
        ' GameTypeChoiceLabel
        Processes.RunTimeContructor(GameTypeChoiceLabel, 0, 70, 500, 20, "Choose an option to continue:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        Controls.Add(GameTypeChoiceLabel)
        GameTypeChoiceLabel.Parent = BackgroundImage 'sets the labels parent to the background image so that when it changes the background to "transparent" which copies the background of its parent (initially the form)
        GameTypeChoiceLabel.Refresh() 'refreshes the label, to take on its new properties
        ' PlayerVsPlayerButton
        Processes.RunTimeContructor(PlayerVsPlayerButton, 342, 200, 300, 50, "Player vs Player", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler PlayerVsPlayerButton.Click, AddressOf PlayerVsPlayerButton_Click
        Controls.Add(PlayerVsPlayerButton)
        ' PlayerVsComputerButton
        Processes.RunTimeContructor(PlayerVsComputerButton, 342, 260, 300, 50, "Player vs Computer", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler PlayerVsComputerButton.Click, AddressOf PlayerVsComputerButton_Click
        Controls.Add(PlayerVsComputerButton)
        ' PreviousGamesButton
        Processes.RunTimeContructor(PreviousGamesButton, 342, 320, 300, 50, "View Previous Games", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler PreviousGamesButton.Click, AddressOf PreviousGamesButton_Click
        Controls.Add(PreviousGamesButton)
        ' QuitButton
        Processes.RunTimeContructor(QuitButton, 342, 380, 300, 50, "Quit Game", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler QuitButton.Click, AddressOf QuitButton_Click
        Controls.Add(QuitButton)
        BackgroundImage.SendToBack() 'moves image to back
    End Sub
    Public Sub PlayerVsPlayerButton_Click(sender As Object, e As EventArgs)
        Me.Hide()
        Player_Vs_Player_Menu.This_Form_Setup()
    End Sub
    Public Sub PlayerVsComputerButton_Click(sender As Object, e As EventArgs)
        Me.Hide()
        Player_Vs_Computer_Menu.This_Form_Setup()
    End Sub
    Public Sub PreviousGamesButton_Click(sender As Object, e As EventArgs)
        Me.Hide()
        Previous_Games.This_Form_Setup()
    End Sub
    Sub QuitButton_Click(sender As Object, e As EventArgs) 'if quitbutton is clicked
        Me.Close()
    End Sub

End Class
