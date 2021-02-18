Public Class Player_Vs_Computer_Menu

    ' Control Declarations

    Private Title As New PictureBox
    Private CreditLabel As New Label
    Private Shadows BackgroundImage As New PictureBox
    Private ChoiceSubtitleLabel As New Label
    Private ContinueButton As New Button
    Private UsernameEnterButton As New Button
    Private UsernameEnterTextbox As New TextBox
    Private UsernameEnterLabel As New Label
    Private DifficultySubtitleLabel As New Label
    Private DifficultyToggleButton As New Button
    Private DifficultyDescriptionLabel As New Label
    Private IncorrectInputLabel As New Label
    Private ColourToggleDescriptionLabel As New Label
    Private ColourToggleButton As New Button
    Private CanContinueIndicatorLabel As New Label
    Private StartingPlayerLabel As New Label
    Private StartingPlayerButton As New Button

    ' Variable Declarations
    Private Username As String = "Player"
    Private FirstTimeRunning As Boolean = True ' some code only needs to run when it is not the first time 

    ' Main Code
    Public Sub Form_Setup()
        Me.Show()
        If FirstTimeRunning Then
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
            ' Background Image
            BackgroundImage = New PictureBox
            BackgroundImage.Location = New Point(242, 100)
            BackgroundImage.Size = New Size(500, 500)
            BackgroundImage.SizeMode = PictureBoxSizeMode.Zoom 'sets size mode to allow the whole image to be shown
            BackgroundImage.Image = Image.FromFile("BackgroundImage.png")
            Controls.Add(BackgroundImage)
            BackgroundImage.SendToBack() 'moves image to back
            FirstTimeRunning = False
            Menu_Setup()
        Else 'if code is not running for the first time
            Username = "" 'resets username
            UsernameEnterButton.BackColor = Color.Transparent
        End If
    End Sub

    Public Sub Menu_Setup()
        ' ChoiceSubtitleLabel
        Main_Menu.RunTimeContructor(ChoiceSubtitleLabel, 0, 40, 500, 20, "Customise settings for the game:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        Controls.Add(ChoiceSubtitleLabel)
        ChoiceSubtitleLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        ChoiceSubtitleLabel.Refresh() 'refreshes the label, to take on its new properties

        ' UsernameEnterLabel
        Main_Menu.RunTimeContructor(UsernameEnterLabel, 0, 65, 500, 20, "Enter username below:", ContentAlignment.TopCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(UsernameEnterLabel)
        UsernameEnterLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        UsernameEnterLabel.Refresh()

        ' UsernameEnterTextbox
        Main_Menu.RunTimeContructor(UsernameEnterTextbox, 342, 190, 300, 50, "", ContentAlignment.TopCenter, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Empty) 'color.empty is a null colour, allowing for controls that do not allow customised back colours
        Controls.Add(UsernameEnterTextbox)

        ' UsernameEnterButton
        Main_Menu.RunTimeContructor(UsernameEnterButton, 435, 220, 120, 50, "Submit Username", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler UsernameEnterButton.Click, AddressOf UsernameEnterButton_Click
        Controls.Add(UsernameEnterButton)

        ' IncorrectInputLabel
        Main_Menu.RunTimeContructor(IncorrectInputLabel, 342, 280, 300, 22, "INVALID INPUT: Name must be 1-16 characters", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.IndianRed)
        Controls.Add(IncorrectInputLabel)
        IncorrectInputLabel.Visible = False

        ' DifficultySubtitleLabel
        Main_Menu.RunTimeContructor(DifficultySubtitleLabel, 100, 205, 300, 20, "Customise the difficulty of the computer:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(DifficultySubtitleLabel)
        DifficultySubtitleLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        DifficultySubtitleLabel.Refresh()

        ' DifficultyToggleButton
        Main_Menu.RunTimeContructor(DifficultyToggleButton, 392, 330, 200, 30, "Easy Difficulty", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler DifficultyToggleButton.Click, AddressOf DifficultyToggleButton_Click
        Controls.Add(DifficultyToggleButton)

        ' DifficultyDescriptionLabel
        Main_Menu.RunTimeContructor(DifficultyDescriptionLabel, 25, 255, 450, 40, "Easy Difficulty: The computer controlled player is easy to beat." & vbNewLine & "Recommended for new players.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.Transparent)
        Controls.Add(DifficultyDescriptionLabel)
        DifficultyDescriptionLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        DifficultyDescriptionLabel.Refresh()

        ' ColourToggleDescriptionLabel
        Main_Menu.RunTimeContructor(ColourToggleDescriptionLabel, 25, 285, 450, 40, "Choose your colour:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(ColourToggleDescriptionLabel)
        ColourToggleDescriptionLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        ColourToggleDescriptionLabel.Refresh()

        ' ColourToggleButton
        Main_Menu.RunTimeContructor(ColourToggleButton, 392, 415, 200, 30, "Red", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        ColourToggleButton.ForeColor = Color.Red
        AddHandler ColourToggleButton.Click, AddressOf ColourToggleButton_Click
        Controls.Add(ColourToggleButton)

        ' StartingPlayerLabel
        Main_Menu.RunTimeContructor(StartingPlayerLabel, 25, 335, 450, 40, "Choose the starting player:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(StartingPlayerLabel)
        StartingPlayerLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        StartingPlayerLabel.Refresh()

        ' StartingPlayerButton
        Main_Menu.RunTimeContructor(StartingPlayerButton, 392, 465, 200, 30, "Computer Player", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        AddHandler StartingPlayerButton.Click, AddressOf StartingPlayerButton_Click
        Controls.Add(StartingPlayerButton)

        ' ContinueButton
        Main_Menu.RunTimeContructor(ContinueButton, 392, 505, 200, 30, "Start Game", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        ContinueButton.Font = New Font(ContinueButton.Font, FontStyle.Bold)
        AddHandler ContinueButton.Click, AddressOf ContinueButton_Click
        Controls.Add(ContinueButton)

        ' CanContinueIndicatorLabel
        Main_Menu.RunTimeContructor(CanContinueIndicatorLabel, 342, 535, 300, 22, "Please enter a valid username to continue", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.IndianRed)
        Controls.Add(CanContinueIndicatorLabel)
        CanContinueIndicatorLabel.Visible = False

        BackgroundImage.SendToBack()
    End Sub

    Public Sub UsernameEnterButton_Click(sender As Object, e As EventArgs)
        If UsernameEnterButton.BackColor = Color.IndianRed Then UsernameEnterButton.BackColor = Color.Transparent 'if the colour is red, resets it
        IncorrectInputLabel.Visible = False
        If NameVerification() Then 'if the input was valid
            sender.backcolor = Color.LightGreen
            Username = UsernameEnterTextbox.Text
            UsernameEnterTextbox.Text = ""
        Else 'invalid input
            IncorrectInputLabel.Visible = True
            sender.backcolor = Color.IndianRed
            UsernameEnterTextbox.Text = ""
        End If
    End Sub
    Public Function NameVerification() ' checks if the inputted username is valid
        Dim Valid As Boolean = True
        If UsernameEnterTextbox.Text = "" Then Valid = False 'input can't be blank
        If UsernameEnterTextbox.Text.Length > 16 Then Valid = False 'input can't be longer than 16 characters
        Return Valid
    End Function
    Public Sub DifficultyToggleButton_Click(sender As Object, e As EventArgs)
        If sender.Text = "Easy Difficulty" Then
            DifficultyToggleButton.Text = "Medium Difficulty"
            DifficultyDescriptionLabel.Text = "Medium Difficulty: This difficulty is recommended for most players."
        ElseIf sender.Text = "Medium Difficulty" Then
            DifficultyToggleButton.Text = "Hard Difficulty"
            DifficultyDescriptionLabel.Text = "Hard Difficulty: The computer controlled player is hard to beat." & vbNewLine & "Recommended for experienced players."
        ElseIf sender.Text = "Hard Difficulty" Then
            DifficultyToggleButton.Text = "Impossible Difficulty"
            DifficultyDescriptionLabel.Text = "Impossible Difficulty: The computer controlled player is impossible to beat." & vbNewLine & "You WILL lose or draw."
        ElseIf sender.Text = "Impossible Difficulty" Then
            DifficultyToggleButton.Text = "Easy Difficulty"
            DifficultyDescriptionLabel.Text = "Easy Difficulty: The computer controlled player is easy to beat." & vbNewLine & "Recommended for new players."
        End If
    End Sub

    Public Sub ColourToggleButton_Click(sender As Object, e As EventArgs)
        If sender.Text = "Red" Then
            ColourToggleButton.Text = "Yellow"
            ColourToggleButton.ForeColor = Color.Yellow
        Else
            ColourToggleButton.Text = "Red"
            ColourToggleButton.ForeColor = Color.Red
        End If
    End Sub

    Public Sub StartingPlayerButton_Click(sender As Object, e As EventArgs)
        If sender.Text = Username Then
            StartingPlayerButton.Text = "Computer Player"
        ElseIf sender.Text = "Computer Player" Then
            StartingPlayerButton.Text = "Random"
        Else
            StartingPlayerButton.Text = Username
        End If
    End Sub
    Public Sub ContinueButton_Click(sender As Object, e As EventArgs)
        If Username <> "" Then
            CanContinueIndicatorLabel.Visible = False
            Me.Hide()
            Dim Difficulty As String = DifficultyToggleButton.Text ''JOEQUESTION: Should I get these using a getter
            Dim ColourChoice As String = ColourToggleButton.Text
            Dim StartingPlayer As String = StartingPlayerButton.Text
            If StartingPlayer <> "Computer Player" And StartingPlayer <> "Random" Then 'if the starting player is a username, ensures the username is updated
                StartingPlayer = Username
            End If
            Game_Form.Game_Setup("Computer", Difficulty, Username, ColourChoice, "", "", StartingPlayer, 0)
        Else
            CanContinueIndicatorLabel.Visible = True
        End If
    End Sub
End Class
