Public Class Player_Vs_Computer_Menu 'Menu for configuring the player vs computer mode
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
    Private Username As String = ""
    Private FirstTimeRunning As Boolean = True ' some code only needs to run when it is not the first time 

    ' Code
    Public Sub This_Form_Setup() 'sets up form properties
        Me.Show()
        If FirstTimeRunning Then
            Processes.Form_Setup(Me)
            ' Form Title
            Title.Location = New Point(242, 5)
            Title.Size = New Size(500, 50)
            Title.SizeMode = PictureBoxSizeMode.Zoom
            Title.Image = Image.FromFile("Title.png")
            Controls.Add(Title)
            Processes.RunTimeContructor(CreditLabel, 0, 748, 984, 13, "James Cracknell - 191673", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 8.25, Cursors.Default, Color.Transparent)
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
            IncorrectInputLabel.Visible = False
            UsernameEnterButton.BackColor = Color.Transparent
        End If
    End Sub
    Private Sub Menu_Setup()  'adds form controls
        ' ChoiceSubtitleLabel
        Processes.RunTimeContructor(ChoiceSubtitleLabel, 0, 40, 500, 20, "Customise settings for the game:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        Controls.Add(ChoiceSubtitleLabel)
        ChoiceSubtitleLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        ChoiceSubtitleLabel.Refresh() 'refreshes the label, to take on its new properties

        ' UsernameEnterLabel
        Processes.RunTimeContructor(UsernameEnterLabel, 0, 65, 500, 20, "Enter username below:", ContentAlignment.TopCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(UsernameEnterLabel)
        UsernameEnterLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        UsernameEnterLabel.Refresh()

        ' UsernameEnterTextbox
        Processes.RunTimeContructor(UsernameEnterTextbox, 342, 190, 300, 50, "", ContentAlignment.TopCenter, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Empty) 'color.empty is a null colour, allowing for controls that do not allow customised back colours
        Controls.Add(UsernameEnterTextbox)

        ' UsernameEnterButton
        Processes.RunTimeContructor(UsernameEnterButton, 435, 220, 120, 50, "Submit Username", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler UsernameEnterButton.Click, AddressOf UsernameEnterButton_Click
        Controls.Add(UsernameEnterButton)

        ' IncorrectInputLabel
        Processes.RunTimeContructor(IncorrectInputLabel, 342, 280, 300, 22, "INVALID INPUT: Name must be 1-16 characters", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.IndianRed)
        Controls.Add(IncorrectInputLabel)
        IncorrectInputLabel.Visible = False

        ' DifficultySubtitleLabel
        Processes.RunTimeContructor(DifficultySubtitleLabel, 100, 205, 300, 20, "Customise the difficulty of the computer:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(DifficultySubtitleLabel)
        DifficultySubtitleLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        DifficultySubtitleLabel.Refresh()

        ' DifficultyToggleButton
        Processes.RunTimeContructor(DifficultyToggleButton, 392, 330, 200, 30, "Easy Difficulty", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler DifficultyToggleButton.Click, AddressOf DifficultyToggleButton_Click
        Controls.Add(DifficultyToggleButton)

        ' DifficultyDescriptionLabel
        Processes.RunTimeContructor(DifficultyDescriptionLabel, 25, 255, 450, 40, "Easy Difficulty: The computer controlled player is easy to beat." & vbNewLine & "Recommended for new players.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.Transparent)
        Controls.Add(DifficultyDescriptionLabel)
        DifficultyDescriptionLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        DifficultyDescriptionLabel.Refresh()

        ' ColourToggleDescriptionLabel
        Processes.RunTimeContructor(ColourToggleDescriptionLabel, 25, 285, 450, 40, "Choose your colour:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(ColourToggleDescriptionLabel)
        ColourToggleDescriptionLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        ColourToggleDescriptionLabel.Refresh()

        ' ColourToggleButton
        Processes.RunTimeContructor(ColourToggleButton, 392, 415, 200, 30, "Red", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        ColourToggleButton.ForeColor = Color.Red
        AddHandler ColourToggleButton.Click, AddressOf ColourToggleButton_Click
        Controls.Add(ColourToggleButton)

        ' StartingPlayerLabel
        Processes.RunTimeContructor(StartingPlayerLabel, 25, 335, 450, 40, "Choose the starting player:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(StartingPlayerLabel)
        StartingPlayerLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        StartingPlayerLabel.Refresh()

        ' StartingPlayerButton
        Processes.RunTimeContructor(StartingPlayerButton, 392, 465, 200, 30, "Computer Player", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        AddHandler StartingPlayerButton.Click, AddressOf StartingPlayerButton_Click
        Controls.Add(StartingPlayerButton)

        ' ContinueButton
        Processes.RunTimeContructor(ContinueButton, 392, 505, 200, 30, "Start Game", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        ContinueButton.Font = New Font(ContinueButton.Font, FontStyle.Bold)
        AddHandler ContinueButton.Click, AddressOf ContinueButton_Click
        Controls.Add(ContinueButton)

        ' CanContinueIndicatorLabel
        Processes.RunTimeContructor(CanContinueIndicatorLabel, 342, 535, 300, 22, "Please enter a valid username to continue", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.IndianRed)
        Controls.Add(CanContinueIndicatorLabel)
        CanContinueIndicatorLabel.Visible = False

        BackgroundImage.SendToBack()
    End Sub
    Private Sub UsernameEnterButton_Click(sender As Object, e As EventArgs) 'When the username enter button is clicked
        If UsernameEnterButton.BackColor = Color.IndianRed Then UsernameEnterButton.BackColor = Color.Transparent 'if the colour is red, resets it
        IncorrectInputLabel.Visible = False
        If NameVerification() Then 'if the input was valid
            sender.backcolor = Color.LightGreen
            Username = UsernameEnterTextbox.Text
            UsernameEnterTextbox.Text = "" 'clears inputted text
        Else 'invalid input
            IncorrectInputLabel.Visible = True
            sender.backcolor = Color.IndianRed 'sets background to re
            UsernameEnterTextbox.Text = "" 'clears inputted text
        End If
    End Sub
    Private Function NameVerification() As Boolean ' checks if the inputted username is valid
        Dim Valid As Boolean = True
        If UsernameEnterTextbox.Text = "" Then Valid = False 'input can't be blank
        If UsernameEnterTextbox.Text.Length > 16 Then Valid = False 'input can't be longer than 16 characters
        Return Valid
    End Function
    Private Sub DifficultyToggleButton_Click(sender As Object, e As EventArgs) 'toggles between different difficulties upon click
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
    Private Sub ColourToggleButton_Click(sender As Object, e As EventArgs) 'toggles between colours for player upon click
        If sender.Text = "Red" Then
            ColourToggleButton.Text = "Yellow"
            ColourToggleButton.ForeColor = Color.Yellow
        Else
            ColourToggleButton.Text = "Red"
            ColourToggleButton.ForeColor = Color.Red
        End If
    End Sub
    Private Sub StartingPlayerButton_Click(sender As Object, e As EventArgs) 'toggles between the player who starts upon click
        If sender.Text = Username Then
            StartingPlayerButton.Text = "Computer Player"
        ElseIf sender.Text = "Computer Player" Then
            StartingPlayerButton.Text = "Random"
        Else
            StartingPlayerButton.Text = Username
        End If
    End Sub
    Private Sub ContinueButton_Click(sender As Object, e As EventArgs) 'Starts the game based on inputted parameters
        If Username <> "" Then 'requirement for game to continue
            CanContinueIndicatorLabel.Visible = False
            Me.Hide()
            Dim Difficulty As String = DifficultyToggleButton.Text
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
    Private Sub Player_Vs_Computer_Menu_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing 'when x is pressed
        Main_Menu.Close() 'closing first form closes whole program
    End Sub
End Class
