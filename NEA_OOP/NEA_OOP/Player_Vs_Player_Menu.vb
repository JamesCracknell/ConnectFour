Public Class Player_Vs_Player_Menu

    ' Control Declarations

    Private Title As New PictureBox
    Private CreditLabel As New Label
    Private Shadows BackgroundImage As New PictureBox
    Private PlayerChoiceLabel As New Label
    Private InputPlayerNameLabel As New Label
    Private PlayerNameTextBox As New TextBox
    Private SubmitPlayerOneNameButton As New Button
    Private SubmitPlayerTwoNameButton As New Button
    Private TimerOptionButton As New Button 'MAKE THIS A VISUAL TOGGLE
    Private TimerOptionLabel As New Label
    Private TimerExplanationLabel As New Label
    Private PlayerOneCurrentNameLabel As New Label
    Private PlayerTwoCurrentNameLabel As New Label
    Private IncorrectInputLabel As New Label
    Private StartingPlayerLabel As New Label
    Private StartingPlayerButton As New Button
    Private ColourToggleDescriptionLabel As New Label
    Private ColourToggleButton As New Button
    Private CanContinueIndicatorLabel As New Label
    Private ContinueButton As New Button
    Private TimeCustomiserTextbox As New TextBox
    Private TimeCustomiserLabel As New Label

    ' Variable Declarations
    Private FirstTimeRunning As Boolean = True ' some code only needs to run when it is not the first time 

    ' Main Code
    Public Sub This_Form_Setup()
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
            FirstTimeRunning = False
            Menu_Setup()
        Else
            PlayerOneCurrentNameLabel.Text = ""
            PlayerTwoCurrentNameLabel.Text = ""
            TimeCustomiserTextbox.Text = 120
            SubmitPlayerOneNameButton.BackColor = Color.Transparent
            SubmitPlayerTwoNameButton.BackColor = Color.Transparent
        End If
    End Sub
    Public Sub Menu_Setup()
        ' PlayerChoiceLabel
        Processes.RunTimeContructor(PlayerChoiceLabel, 0, 50, 500, 20, "Customise settings for the game:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        Controls.Add(PlayerChoiceLabel)
        PlayerChoiceLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        PlayerChoiceLabel.Refresh() 'refreshes the label, to take on its new properties

        ' InputPlayerNameLabel
        Processes.RunTimeContructor(InputPlayerNameLabel, 0, 70, 500, 20, "Enter player names below:", ContentAlignment.TopCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(InputPlayerNameLabel)
        InputPlayerNameLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        InputPlayerNameLabel.Refresh()

        ' PlayerNameTextBox
        Processes.RunTimeContructor(PlayerNameTextBox, 342, 190, 300, 50, "", ContentAlignment.TopCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Empty) 'color.empty is a null colour, allowing for controls that do not allow customised back colours
        Controls.Add(PlayerNameTextBox)

        ' SubmitPlayerOneNameButton
        Processes.RunTimeContructor(SubmitPlayerOneNameButton, 362, 220, 120, 50, "Submit Player One Name", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler SubmitPlayerOneNameButton.Click, AddressOf SubmitPlayerNameButton_Click
        SubmitPlayerOneNameButton.Name = "SubmitPlayerOneNameButton"
        Controls.Add(SubmitPlayerOneNameButton)

        ' PlayerOneCurrentNameLabel
        Processes.RunTimeContructor(PlayerOneCurrentNameLabel, 113, 165, 132, 20, "", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.Transparent)
        Controls.Add(PlayerOneCurrentNameLabel)
        PlayerOneCurrentNameLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        PlayerOneCurrentNameLabel.Refresh()

        ' SubmitPlayerTwoNameButton
        Processes.RunTimeContructor(SubmitPlayerTwoNameButton, 502, 220, 120, 50, "Submit Player Two Name", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler SubmitPlayerTwoNameButton.Click, AddressOf SubmitPlayerNameButton_Click
        SubmitPlayerTwoNameButton.Name = "SubmitPlayerTwoNameButton"
        Controls.Add(SubmitPlayerTwoNameButton)

        ' PlayerTwoCurrentNameLabel
        Processes.RunTimeContructor(PlayerTwoCurrentNameLabel, 253, 165, 132, 20, "", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.Transparent)
        Controls.Add(PlayerTwoCurrentNameLabel)
        PlayerTwoCurrentNameLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        PlayerTwoCurrentNameLabel.Refresh()

        ' IncorrectInputLabel
        Processes.RunTimeContructor(IncorrectInputLabel, 342, 285, 300, 22, "", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.IndianRed)
        Controls.Add(IncorrectInputLabel)
        IncorrectInputLabel.Visible = False

        ' TimerOptionLabel
        Processes.RunTimeContructor(TimerOptionLabel, 100, 205, 300, 20, "Would you like to add a timer to the game?", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(TimerOptionLabel)
        TimerOptionLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        TimerOptionLabel.Refresh()

        ' TimerOptionButton
        Processes.RunTimeContructor(TimerOptionButton, 392, 327, 200, 30, "Timed Game", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler TimerOptionButton.Click, AddressOf TimerOptionButton_Click
        Controls.Add(TimerOptionButton)

        ' TimerExplanationLabel
        Processes.RunTimeContructor(TimerExplanationLabel, 25, 255, 450, 40, "Timed Game: The game is timed." & vbNewLine & "In the event of a draw, the player with the lowest time wins.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.Transparent)
        Controls.Add(TimerExplanationLabel)
        TimerExplanationLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        TimerExplanationLabel.Refresh()

        ' TimeCustomiserLabel
        Processes.RunTimeContructor(TimeCustomiserLabel, 745, 320, 225, 40, "Enter the time to count down from:" & vbNewLine & "(in seconds)", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        TimeCustomiserLabel.Visible = False
        Controls.Add(TimeCustomiserLabel)

        ' TimeCustomiserTextbox
        Processes.RunTimeContructor(TimeCustomiserTextbox, 755, 360, 205, 30, "120", ContentAlignment.TopCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Empty)
        AddHandler TimeCustomiserTextbox.Click, AddressOf TimerOptionButton_Click
        TimeCustomiserTextbox.Visible = False
        Controls.Add(TimeCustomiserTextbox)

        ' StartingPlayerLabel
        Processes.RunTimeContructor(StartingPlayerLabel, 25, 280, 450, 40, "Choose the starting player:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(StartingPlayerLabel)
        StartingPlayerLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        StartingPlayerLabel.Refresh()

        ' StartingPlayerButton
        Processes.RunTimeContructor(StartingPlayerButton, 392, 415, 200, 30, "Player One", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        AddHandler StartingPlayerButton.Click, AddressOf StartingPlayerButton_Click
        Controls.Add(StartingPlayerButton)

        ' ColourToggleDescriptionLabel
        Processes.RunTimeContructor(ColourToggleDescriptionLabel, 25, 335, 450, 40, "Choose player one's colour:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(ColourToggleDescriptionLabel)
        ColourToggleDescriptionLabel.Parent = BackgroundImage 'sets the label's parent to the background image
        ColourToggleDescriptionLabel.Refresh()

        ' ColourToggleButton
        Processes.RunTimeContructor(ColourToggleButton, 392, 470, 200, 30, "Red", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        ColourToggleButton.ForeColor = Color.Red
        AddHandler ColourToggleButton.Click, AddressOf ColourToggleButton_Click
        Controls.Add(ColourToggleButton)

        ' ContinueButton
        Processes.RunTimeContructor(ContinueButton, 392, 505, 200, 30, "Start Game" & vbNewLine & "In the event of a draw, the player with the lowest time wins.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        ContinueButton.Font = New Font(ContinueButton.Font, FontStyle.Bold)
        AddHandler ContinueButton.Click, AddressOf ContinueButton_Click
        Controls.Add(ContinueButton)

        ' CanContinueIndicatorLabel
        Processes.RunTimeContructor(CanContinueIndicatorLabel, 392, 527, 200, 30, "Please enter both usernames.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.IndianRed)
        Controls.Add(CanContinueIndicatorLabel)
        CanContinueIndicatorLabel.Visible = False
        BackgroundImage.SendToBack() 'sends the image to the back so that all text and buttons appear in front of it
    End Sub
    Public Sub SubmitPlayerNameButton_Click(sender As Object, e As EventArgs)
        CanContinueIndicatorLabel.Visible = False
        CanContinueIndicatorLabel.Refresh()
        If SubmitPlayerOneNameButton.BackColor = Color.IndianRed Then SubmitPlayerOneNameButton.BackColor = Color.Transparent 'if the colour is red, resets it
        If SubmitPlayerTwoNameButton.BackColor = Color.IndianRed Then SubmitPlayerTwoNameButton.BackColor = Color.Transparent 'if the colour is red, resets it
        IncorrectInputLabel.Visible = False
        NameVerification()
        If NameVerification() Then 'if the input was valid
            sender.backcolor = Color.LightGreen
            If sender.name = "SubmitPlayerOneNameButton" Then
                PlayerOneCurrentNameLabel.Text = PlayerNameTextBox.Text
            Else
                PlayerTwoCurrentNameLabel.Text = PlayerNameTextBox.Text
            End If
            PlayerNameTextBox.Text = ""
        Else 'invalid input
            IncorrectInputLabel.Visible = True
            sender.backcolor = Color.IndianRed
            PlayerNameTextBox.Text = ""
        End If
    End Sub
    Public Function NameVerification()
        IncorrectInputLabel.Text = "INVALID INPUT: Name must be 1-16 characters" 'reset default invalid input message
        Dim valid As Boolean = True
        If PlayerNameTextBox.Text = "" Then valid = False 'input can't be blank
        If PlayerNameTextBox.Text.Length > 16 Then valid = False 'input can't be longer than 16 characters
        If (PlayerNameTextBox.Text = PlayerTwoCurrentNameLabel.Text Or PlayerNameTextBox.Text = PlayerOneCurrentNameLabel.Text) And PlayerNameTextBox.Text <> "" Then
            valid = False
            IncorrectInputLabel.Text = "INVALID INPUT: Names must not be the same" 'set special invalid input message
        End If
        Return valid
    End Function
    Public Sub TimerOptionButton_Click(sender As Object, e As EventArgs)
        If sender.text = "Timed Game" Then
            TimeCustomiserLabel.Visible = True
            TimeCustomiserTextbox.Visible = True
            TimerOptionButton.Text = "Countdown"
            TimerExplanationLabel.Text = "Countdown: Set a timer for each player that counts down during their turn." & vbNewLine & "If the time runs out on your turn then you lose!"
        ElseIf sender.text = "Countdown" Then
            TimeCustomiserLabel.Visible = False
            TimeCustomiserTextbox.Visible = False
            TimerOptionButton.Text = "No Timer"
            TimerExplanationLabel.Text = "No Timer:  There is no timer." & vbNewLine & "A draw is a draw!"
        ElseIf sender.text = "No Timer" Then
            TimerOptionButton.Text = "Timed Game"
            TimerExplanationLabel.Text = "Timed Game: The game is timed." & vbNewLine & "In the event of a draw, the player with the lowest time wins."
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
        If sender.Text = "Player One" Then
            StartingPlayerButton.Text = "Player Two"
        ElseIf sender.Text = "Player Two" Then
            StartingPlayerButton.Text = "Random"
        Else
            StartingPlayerButton.Text = "Player One"
        End If
    End Sub
    Public Function getOppositeColour(ColourChoice As String)
        If ColourChoice = "Red" Then
            Return "Yellow"
        Else
            Return "Red"
        End If
    End Function
    Public Sub ContinueButton_Click(sender As Object, e As EventArgs)
        If CanContinue() Then
            CanContinueIndicatorLabel.Visible = False
            Me.Hide()
            Dim SubType As String = TimerOptionButton.Text
            Dim ColourChoice As String = ColourToggleButton.Text
            Dim OppositeColour As String = getOppositeColour(ColourChoice)
            Dim StartingPlayer As String
            If StartingPlayerButton.Text = "Player One" Then
                StartingPlayer = PlayerOneCurrentNameLabel.Text
            ElseIf StartingPlayerButton.Text = "Player Two" Then
                StartingPlayer = PlayerTwoCurrentNameLabel.Text
            Else
                If (CInt(Math.Floor(Rnd() * 2))) = 0 Then 'random number, either 1 or 0 resulting in 50% chance to be either player
                    StartingPlayer = PlayerOneCurrentNameLabel.Text
                Else
                    StartingPlayer = PlayerTwoCurrentNameLabel.Text
                End If
            End If
            Game_Form.Game_Setup("Player", SubType, PlayerOneCurrentNameLabel.Text, ColourChoice, PlayerTwoCurrentNameLabel.Text, OppositeColour, StartingPlayer, TimeCustomiserTextbox.Text)
        Else
            CanContinueIndicatorLabel.Visible = True
        End If
    End Sub
    Private Function CanContinue() 'determines if the user has performed the necessary steps to continue
        If PlayerOneCurrentNameLabel.Text <> "" And PlayerTwoCurrentNameLabel.Text <> "" Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub Player_Vs_Player_Menu_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing 'when x is pressed
        Main_Menu.Close() 'closing first form closes whole program
    End Sub
End Class