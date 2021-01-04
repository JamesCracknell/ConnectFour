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
    Private ContinueButton As New Button
    Private ContinueMessageLabel As New Label
    Private PlayerOneCurrentNameLabel As New Label
    Private PlayerTwoCurrentNameLabel As New Label
    Private IncorrectInputLabel As New Label

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
        ' Background Image
        BackgroundImage = New PictureBox
        BackgroundImage.Location = New Point(242, 100)
        BackgroundImage.Size = New Size(500, 500)
        BackgroundImage.SizeMode = PictureBoxSizeMode.Zoom 'sets size mode to allow the whole image to be shown
        BackgroundImage.Image = Image.FromFile("BackgroundImage.png")
        Controls.Add(BackgroundImage)
        Menu_Setup()
    End Sub

    Public Sub Menu_Setup()
        ' PlayerChoiceLabel
        Main_Menu.RunTimeContructor(PlayerChoiceLabel, 0, 70, 500, 20, "Customise settings for the game:", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        Controls.Add(PlayerChoiceLabel)
        PlayerChoiceLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        PlayerChoiceLabel.Refresh() 'refreshes the label, to take on its new properties
        ' InputPlayerNameLabel
        Main_Menu.RunTimeContructor(InputPlayerNameLabel, 0, 105, 500, 20, "Enter player names below:", ContentAlignment.TopCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(InputPlayerNameLabel)
        InputPlayerNameLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        InputPlayerNameLabel.Refresh()
        ' PlayerNameTextBox
        Main_Menu.RunTimeContructor(PlayerNameTextBox, 342, 230, 300, 50, "", ContentAlignment.TopCenter, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Empty) 'color.empty is a null colour, allowing for controls that do not allow customised back colours
        Controls.Add(PlayerNameTextBox)
        ' SubmitPlayerOneNameButton
        Main_Menu.RunTimeContructor(SubmitPlayerOneNameButton, 362, 260, 120, 50, "Submit Red Player Name", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler SubmitPlayerOneNameButton.Click, AddressOf SubmitPlayerNameButton_Click
        SubmitPlayerOneNameButton.Name = "SubmitPlayerOneNameButton"
        Controls.Add(SubmitPlayerOneNameButton)
        ' PlayerOneCurrentNameLabel
        Main_Menu.RunTimeContructor(PlayerOneCurrentNameLabel, 113, 210, 132, 20, "", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.Transparent)
        Controls.Add(PlayerOneCurrentNameLabel)
        PlayerOneCurrentNameLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        PlayerOneCurrentNameLabel.Refresh()
        ' SubmitPlayerTwoNameButton
        Main_Menu.RunTimeContructor(SubmitPlayerTwoNameButton, 502, 260, 120, 50, "Submit Yellow Player Name", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler SubmitPlayerTwoNameButton.Click, AddressOf SubmitPlayerNameButton_Click
        SubmitPlayerTwoNameButton.Name = "SubmitPlayerTwoNameButton"
        Controls.Add(SubmitPlayerTwoNameButton)
        ' PlayerTwoCurrentNameLabel
        Main_Menu.RunTimeContructor(PlayerTwoCurrentNameLabel, 253, 210, 132, 20, "", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.Transparent)
        Controls.Add(PlayerTwoCurrentNameLabel)
        PlayerTwoCurrentNameLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        PlayerTwoCurrentNameLabel.Refresh()
        ' IncorrectInputLabel
        Main_Menu.RunTimeContructor(IncorrectInputLabel, 342, 320, 300, 22, "", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.IndianRed)
        Controls.Add(IncorrectInputLabel)
        IncorrectInputLabel.Visible = False
        ' TimerOptionButton
        Main_Menu.RunTimeContructor(TimerOptionButton, 392, 375, 200, 30, "Timed Game", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler TimerOptionButton.Click, AddressOf TimerOptionButton_Click
        Controls.Add(TimerOptionButton)
        ' TimerOptionLabel
        Main_Menu.RunTimeContructor(TimerOptionLabel, 100, 250, 300, 20, "Would you like to add a timer to the game?", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.Transparent)
        Controls.Add(TimerOptionLabel)
        TimerOptionLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        TimerOptionLabel.Refresh()
        ' TimerExplanationLabel
        Main_Menu.RunTimeContructor(TimerExplanationLabel, 25, 300, 450, 40, "Timed Game: The game is timed." & vbNewLine & "In the event of a draw, the player with the lowest time wins.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.Transparent)
        Controls.Add(TimerExplanationLabel)
        TimerExplanationLabel.Parent = BackgroundImage 'sets the label's parent to the background image 
        TimerExplanationLabel.Refresh()
        ' ContinueButton
        Main_Menu.RunTimeContructor(ContinueButton, 392, 450, 200, 30, "Start Game" & vbNewLine & "In the event of a draw, the player with the lowest time wins.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Hand, Color.Transparent)
        ContinueButton.Font = New Font(ContinueButton.Font, FontStyle.Bold)
        AddHandler ContinueButton.Click, AddressOf ContinueButton_Click
        Controls.Add(ContinueButton)
        ' ContinueMessageLabel
        Main_Menu.RunTimeContructor(ContinueMessageLabel, 392, 475, 200, 30, "Please enter both usernames.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9, Cursors.Default, Color.IndianRed)
        Controls.Add(ContinueMessageLabel)
        ContinueMessageLabel.Visible = False
        BackgroundImage.SendToBack() 'sends the image to the back so that all text and buttons appear in front of it
    End Sub
    Public Sub SubmitPlayerNameButton_Click(sender As Object, e As EventArgs)
        ContinueMessageLabel.Visible = False
        ContinueMessageLabel.Refresh()
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
            TimerOptionButton.Text = "Countdown"
            TimerExplanationLabel.Text = "Countdown: Set a timer for each player that counts down during their turn." & vbNewLine & "If the time runs out on your turn then you lose!"
        ElseIf sender.text = "Countdown" Then
            TimerOptionButton.Text = "No Timer"
            TimerExplanationLabel.Text = "No Timer:  There is no timer." & vbNewLine & "A draw is a draw!"
        ElseIf sender.text = "No Timer" Then
            TimerOptionButton.Text = "Timed Game"
            TimerExplanationLabel.Text = "Timed Game: The game is timed." & vbNewLine & "In the event of a draw, the player with the lowest time wins."
        End If
    End Sub
    Public Sub ContinueButton_Click() ' this code is very wrong
        If CanContinue() = True Then

            Height = 900
            CreditLabel.Location = New Point(0, 848)
            'SetUpPlayerVsPlayerGame()
            ' Dim GameSelected As String = TimerOptionButton.Text
            ' Dim CurrentGame As New PlayerVsPlayerGame(GameSelected)
        Else
            ContinueMessageLabel.Visible = True
            ContinueMessageLabel.Refresh()
        End If
    End Sub
    Private Function CanContinue() 'determines if the user has performed the necessary steps to continue
        If PlayerOneCurrentNameLabel.Text <> "" And PlayerTwoCurrentNameLabel.Text <> "" Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Player_Vs_Player_Menu_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class