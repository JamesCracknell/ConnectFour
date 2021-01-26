Public Class Game_Form

    ' Control Declarations

    Private Title As New PictureBox
    Private CreditLabel As New Label
    Private BoardLocations(6, 5) As PictureBox
    Private LocationIndicators(6) As PictureBox
    Private StartGameButton As New Button
    Private CurrentGame
    Private CurrentMoveLabel As New Label
    Private CurrentPlayerLabel As New Label
    Private PlayerOneNameLabel As New Label
    Private PlayerTwoNameLabel As New Label
    Private PlayerOneTimer As New Timer
    Private PlayerTwoTimer As New Timer
    Private PlayerOneTimerLabel As New Label
    Private PlayerTwoTimerLabel As New Label


    ' Variable Declarations

    Private BoardEnabled As Boolean = False
    Private StartingPlayer As String
    Private FirstTimeRunning As Boolean = True ' some code only needs to run when it is not the first time 
    Private PreviousGameType As String
    Private BoardPainted As Boolean = False 'only allows board to be painted once a game
    Private PlayerOneTimerValue As Decimal
    Private PlayerTwoTimerValue As Decimal
    Private CurrentGameType As String

    ' Main Code
    ''JOEQUESTION do these need byval and byref. should these be obtained using a getter instead.
    ''JOEQUESTION should names sent down be different to names saved as shown below
    Public Sub Game_Setup(GameType As String, SubType As String, PlayerOneName As String, PlayerOneColour As String, PlayerTwoName As String, PlayerTwoColour As String, GameStartingPlayer As String, CountdownTime As Integer) 'decide game type here
        Me.Show()
        StartingPlayer = GameStartingPlayer
        CurrentGameType = GameType
        If GameType = "Computer" Then
            CurrentGame = New Player_Vs_Computer_Game(SubType, PlayerOneName, PlayerOneColour, GameStartingPlayer)
            If FirstTimeRunning Then
                Board_Setup()
                Form_Setup()
                Computer_Form_SetUp()
            Else
                Board_Update()
                UpdateCurrentMoveVisual()
                Controls.Add(StartGameButton)
                If PreviousGameType <> "Computer" Then 'previous game was NOT a computer game
                    DisableTimers()
                    'needs to switch menu
                End If
            End If
            'set up computer specific forms
        ElseIf GameType = "Player" Then
            CurrentGame = New Player_Vs_Player_Game(SubType, PlayerOneName, PlayerOneColour, PlayerTwoName, PlayerTwoColour, GameStartingPlayer)
            If FirstTimeRunning Then
                Board_Setup()
                Form_Setup()
                Player_Form_Setup(SubType, CountdownTime)
            Else
                DisableTimers()
                Board_Update()
                UpdateCurrentMoveVisual()
                SetTimerValues(CountdownTime) ''''''''''''''''''''''''''''
                Controls.Add(StartGameButton)
            End If
            If SubType = "Countdown" Then
                Countdown_Game_Setup(SubType, CountdownTime)
            ElseIf SubType = "Timed Game" Then
                Timed_Game_Setup(SubType)
            End If
        End If
        BoardPainted = False
        PreviousGameType = GameType
        FirstTimeRunning = False 'first time running set to false after all start up code has run
    End Sub
    Public Sub Form_Setup()
        MaximizeBox = False 'disable user making the window a full screen
        MinimizeBox = False 'disable user minimizing
        FormBorderStyle = FormBorderStyle.FixedSingle 'disable user changing window size
        Width = 1000 'sets width of window
        Height = 900 'sets height of window
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
        Main_Menu.RunTimeContructor(CreditLabel, 0, 848, 984, 13, "James Cracknell - 191673", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 8.25, Cursors.Default, Color.Transparent)
        Controls.Add(CreditLabel)
        ' StartGameButton
        Main_Menu.RunTimeContructor(StartGameButton, 392, 110, 200, 30, "Start Game", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12.25, Cursors.Hand, Color.Transparent)
        StartGameButton.BackColor = Color.LightGreen
        AddHandler StartGameButton.Click, AddressOf StartGameButton_Click
        Controls.Add(StartGameButton)
        AddHandler Me.Paint, AddressOf Board_Paint
    End Sub
    Public Sub Board_Setup()
        Dim XPos, YPos As Integer
        ' Board
        XPos = 150
        For x = 0 To 6
            YPos = 235
            For y = 0 To 5
                BoardLocations(x, y) = New PictureBox
                BoardLocations(x, y) = New PictureBox
                BoardLocations(x, y).Location = New Point(XPos, YPos)
                BoardLocations(x, y).Size = New Size(90, 90)
                BoardLocations(x, y).SizeMode = PictureBoxSizeMode.Zoom ' Set to fully display image without warping it
                BoardLocations(x, y).BackColor = Color.Blue 'since I cannot make it transparent, I made the background blue to match the box
                BoardLocations(x, y).Enabled = False
                BoardLocations(x, y).Tag = (x * 6 + y)
                AddHandler BoardLocations(x, y).Click, AddressOf Board_Click
                AddHandler BoardLocations(x, y).MouseHover, AddressOf Board_Hover
                Controls.Add(BoardLocations(x, y))
                YPos += 100
            Next
            XPos += 100
        Next

        ' Hover Indicators
        XPos = 145
        YPos = 160
        For i = 0 To 6
            LocationIndicators(i) = New PictureBox
            LocationIndicators(i).Location = New Point(XPos, YPos)
            LocationIndicators(i).Size = New Size(100, 65)
            LocationIndicators(i).SizeMode = PictureBoxSizeMode.Zoom ' Set to fully display image without warping it
            Controls.Add(LocationIndicators(i))
            XPos += 100
        Next

        ' Other Elements
        Main_Menu.RunTimeContructor(CurrentMoveLabel, 745, 60, 150, 50, "Current Move:" & vbNewLine & CurrentGame.GetCurrentMove, ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.LightGray)
        Controls.Add(CurrentMoveLabel)

        Main_Menu.RunTimeContructor(CurrentPlayerLabel, 745, 113, 150, 50, "", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.LightGray)
        If StartingPlayer = "Player One" Then
            CurrentPlayerLabel.Text = "Current Player:" & vbNewLine & CurrentGame.GetPlayerOneName
        Else
            CurrentPlayerLabel.Text = "Current Player:" & vbNewLine & CurrentGame.GetPlayertwoName
        End If
        Controls.Add(CurrentPlayerLabel)

    End Sub
    Public Sub ToggleBoardInteractivity()
        If BoardEnabled = True Then
            BoardEnabled = False
            For x = 0 To 6
                For y = 0 To 5
                    BoardLocations(x, y).Enabled = False
                    BoardLocations(x, y).Cursor = Cursors.Default
                Next
            Next
        Else
            BoardEnabled = True
            For x = 0 To 6
                For y = 0 To 5
                    BoardLocations(x, y).Enabled = True
                    BoardLocations(x, y).Cursor = Cursors.Hand
                Next
            Next
        End If
    End Sub
    Public Sub StartGameButton_Click()
        Controls.Remove(StartGameButton)
        CurrentGame.startgame
    End Sub
    Public Function GetBoardInteractivity()
        Return BoardEnabled
    End Function
    Public Sub Computer_Form_SetUp() 'add elements exclusive to computer vs player game
        ' computer making move indicator
    End Sub
    Public Sub Player_Form_Setup(SubType As String, CountdownTime As String)
        ' set up player names
        ' call subtype games

        ' PlayerOneNameLabel
        Main_Menu.RunTimeContructor(PlayerOneNameLabel, 292, 75, 150, 30, CStr(CurrentGame.GetPlayerOneName()), ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        PlayerOneNameLabel.Visible = False
        Controls.Add(PlayerOneNameLabel)

        ' PlayerTwoNameLabel
        Main_Menu.RunTimeContructor(PlayerTwoNameLabel, 542, 75, 150, 30, CStr(CurrentGame.GetPlayerTwoName()), ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        PlayerTwoNameLabel.Visible = False
        Controls.Add(PlayerTwoNameLabel)
        If SubType.ToLower <> "no timer" Then
            ' PlayerOneTimerLabel
            Main_Menu.RunTimeContructor(PlayerOneTimerLabel, 292, 100, 150, 30, "0.00", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
            PlayerOneTimerLabel.Visible = False
            Controls.Add(PlayerOneTimerLabel)

            'PlayerTwoTimerLabel
            Main_Menu.RunTimeContructor(PlayerTwoTimerLabel, 542, 100, 150, 30, "0.00", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
            PlayerTwoTimerLabel.Visible = False
            Controls.Add(PlayerTwoTimerLabel)
        End If
    End Sub
    Public Sub MakeTimersVisible()
        PlayerOneNameLabel.Visible = True
        PlayerTwoNameLabel.Visible = True
        PlayerOneTimerLabel.Visible = True
        PlayerTwoTimerLabel.Visible = True
    End Sub
    Public Sub DisableTimers()
        PlayerOneTimer.Enabled = False
        PlayerTwoTimer.Enabled = False
        PlayerOneNameLabel.Visible = False
        PlayerTwoNameLabel.Visible = False
        PlayerOneTimerLabel.Visible = False
        PlayerTwoTimerLabel.Visible = False
    End Sub
    Public Sub Countdown_Game_Setup(SubType As String, CountdownValue As Integer)
        TimerCreate(SubType)
        SetTimerValues(CountdownValue)
        PlayerOneTimerLabel.Refresh()
        PlayerTwoTimerLabel.Refresh()
    End Sub
    Public Sub Timed_Game_Setup(SubType As String)
        TimerCreate(SubType)
        SetTimerValues(0)
    End Sub
    Private Sub TimerCreate(SubType As String)
        If SubType = "Countdown" Then
            RemoveHandler PlayerOneTimer.Tick, AddressOf CountupTimer_Tick
            RemoveHandler PlayerTwoTimer.Tick, AddressOf CountupTimer_Tick
            AddHandler PlayerOneTimer.Tick, AddressOf CountdownTimer_Tick
            AddHandler PlayerTwoTimer.Tick, AddressOf CountdownTimer_Tick
        Else
            RemoveHandler PlayerOneTimer.Tick, AddressOf CountdownTimer_Tick
            RemoveHandler PlayerTwoTimer.Tick, AddressOf CountdownTimer_Tick
            AddHandler PlayerOneTimer.Tick, AddressOf CountupTimer_Tick
            AddHandler PlayerTwoTimer.Tick, AddressOf CountupTimer_Tick
        End If
        PlayerOneTimer.Interval = 10 '0.01 second interval
        PlayerTwoTimer.Interval = 10
    End Sub
    Private Sub SetTimerValues(CountdownValue As Integer)
        PlayerOneTimerValue = CountdownValue ' Initial value for timer chosen by user, 0 if counting up.
        PlayerTwoTimerValue = CountdownValue
        PlayerOneTimerLabel.Text = PlayerOneTimerValue
        PlayerTwoTimerLabel.Text = PlayerTwoTimerValue
    End Sub
    Public Sub CountdownTimer_Tick(sender As Object, e As EventArgs) 'functionality must be added to swap the user
        If PlayerOneTimer.Enabled = True Then 'if player one timer is counting up
            PlayerOneTimerValue -= 0.01
            PlayerOneTimerLabel.Text = PlayerOneTimerValue
        Else 'if player two timer is counting 
            PlayerTwoTimerValue -= 0.01
            PlayerTwoTimerLabel.Text = PlayerTwoTimerValue
        End If

        If PlayerOneTimerValue <= 0 Or PlayerTwoTimerValue <= 0 Then
            PlayerOneTimer.Enabled = False
            PlayerTwoTimer.Enabled = False
            CurrentGame.GameOver("TimeOut")
        End If
    End Sub
    Public Sub CountupTimer_Tick(sender As Object, e As EventArgs)
        If PlayerOneTimer.Enabled = True Then 'if player one timer is counting up
            PlayerOneTimerValue += 0.01
            PlayerOneTimerLabel.Text = PlayerOneTimerValue
        Else 'if player two timer is counting 
            PlayerTwoTimerValue += 0.01
            PlayerTwoTimerLabel.Text = PlayerTwoTimerValue
        End If
    End Sub
    Public Function GetPlayerOneTimerStatus() As String
        If PlayerOneTimer.Enabled = True Then
            Return ("Enabled")
        Else
            Return ("Disabled")
        End If
    End Function
    Public Sub SetPlayerOneTimerStatus(NewStatus As String)
        If NewStatus = "Enable" Then
            PlayerOneTimer.Enabled = True
        Else
            PlayerOneTimer.Enabled = False
        End If
    End Sub
    Public Sub SetPlayerTwoTimerStatus(NewStatus As String)
        If NewStatus = "Enable" Then
            PlayerTwoTimer.Enabled = True
        Else
            PlayerTwoTimer.Enabled = False
        End If
    End Sub
    Public Function GetPlayerOneTimerValue() As Decimal
        Return PlayerOneTimerValue
    End Function
    Public Function GetPlayerTwoTimerValue() As Decimal
        Return PlayerTwoTimerValue
    End Function
    Private Sub Board_Paint(sender As Object, e As PaintEventArgs)
        If BoardPainted = False Then 'prevents the board being redrawn. It should only be drawn once.
            BoardPainted = True
            Dim g As Graphics
            g = Me.CreateGraphics

            ' Rectangle Background
            Dim myBrush As New SolidBrush(Color.Blue) 'creates a new blue brush 
            Dim formGraphics As Graphics
            formGraphics = Me.CreateGraphics()
            formGraphics.FillRectangle(myBrush, New Rectangle(145, 230, 700, 600))
            myBrush.Dispose()
            formGraphics.Dispose()
            Dim BlackPen As New Pen(Brushes.Black)
            BlackPen.Width = 6.0F
            ' Rectangle Outline
            g.DrawRectangle(BlackPen, 145, 230, 700, 600) '(pentype, startingx, startingy, xsize, ysize)

            'vertical lines
            Dim x1 As Integer = 245
            For i = 0 To 6
                g.DrawLine(BlackPen, x1, 230, x1, 830) '(pentype, startingx, startingy, endingx, endingy)
                x1 += 100
            Next

            'horizontal lines
            Dim y1 As Integer = 330
            For j = 0 To 5
                g.DrawLine(BlackPen, 145, y1, 845, y1)
                y1 += 100
            Next
        End If
    End Sub
    Public Sub Board_Update()
        For x = 0 To 6
            For y = 0 To 5
                If CurrentGame.GetFilledStatus(x, y) = "Y" Then
                    BoardLocations(x, y).Image = Image.FromFile("YellowToken.png")
                ElseIf CurrentGame.GetFilledStatus(x, y) = "R" Then
                    BoardLocations(x, y).Image = Image.FromFile("RedToken.png")
                Else
                    BoardLocations(x, y).Image = Nothing 'if an image is there that shouldn't be there it is cleared (such as when the game is reset)
                End If
            Next
        Next
    End Sub
    Private Sub Board_Click(sender As Object, e As EventArgs)
        If CurrentGame.make_move(sender.tag) Then
            CurrentGame.SwitchMove()
        End If
    End Sub
    Private Sub Board_Hover(sender As Object, e As EventArgs)
        For i = 0 To 6
            LocationIndicators(i).Image = Nothing
        Next
        LocationIndicators(CInt(sender.tag) \ 6).Image = Image.FromFile("Arrow.png")
        Controls.Add(LocationIndicators(CInt(sender.tag) \ 6))
    End Sub
    Public Sub UpdateCurrentPlayerVisual()
        If CurrentGame.GetCurrentColour() = CurrentGame.GetPlayerOneColour Then
            CurrentPlayerLabel.Text = "Current Player:" & vbNewLine & CurrentGame.GetPlayerOneName
        Else 'player two turn
            CurrentPlayerLabel.Text = "Current Player:" & vbNewLine & CurrentGame.GetPlayerTwoName
        End If
    End Sub
    Public Function GetCurrentPlayer()
        Return Replace(CStr(CurrentPlayerLabel.Text), "Current Player:", "") 'replace command is used to clean up the output
    End Function
    Public Sub UpdateCurrentMoveVisual()
        CurrentMoveLabel.Text = "Current Move:" & vbNewLine & CurrentGame.getCurrentMove
    End Sub
    Public Function GetStartingPlayer() As String
        Return StartingPlayer
    End Function
    Public Function GetGameType()
        Return CurrentGameType
    End Function
End Class

Public Class Game
    Protected BoardState(6, 5) As Char
    Protected PlayerOneColour As Colours
    Protected PlayerOneName As String
    Protected PlayerTwoColour As Colours
    Protected PlayerTwoName As String
    Private MoveNumber As Integer
    Private CurrentColour As Colours
    Private CurrentMove As Integer
    Private GameOverIndicator As Boolean
    Enum Colours
        Red
        Yellow
    End Enum
    Public Sub New(SubType As String, PlayerOneUsername As String, PlayerOneColour As String, PlayerTwoUsername As String, PlayerTwoColour As String)
        If PlayerOneColour = "Yellow" Then 'set player colour to their choice from previous menu
            SetPlayerOneColour(Colours.Yellow)
            SetPlayerTwoColour(Colours.Red)
        Else
            SetPlayerOneColour(Colours.Red)
            SetPlayerTwoColour(Colours.Yellow)
        End If
        PlayerOneName = PlayerOneUsername
        PlayerTwoName = PlayerTwoUsername
        GameOverIndicator = False
        CurrentMove = 1
        For x = 0 To 6 'resets grid
            For y = 0 To 5
                BoardState(x, y) = Nothing
            Next
        Next
    End Sub
    Public Function GetFilledStatus(x, y) ' return is empty if the space is empty
        If BoardState(x, y) = "Y" Then
            Return "Y"
        ElseIf BoardState(x, y) = "R" Then
            Return "R"
        End If
    End Function
    Public Sub SetPlayerOneColour(Colour)
        PlayerOneColour = Colour
    End Sub
    Public Overridable Function CheckForWinningState(BoardState, CurrentPlayer) 'searches 2D array of char (boardstate) for 4 in a row
        Dim WinDiscovered As Boolean = False
        Dim CurrentPlayerChar As Char = GetColourChar(CurrentPlayer)
        ' Verticals
        For y = 0 To 2
            For x = 0 To 6
                If BoardState(x, y) = BoardState(x, y + 1) And BoardState(x, y + 1) = BoardState(x, y + 2) And BoardState(x, y + 2) = BoardState(x, y + 3) And BoardState(x, y) = CurrentPlayerChar Then 'cant be blank
                    'if a sequence of 4 is detected, it must be a win for the current player, else it would have been detected by a previous check.
                    WinDiscovered = True
                End If
            Next
        Next
        ' Horizontals
        For x = 0 To 3
            For y = 0 To 5
                If BoardState(x, y) = BoardState(x + 1, y) And BoardState(x + 1, y) = BoardState(x + 2, y) And BoardState(x + 2, y) = BoardState(x + 3, y) And BoardState(x, y) = CurrentPlayerChar Then 'if all 4 are equal
                    WinDiscovered = True
                End If
            Next
        Next
        ' Diagonals
        For x = 0 To 3
            For y = 0 To 2
                If BoardState(x, y) = BoardState(x + 1, y + 1) And BoardState(x + 1, y + 1) = BoardState(x + 2, y + 2) And BoardState(x + 2, y + 2) = BoardState(x + 3, y + 3) And BoardState(x, y) = CurrentPlayerChar Then
                    WinDiscovered = True
                End If
            Next
        Next

        For x = 0 To 3
            For y = 3 To 5
                If BoardState(x, y) = BoardState(x + 1, y - 1) And BoardState(x + 1, y - 1) = BoardState(x + 2, y - 2) And BoardState(x + 2, y - 2) = BoardState(x + 3, y - 3) And BoardState(x, y) = CurrentPlayerChar Then
                    WinDiscovered = True
                End If
            Next
        Next
        Return WinDiscovered
    End Function
    Public Function CheckForDrawingState(BoardState) 'checks to see if the board is full
        Dim Full As Boolean = True 'assumes board is full
        For x = 0 To 6
            For y = 0 To 5
                If BoardState(x, y) = Nothing Then 'if an empty space is found
                    Full = False 'board is hence not full
                End If
            Next
        Next
        Return Full
    End Function
    Public Sub GameOver(WinMode As String)
        Game_Form.ToggleBoardInteractivity()
        ' record game
        Game_Form.Board_Update()
        If WinMode.ToLower = "timeout" Then
            ' opposite player wins
            If Game_Form.GetCurrentPlayer() = PlayerOneName Then
                MsgBox(Game_Form.GetCurrentPlayer() & " ran out of time. " & vbNewLine & PlayerTwoName & " has won the game.", MsgBoxStyle.DefaultButton1, "Game Ended")
            Else
                MsgBox(Game_Form.GetCurrentPlayer() & " ran out of time. " & vbNewLine & PlayerOneName & " has won the game.", MsgBoxStyle.DefaultButton1, "Game Ended")
            End If
        ElseIf CheckForDrawingState(BoardState) Then 'game is drawn
            GameOverIndicator = True
            If WinMode = "Timed Game" Then 'If it is a timed game, the player with the lowest time wins
                If Game_Form.GetPlayerOneTimerValue >= Game_Form.GetPlayerTwoTimerValue Then 'if users had same time left (very unlikely), player one wins
                    MsgBox("The game was a draw. " & PlayerTwoName & " completed their moves quicker, so wins.", MsgBoxStyle.DefaultButton1, "Game Ended")
                ElseIf Game_Form.GetPlayerOneTimerValue < Game_Form.GetPlayerTwoTimerValue Then
                    MsgBox("The game was a draw." & PlayerOneName & " completed their moves quicker, so wins.", MsgBoxStyle.DefaultButton1, "Game Ended")
                End If
            Else
                MsgBox("The game was a draw.", MsgBoxStyle.DefaultButton1, "Game Ended")
            End If
        Else
            GameOverIndicator = True
            MsgBox(Game_Form.GetCurrentPlayer() & vbNewLine & "has won the game.", MsgBoxStyle.DefaultButton1, "Game Ended")
        End If
        Game_Form.Hide()
        Main_Menu.Show()
    End Sub
    Public Function GetGameOverIndicator()
        Return GameOverIndicator
    End Function
    Public Function Make_Move(Location)
        Dim Column As Integer = (CInt(Location)) \ 6
        Dim MoveMade As Boolean = False
        For y = 5 To 0 Step -1 'counts backwards (up)
            If GetFilledStatus(Column, y) <> "Y" And GetFilledStatus(Column, y) <> "R" And MoveMade = False Then
                BoardState(Column, y) = GetColourChar(GetCurrentColour())
                MoveMade = True
            End If
        Next
        Return MoveMade
    End Function
    Public Overridable Sub SwitchMove() '
        UpdateCurrentMove()
        Game_Form.Board_Update()
        Game_Form.UpdateCurrentMoveVisual()
        If CheckForWinningState(BoardState, CurrentColour) Or CheckForDrawingState(BoardState) Then 'game over
            GameOver(Game_Form.GetGameType())
        Else
            If CurrentColour = Colours.Red Then
                CurrentColour = Colours.Yellow
            Else
                CurrentColour = Colours.Red
            End If
            Game_Form.UpdateCurrentPlayerVisual()
        End If
    End Sub
    Public Function GetCurrentMove()
        Return CurrentMove
    End Function
    Public Function UpdateCurrentMove()
        CurrentMove += 1
    End Function
    Public Function GetPlayerOneName()
        Return PlayerOneName
    End Function
    Public Function GetPlayerTwoName()
        Return PlayerTwoName
    End Function
    Public Function GetPlayerOneColour()
        Return PlayerOneColour
    End Function
    Public Overridable Sub SetPlayerTwoColour(Colour)
        PlayerTwoColour = Colour
    End Sub
    Public Function SetCurrentColour(CurrentColourInput)
        CurrentColour = CurrentColourInput
    End Function
    Public Function GetCurrentColour()
        Return CurrentColour
    End Function
    Public Function GetColourChar(ColourChoice)
        If ColourChoice = Colours.Red Then
            Return "R"
        Else
            Return "Y"
        End If
    End Function
    Public Overridable Function GetPlayerTwoColour()
        Return PlayerTwoColour
    End Function
End Class

Public Class Player_Vs_Computer_Game
    Inherits Game
    Private InitialDepth As Integer
    Private Returns
    Private ComputerColour As Colours
    Private ChanceOfMistake As Decimal
    Public Sub New(Difficulty As String, Username As String, ChosenColour As String, StartingPlayer As String)
        MyBase.New(Difficulty, Username, ChosenColour, "Computer", "")
        SetDepth(Difficulty) 'this must be set based on difficulty.
        If StartingPlayer = Username Then
            SetCurrentColour(PlayerOneColour)
        ElseIf StartingPlayer = "Computer Player" Then
            SetCurrentColour(ComputerColour)
        Else
            If (CInt(Math.Floor(Rnd() * 2))) = 0 Then 'random number, either 1 or 0 resulting in 50% chance to be either player
                SetCurrentColour(PlayerOneColour)
            Else
                SetCurrentColour(ComputerColour)
            End If
        End If
    End Sub
    Public Sub SetDepth(SpecifiedDifficulty) 'sets attributes based on difficulty
        If SpecifiedDifficulty = "Easy Difficulty" Then
            InitialDepth = 2
            ChanceOfMistake = 25
        ElseIf SpecifiedDifficulty = "Medium Difficulty" Then
            ChanceOfMistake = 10
            InitialDepth = 3
        ElseIf SpecifiedDifficulty = "Hard Difficulty" Then
            ChanceOfMistake = 5
            InitialDepth = 4
        Else 'Impossible difficulty
            ChanceOfMistake = 0
            InitialDepth = 5
        End If
    End Sub
    Public Sub ComputerMove()
        Dim Returns
        Dim BestXCoordinate As Integer
        Dim RandomMoveMade As Boolean
        Returns = Minimax(BoardState, Double.NegativeInfinity, Double.PositiveInfinity, 4, True)
        If ChanceOfMistake <= CInt((100) * Rnd()) Then 'random number from 0 to 100. If it is smaller than the chance of mistake then it makes a mistake and picks a random spot.
            BestXCoordinate = Returns.item2
            MakeComputerMove(BestXCoordinate)
        Else
            ' random spot
            MsgBox("random move")
            Do
                RandomMoveMade = True
                BestXCoordinate = CInt((6) * Rnd())
                If MakeComputerMove(BestXCoordinate) = False Then
                    RandomMoveMade = False 'column was full, so tries again
                End If
            Loop Until RandomMoveMade
        End If
        SwitchMove()
    End Sub
    Public Function MakeComputerMove(XCoordinate)
        Dim MoveMade As Boolean = False
        For y = 5 To 0 Step -1 'counts backwards (up)
            If GetFilledStatus(XCoordinate, y) <> "Y" And GetFilledStatus(XCoordinate, y) <> "R" And MoveMade = False Then
                BoardState(XCoordinate, y) = GetColourChar(ComputerColour)
                MoveMade = True
            End If
        Next
        Return MoveMade
    End Function
    Public Function Minimax(ByVal CurrentBoardState(,) As Char, Alpha As Double, Beta As Double, Depth As Integer, MaximisingPlayer As Boolean)
        Dim Evaluation
        Dim BestEvaluation As Double
        Dim BestColumn As Integer
        Dim Searched As Boolean = False
        Dim YCoordinate As Integer
        Dim Order() As Integer = {3, 4, 2, 5, 1, 6, 0} 'the order that board states should be searched
        If Depth = 0 Then  'if it has reached the end of its search depth
            Return (StaticEvaluation(CurrentBoardState), Nothing)
        Else
            If CheckForDrawingState(CurrentBoardState) Then 'if it is a draw (board is full)
                Return (0, Nothing)
            Else
                If CheckForWinningState(CurrentBoardState, GetPlayerTwoColour) Then 'computer got a 4 in a row
                    Return (10000000, Nothing)
                ElseIf CheckForWinningState(CurrentBoardState, GetPlayerOneColour) Then 'player got a 4 in a row
                    Return (-1000000, Nothing)
                End If
            End If
        End If

        If MaximisingPlayer Then 'computer move
            BestEvaluation = Double.NegativeInfinity
            For x = 0 To 6
                Searched = False
                YCoordinate = 5
                Do
                    If CurrentBoardState(x, YCoordinate) = Nothing And Searched = False Then
                        Searched = True ' so loop only runs once per column
                        CurrentBoardState(x, YCoordinate) = GetColourChar(GetPlayerTwoColour())
                        Evaluation = Minimax(CurrentBoardState, Alpha, Beta, Depth - 1, False)
                        If (Evaluation.item1 + Depth) > BestEvaluation Then 'if it is less moves in, the score is higher. if two moves win, the shortest win is favoured.
                            BestColumn = x
                            BestEvaluation = Evaluation.item1 + Depth
                        End If
                        CurrentBoardState(x, YCoordinate) = Nothing
                    End If
                    YCoordinate -= 1
                Loop Until YCoordinate = 0 Or Searched = True
                Alpha = Math.Max(Alpha, BestEvaluation)
                If Alpha >= Beta Then
                    ' Exit For
                    ' Return (BestEvaluation, BestColumn)
                End If
            Next
            Return (BestEvaluation, BestColumn)
        Else
            BestEvaluation = Double.PositiveInfinity
            For x = 0 To 6
                Searched = False
                YCoordinate = 5
                Do
                    If CurrentBoardState(x, YCoordinate) = Nothing And Searched = False Then
                        Searched = True ' so loop only runs once
                        CurrentBoardState(x, YCoordinate) = GetColourChar(GetPlayerOneColour())
                        Evaluation = Minimax(CurrentBoardState, Alpha, Beta, Depth - 1, True)
                        If (Evaluation.item1 + Depth) < BestEvaluation Then 'if it is less moves in, the score is higher. if two moves win, the shortest win is favoured. 
                            BestColumn = x
                            BestEvaluation = Evaluation.item1 + Depth
                        End If
                        CurrentBoardState(x, YCoordinate) = Nothing
                    End If
                    YCoordinate -= 1
                Loop Until YCoordinate = 0 Or Searched = True
                Beta = Math.Min(Beta, BestEvaluation)
                If Alpha >= Beta Then
                    'Exit For
                    ' Return (BestEvaluation, BestColumn)
                End If
            Next
            Return (BestEvaluation, BestColumn)
        End If
    End Function
    Public Function EvaluateScore(EvaluatingSection() As Char, AnalysisColour As Char)
        Dim NumberOfTilesInSequence = 0
        Dim NumberOfOppositeTilesInSequence = 0
        Dim SectionScore As Integer
        For i = 0 To 3
            If EvaluatingSection(i) <> AnalysisColour And EvaluatingSection(i) <> Nothing Then 'if the sequence has the opponents colour anywhere in it, it is not viable
                NumberOfOppositeTilesInSequence += 1
                ' Return 0
            ElseIf EvaluatingSection(i) = AnalysisColour Then
                NumberOfTilesInSequence += 1
            End If
        Next
        If NumberOfTilesInSequence = 1 Then '1/4 of the tiles is filled 
            SectionScore += 1
        ElseIf NumberOfTilesInSequence = 2 Then '2/4 of the tiles is filled
            SectionScore += 10
        ElseIf NumberOfTilesInSequence = 3 Then '3/4 of the tiles is filled
            SectionScore += 50
        ElseIf NumberOfTilesInSequence = 4 Then '4/4 tiles are filled by one colour
            SectionScore += 2000
        End If
        If NumberOfOppositeTilesInSequence = 3 And NumberOfTilesInSequence = 0 Then '3 of opposite player, unblocked
            SectionScore -= 18
        End If
        Return SectionScore
    End Function
    Public Function StaticEvaluation(EvaluatingBoardState(,) As Char) 'returns the score of that board state
        Dim EvaluatedScore As Integer = 0
        Dim AnalysisColour As Char
        Dim AnalysisSection(3) As Char
        Dim CentreCount As Integer
        Dim NextToCentreCount As Integer
        Dim TwoFromCentreCount As Integer
        AnalysisColour = GetColourChar(GetPlayerTwoColour())

        For Y = 5 To 0 Step -1  'more central moves are worth more
            If EvaluatingBoardState(1, Y) = AnalysisColour Then TwoFromCentreCount += 1
            If EvaluatingBoardState(2, Y) = AnalysisColour Then NextToCentreCount += 1
            If EvaluatingBoardState(3, Y) = AnalysisColour Then CentreCount += 1
            If EvaluatingBoardState(4, Y) = AnalysisColour Then NextToCentreCount += 1
            If EvaluatingBoardState(5, Y) = AnalysisColour Then TwoFromCentreCount += 1
        Next
        EvaluatedScore += CentreCount * 3
        EvaluatedScore += NextToCentreCount
        EvaluatedScore += TwoFromCentreCount * 0.5

        ' Verticals
        For X = 0 To 6
            For StartingY = 0 To 2
                For Increase = 0 To 3
                    AnalysisSection(Increase) = EvaluatingBoardState(X, StartingY + Increase)
                Next
                EvaluatedScore += EvaluateScore(AnalysisSection, AnalysisColour)
            Next
        Next

        ' Horizontals
        For Y = 0 To 5
            For StartingX = 0 To 3
                For Increase = 0 To 3
                    AnalysisSection(Increase) = EvaluatingBoardState(StartingX + Increase, Y)
                Next
                EvaluatedScore += EvaluateScore(AnalysisSection, AnalysisColour)
            Next
        Next

        ' Diagonals

        For X = 0 To 3
            For Y = 0 To 2
                For Increase = 0 To 3
                    AnalysisSection(Increase) = EvaluatingBoardState(X + Increase, Y + Increase)
                Next
                EvaluatedScore += EvaluateScore(AnalysisSection, AnalysisColour)
            Next

        Next

        For X = 0 To 3
            For Y = 3 To 5
                For Increase = 0 To 3
                    AnalysisSection(Increase) = EvaluatingBoardState(X + Increase, Y - Increase)
                Next
                EvaluatedScore += EvaluateScore(AnalysisSection, AnalysisColour)
            Next
        Next
        Return EvaluatedScore
    End Function
    Public Sub StartGame()
        If GetCurrentColour() = PlayerTwoColour Then 'Computer Move
            ComputerMove()
        Else 'player move
            Game_Form.ToggleBoardInteractivity() 'activate buttons
        End If
    End Sub
    Public Overrides Sub SwitchMove()
        MyBase.SwitchMove()
        Game_Form.ToggleBoardInteractivity()
        If GetCurrentColour() = PlayerTwoColour And GetGameOverIndicator() = False Then 'PlayerOnesMove
            ComputerMove()
        End If
    End Sub
    Public Overrides Sub SetPlayerTwoColour(Colour)
        MyBase.SetPlayerTwoColour(Colour)
        ComputerColour = Colour
    End Sub
    Public Overrides Function GetPlayerTwoColour()
        MyBase.GetPlayerTwoColour()
        Return ComputerColour
    End Function
End Class

Class Player_Vs_Player_Game
    Inherits Game
    Private GameSubType As String
    Public Sub New(SubType As String, PlayerOneUsername As String, PlayerOneChosenColour As String, PlayerTwoUsername As String, PlayerTwoChosenColour As String, StartingPlayer As String)
        MyBase.New(SubType, PlayerOneUsername, PlayerOneChosenColour, PlayerTwoUsername, PlayerTwoChosenColour)
        GameSubType = SubType
        If StartingPlayer = PlayerOneUsername Then
            SetCurrentColour(PlayerOneColour)
        ElseIf StartingPlayer = PlayerTwoUsername Then
            SetCurrentColour(PlayerTwoColour)
        End If
    End Sub
    Public Sub StartGame()
        Game_Form.MakeTimersVisible()
        If Game_Form.GetStartingPlayer() = PlayerOneName Then
            Game_Form.SetPlayerOneTimerStatus("Enable")
        Else
            Game_Form.SetPlayerTwoTimerStatus("Enable")
        End If

        'End If
        ' start starting players timer
        Game_Form.ToggleBoardInteractivity() 'activate buttons
    End Sub
    Public Overrides Sub SwitchMove()
        MyBase.SwitchMove()
        If Game_Form.GetPlayerOneTimerStatus() = "Enabled" Then 'if player one timer is counting up
            Game_Form.SetPlayerOneTimerStatus("Disable")
            Game_Form.SetPlayerTwoTimerStatus("Enable")
        Else 'if player two timer is counting 
            Game_Form.SetPlayerOneTimerStatus("Enable")
            Game_Form.SetPlayerTwoTimerStatus("Disable")
        End If
    End Sub
End Class