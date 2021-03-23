Imports System.IO
Public Class Game_Form ' The main code for the game

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
    Private ComputerDifficultyLabel As New Label

    ' Variable Declarations

    Private BoardEnabled As Boolean = False
    Private StartingPlayer As String
    Private FirstTimeRunning As Boolean = True ' some code only needs to run when it is not the first time 
    Private PreviousGameType As String
    Private BoardPainted As Boolean = False 'only allows board to be painted once a game
    Private PlayerOneTimerValue As Decimal
    Private PlayerTwoTimerValue As Decimal
    Private CurrentGameType As String
    Private CurrentSubType As String

    ' Code

    Public Sub Game_Setup(GameType As String, SubType As String, PlayerOneName As String, PlayerOneColour As String, PlayerTwoName As String, PlayerTwoColour As String, GameStartingPlayer As String, CountdownTime As Decimal) 'Creates the game object, based on parameters passed in
        Me.Show()
        StartingPlayer = GameStartingPlayer
        CurrentGameType = GameType
        CurrentSubType = SubType
        PlayerOneTimerValue = 0
        PlayerTwoTimerValue = 0
        If GameType = "Computer" Then
            CurrentGame = New Player_Vs_Computer_Game(SubType, PlayerOneName, PlayerOneColour, GameStartingPlayer)
            Computer_Form_Setup(SubType) 'set up player vs computer specific elements
            If FirstTimeRunning Then 'some code only needs to run if it is the first time it is being run
                Board_Setup() 'creates board
                This_Form_Setup() 'sets up form
            Else
                Board_Update() 'resets board
                UpdateCurrentMoveVisual()
                Controls.Add(StartGameButton)
                If PreviousGameType <> "Computer" Then 'previous game was NOT a computer game
                    DisableTimers()
                    RemoveHandler PlayerOneTimer.Tick, AddressOf CountupTimer_Tick 'removes previous timer handlers (from previous games played) so that they cannot interfere with each other
                    RemoveHandler PlayerTwoTimer.Tick, AddressOf CountupTimer_Tick
                    RemoveHandler PlayerOneTimer.Tick, AddressOf CountdownTimer_Tick
                    RemoveHandler PlayerTwoTimer.Tick, AddressOf CountdownTimer_Tick
                End If
            End If
        ElseIf GameType = "Player" Then
            CurrentGame = New Player_Vs_Player_Game(SubType, PlayerOneName, PlayerOneColour, PlayerTwoName, PlayerTwoColour, GameStartingPlayer)
            If FirstTimeRunning Then
                Board_Setup() 'creates board
                This_Form_Setup() 'sets up form
                Player_Form_Setup(SubType, CountdownTime) 'set up player vs player specific elements
            Else
                Controls.Remove(ComputerDifficultyLabel)
                DisableTimers() 'turns off timets
                Board_Update() 'rests board
                UpdateCurrentMoveVisual() 'resets current move
                SetTimerValues(CountdownTime) 'sets up timers
                Controls.Add(StartGameButton)
            End If
            If SubType = "Countdown" Then
                Countdown_Game_Setup(SubType, CountdownTime)
            ElseIf SubType = "Timed Game" Then
                Timed_Game_Setup(SubType)
            Else
                DisableTimers()
            End If

        End If
        BoardPainted = False
        PreviousGameType = GameType
        FirstTimeRunning = False 'first time running set to false after all start up code has run
    End Sub
    Private Sub This_Form_Setup() 'sets this forms properties
        Processes.Form_Setup(Me)
        Height = 900 'this form is bigger
        ' Form Title
        Title.Location = New Point(242, 5)
        Title.Size = New Size(500, 50)
        Title.SizeMode = PictureBoxSizeMode.Zoom
        Title.Image = Image.FromFile("Title.png")
        Controls.Add(Title)
        Processes.RunTimeContructor(CreditLabel, 0, 848, 984, 13, "James Cracknell - 191673", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 8.25, Cursors.Default, Color.Transparent)
        Controls.Add(CreditLabel)
        ' StartGameButton
        Processes.RunTimeContructor(StartGameButton, 392, 110, 200, 30, "Start Game", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12.25, Cursors.Hand, Color.Transparent)
        StartGameButton.BackColor = Color.LightGreen
        AddHandler StartGameButton.Click, AddressOf StartGameButton_Click
        Controls.Add(StartGameButton)
        AddHandler Me.Paint, AddressOf Board_Paint
    End Sub
    Private Sub Board_Setup() 'creates the board
        Dim XPos, YPos As Integer
        ' Board
        XPos = 146
        For x = 0 To 6
            YPos = 231
            For y = 0 To 5
                BoardLocations(x, y) = New PictureBox
                BoardLocations(x, y) = New PictureBox
                BoardLocations(x, y).Location = New Point(XPos, YPos)
                BoardLocations(x, y).Size = New Size(98, 98)
                BoardLocations(x, y).SizeMode = PictureBoxSizeMode.Zoom ' Set to fully display image without warping it
                BoardLocations(x, y).BackColor = Color.Blue 'since I cannot make it transparent, I made the background blue to match the box
                BoardLocations(x, y).Enabled = False
                BoardLocations(x, y).Tag = (x * 6 + y)
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

        ' CurrentMoveLabel
        Processes.RunTimeContructor(CurrentMoveLabel, 745, 60, 150, 50, "Current Move:" & vbNewLine & CurrentGame.GetCurrentMove, ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.LightGray)
        Controls.Add(CurrentMoveLabel)

        ' CurrentPlayerLabel
        Processes.RunTimeContructor(CurrentPlayerLabel, 745, 113, 150, 50, "", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.LightGray)

        If StartingPlayer = CurrentGame.GetPlayerOneName Then 'sets initial player label text
            CurrentPlayerLabel.Text = "Current Player:" & vbNewLine & CurrentGame.GetPlayerOneName
        Else
            CurrentPlayerLabel.Text = "Current Player:" & vbNewLine & CurrentGame.GetPlayertwoName
        End If
        Controls.Add(CurrentPlayerLabel)

    End Sub
    Public Sub ToggleBoardInteractivity() ' switches between board enabled and board disabled
        If BoardEnabled = True Then
            BoardEnabled = False
            For x = 0 To 6
                For y = 0 To 5
                    RemoveHandler BoardLocations(x, y).Click, AddressOf Board_Click
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
                    AddHandler BoardLocations(x, y).Click, AddressOf Board_Click
                Next
            Next
        End If
    End Sub
    Public Sub TimersStop()
        PlayerOneTimer.Stop()
        PlayerTwoTimer.Stop()
    End Sub
    Private Sub StartGameButton_Click(sender As Object, e As EventArgs) 'when the start button is clicked, starts the game
        Controls.Remove(StartGameButton)
        CurrentGame.startgame(GetSubType())
    End Sub
    Private Function GetBoardInteractivity() As Boolean
        Return BoardEnabled
    End Function
    Private Sub Player_Form_Setup(SubType As String, CountdownTime As String) 'sets up the controls unique to the player form
        ' PlayerOneNameLabel
        Processes.RunTimeContructor(PlayerOneNameLabel, 292, 75, 150, 30, CStr(CurrentGame.GetPlayerOneName()), ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        PlayerOneNameLabel.Visible = False
        Controls.Add(PlayerOneNameLabel)

        ' PlayerTwoNameLabel
        Processes.RunTimeContructor(PlayerTwoNameLabel, 542, 75, 150, 30, CStr(CurrentGame.GetPlayerTwoName()), ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
        PlayerTwoNameLabel.Visible = False
        Controls.Add(PlayerTwoNameLabel)
        If SubType.ToLower <> "no timer" Then
            ' PlayerOneTimerLabel
            Processes.RunTimeContructor(PlayerOneTimerLabel, 292, 100, 150, 30, "0.00", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
            PlayerOneTimerLabel.Visible = False
            Controls.Add(PlayerOneTimerLabel)

            ' PlayerTwoTimerLabel
            Processes.RunTimeContructor(PlayerTwoTimerLabel, 542, 100, 150, 30, "0.00", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 12, Cursors.Default, Color.Transparent)
            PlayerTwoTimerLabel.Visible = False
            Controls.Add(PlayerTwoTimerLabel)
        End If
    End Sub
    Private Sub Computer_Form_Setup(Difficulty)
        ' ComputerDifficultyLabel
        Processes.RunTimeContructor(ComputerDifficultyLabel, 142, 90, 150, 50, "Computer Difficulty:" & vbNewLine & Difficulty, ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 10.25, Cursors.Default, Color.LightGray)
        Controls.Add(ComputerDifficultyLabel)
    End Sub
    Public Sub MakeTimersVisible()
        PlayerOneNameLabel.Visible = True
        PlayerTwoNameLabel.Visible = True
        PlayerOneTimerLabel.Visible = True
        PlayerTwoTimerLabel.Visible = True
    End Sub
    Private Sub DisableTimers()
        PlayerOneTimer.Enabled = False
        PlayerTwoTimer.Enabled = False
        PlayerOneNameLabel.Visible = False
        PlayerTwoNameLabel.Visible = False
        PlayerOneTimerLabel.Visible = False
        PlayerTwoTimerLabel.Visible = False
    End Sub
    Private Sub Countdown_Game_Setup(SubType As String, CountdownValue As Decimal)
        TimerCreate(SubType)
        SetTimerValues(CountdownValue)
        PlayerOneTimerLabel.Refresh()
        PlayerTwoTimerLabel.Refresh()
    End Sub
    Private Sub Timed_Game_Setup(SubType As String)
        TimerCreate(SubType)
        SetTimerValues(0)
    End Sub
    Private Sub TimerCreate(SubType As String)
        RemoveHandler PlayerOneTimer.Tick, AddressOf CountupTimer_Tick 'removes previous timer handlers (from previous games played) so that they cannot interfere with each other
        RemoveHandler PlayerTwoTimer.Tick, AddressOf CountupTimer_Tick
        RemoveHandler PlayerOneTimer.Tick, AddressOf CountdownTimer_Tick
        RemoveHandler PlayerTwoTimer.Tick, AddressOf CountdownTimer_Tick
        If SubType = "Countdown" Then
            AddHandler PlayerOneTimer.Tick, AddressOf CountdownTimer_Tick
            AddHandler PlayerTwoTimer.Tick, AddressOf CountdownTimer_Tick
        Else
            AddHandler PlayerOneTimer.Tick, AddressOf CountupTimer_Tick
            AddHandler PlayerTwoTimer.Tick, AddressOf CountupTimer_Tick
        End If
        PlayerOneTimer.Interval = 10 '0.01 second interval
        PlayerTwoTimer.Interval = 10
    End Sub
    Private Sub SetTimerValues(CountdownValue As Decimal)
        PlayerOneTimerValue = CountdownValue ' Initial value for timer chosen by user, 0 if counting up.
        PlayerTwoTimerValue = CountdownValue
        PlayerOneTimerLabel.Text = PlayerOneTimerValue
        PlayerTwoTimerLabel.Text = PlayerTwoTimerValue
    End Sub
    Private Sub CountdownTimer_Tick(sender As Object, e As EventArgs) 'functionality must be added to swap the user
        If PlayerOneTimer.Enabled = True Then 'if player one timer is counting up
            PlayerOneTimerValue -= 0.01
            PlayerOneTimerLabel.Text = PlayerOneTimerValue
        Else 'if player two timer is counting 
            PlayerTwoTimerValue -= 0.01
            PlayerTwoTimerLabel.Text = PlayerTwoTimerValue
        End If
        If PlayerOneTimer.Enabled = True Or PlayerTwoTimer.Enabled = True Then
            If PlayerOneTimerValue <= 0 Or PlayerTwoTimerValue <= 0 Then
                PlayerOneTimer.Enabled = False
                PlayerTwoTimer.Enabled = False
                CurrentGame.GameOver("TimeOut")
            End If
        End If
    End Sub
    Private Sub CountupTimer_Tick(sender As Object, e As EventArgs)
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
    Public Sub Board_Update() 'updates the board (front end) from the 2d array (back end)
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
    Public Sub UpdateWinningBoard(WinningColourChar) 'updates the board so it highlights the winning combination
        Dim Returns
        Dim x As Integer
        Dim y As Integer
        Dim WinType As String
        Dim PictureName As String 'the name of the file used
        Dim PictureNickName As String 'the nickname given to the file to be used in recording the winning state
        For x = 0 To 6 'changes all tokens to transparent
            For y = 0 To 5
                If CurrentGame.GetFilledStatus(x, y) = "Y" Then
                    BoardLocations(x, y).Image = Image.FromFile("YellowTokenTransparent.png")
                    BoardLocations(x, y).Name = "Y"
                ElseIf CurrentGame.GetFilledStatus(x, y) = "R" Then
                    BoardLocations(x, y).Image = Image.FromFile("RedTokenTransparent.png")
                    BoardLocations(x, y).Name = "R"
                Else
                    BoardLocations(x, y).Image = Nothing 'if an image is there that shouldn't be there it is cleared (such as when the game is reset)
                End If
            Next
        Next

        'Checks which win occurred
        If CurrentGame.CheckVertical(WinningColourChar).item1 <> "1000" Then ' Verticals 
            Returns = CurrentGame.CheckVertical(WinningColourChar)
            WinType = "Vertical"
        ElseIf CurrentGame.CheckHorizontal(WinningColourChar).item1 <> "1000" Then  ' Horizontals
            Returns = CurrentGame.CheckHorizontal(WinningColourChar)
            WinType = "Horizontal"
        ElseIf CurrentGame.CheckDiagonalOne(WinningColourChar).item1 <> "1000" Then ' Diagonals
            Returns = CurrentGame.CheckDiagonalOne(WinningColourChar)
            WinType = "DiagonalOne"
        ElseIf CurrentGame.CheckDiagonalTwo(WinningColourChar).item1 <> "1000" Then
            Returns = CurrentGame.CheckDiagonalTwo(WinningColourChar)
            WinType = "DiagonalTwo"
        Else
            WinType = "Draw"
        End If
        If WinType <> "Draw" Then 'only updates the board if there was a win
            x = Returns.item1
            y = Returns.item2

            If CurrentGame.GetFilledStatus(x, y) = "Y" Then
                PictureNickName = "WY" 'WY = Winning Yellow
                PictureName = "YellowToken.png"
            ElseIf CurrentGame.GetFilledStatus(x, y) = "R" Then
                PictureNickName = "WR"
                PictureName = "RedToken.png"
            End If

            If WinType = "Vertical" Then
                For i = 0 To 3
                    BoardLocations(x, y + i).Image = Image.FromFile(PictureName)
                    BoardLocations(x, y + i).Name = PictureNickName
                Next
            ElseIf WinType = "Horizontal" Then
                For i = 0 To 3
                    BoardLocations(x + i, y).Image = Image.FromFile(PictureName)
                    BoardLocations(x + i, y).Name = PictureNickName
                Next
            ElseIf WinType = "DiagonalOne" Then
                For i = 0 To 3
                    BoardLocations(x + i, y + i).Image = Image.FromFile(PictureName)
                    BoardLocations(x + i, y + i).Name = PictureNickName
                Next
            ElseIf WinType = "DiagonalTwo" Then
                For i = 0 To 3
                    BoardLocations(x + i, y - i).Image = Image.FromFile(PictureName)
                    BoardLocations(x + i, y - i).Name = PictureNickName
                Next
            End If
        End If

        ' Records winning game state for future access
        Using writer As New StreamWriter("PreviousGameStates.txt", True) 'open in append. records the game state.
            writer.WriteLine()
            For x = 0 To 6 'resets grid
                For y = 0 To 5
                    If BoardLocations(x, y).Image Is Nothing Then 'blank space
                        writer.Write("0" & ",")
                    ElseIf BoardLocations(x, y).Name = "R" Then 'red token
                        writer.Write("R" & ",")
                    ElseIf BoardLocations(x, y).Name = "Y" Then 'yellow token
                        writer.Write("Y" & ",")
                    ElseIf BoardLocations(x, y).Name = "WR" Then 'winning red
                        writer.Write("WR" & ",")
                    ElseIf BoardLocations(x, y).Name = "WY" Then 'winning yellow
                        writer.Write("WY" & ",")
                    End If
                Next
            Next
        End Using
    End Sub
    Public Sub Board_Refresh()
        For x = 0 To 6
            For Y = 0 To 5
                BoardLocations(x, Y).Refresh()
            Next
        Next
    End Sub
    Private Sub Board_Click(sender As Object, e As EventArgs) 'when a square is clicked on
        If CurrentGame.makemove(sender.tag) Then
            CurrentGame.SwitchMove()
        End If
    End Sub
    Private Sub Board_Hover(sender As Object, e As EventArgs) 'when a square is hovered over
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
    Public Function GetCurrentPlayer() As String
        Dim output As String
        output = Replace(CStr(CurrentPlayerLabel.Text), "Current Player:", "") 'replace command is used to clean up the output
        output = Replace(output, vbCrLf, "")
        Return output
    End Function
    Public Sub UpdateCurrentMoveVisual()
        CurrentMoveLabel.Text = "Current Move:" & vbNewLine & CurrentGame.getCurrentMove
    End Sub
    Public Function GetStartingPlayer() As String
        Return StartingPlayer
    End Function
    Public Function GetGameType() As String
        Return CurrentGameType
    End Function
    Public Function GetSubType()
        Return CurrentSubType
    End Function

    Public Function GetBoardEnabled() As Boolean
        Return BoardEnabled
    End Function
    Private Sub Game_Form_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing 'when x is pressed
        Main_Menu.Close() 'closing first form closes whole program
    End Sub
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
    Public Function GetFilledStatus(x, y) As String ' return is empty if the space is empty
        If BoardState(x, y) = "Y" Then
            Return "Y"
        ElseIf BoardState(x, y) = "R" Then
            Return "R"
        End If
    End Function
    Public Sub SetPlayerOneColour(Colour)
        PlayerOneColour = Colour
    End Sub
    Public Overridable Function CheckForWinningState(BoardState, CurrentPlayer) As Boolean 'searches 2D array of char (boardstate) for 4 in a row
        Dim WinDiscovered As Boolean = False
        Dim CurrentPlayerChar As Char = GetColourChar(CurrentPlayer)
        If CheckVertical(CurrentPlayerChar).item1 <> "1000" Then ' Verticals
            WinDiscovered = True
        ElseIf CheckHorizontal(CurrentPlayerChar).item1 <> "1000" Then  ' Horizontals
            WinDiscovered = True
        ElseIf CheckDiagonalOne(CurrentPlayerChar).item1 <> "1000" Then ' Diagonals
            WinDiscovered = True
        ElseIf CheckDiagonalTwo(CurrentPlayerChar).item1 <> "1000" Then
            WinDiscovered = True
        End If
        Return WinDiscovered
    End Function
    Public Function CheckVertical(CurrentPlayerChar)
        For y = 0 To 2
            For x = 0 To 6
                If BoardState(x, y) = BoardState(x, y + 1) And BoardState(x, y + 1) = BoardState(x, y + 2) And BoardState(x, y + 2) = BoardState(x, y + 3) And BoardState(x, y) = CurrentPlayerChar Then 'cant be blank
                    'if a sequence of 4 is detected, it must be a win for the current player, else it would have been detected by a previous check.
                    Return (x, y)
                End If
            Next
        Next
        Return ("1000", "") 'returns 1000 (an impossible position) if the win is not detected
    End Function
    Public Function CheckHorizontal(CurrentPlayerChar)
        For x = 0 To 3
            For y = 0 To 5
                If BoardState(x, y) = BoardState(x + 1, y) And BoardState(x + 1, y) = BoardState(x + 2, y) And BoardState(x + 2, y) = BoardState(x + 3, y) And BoardState(x, y) = CurrentPlayerChar Then 'if all 4 are equal
                    Return (x, y)
                End If
            Next
        Next
        Return ("1000", "")
    End Function
    Public Function CheckDiagonalOne(CurrentPlayerChar)
        For x = 0 To 3
            For y = 0 To 2
                If BoardState(x, y) = BoardState(x + 1, y + 1) And BoardState(x + 1, y + 1) = BoardState(x + 2, y + 2) And BoardState(x + 2, y + 2) = BoardState(x + 3, y + 3) And BoardState(x, y) = CurrentPlayerChar Then
                    Return (x, y)
                End If
            Next
        Next
        Return ("1000", "")
    End Function
    Public Function CheckDiagonalTwo(CurrentPlayerChar)
        For x = 0 To 3
            For y = 3 To 5
                If BoardState(x, y) = BoardState(x + 1, y - 1) And BoardState(x + 1, y - 1) = BoardState(x + 2, y - 2) And BoardState(x + 2, y - 2) = BoardState(x + 3, y - 3) And BoardState(x, y) = CurrentPlayerChar Then
                    Return (x, y)
                End If
            Next
        Next
        Return ("1000", "")
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
        Dim Winner As String
        Game_Form.ToggleBoardInteractivity()
        Game_Form.Board_Update()
        Game_Form.TimersStop()
        Game_Form.UpdateWinningBoard(GetColourChar(GetCurrentColour()))
        If WinMode.ToLower = "timeout" Then
            If Not CheckForDrawingState(BoardState) Then
                If Game_Form.GetCurrentPlayer() = PlayerOneName Then
                    MsgBox(Game_Form.GetCurrentPlayer() & " ran out of time. " & vbNewLine & PlayerTwoName & " has won the game.", MsgBoxStyle.DefaultButton1, "Game Ended")
                    Winner = PlayerTwoName
                Else  ' opposite player wins
                    MsgBox(Game_Form.GetCurrentPlayer() & " ran out of time. " & vbNewLine & PlayerOneName & " has won the game.", MsgBoxStyle.DefaultButton1, "Game Ended")
                    Winner = PlayerOneName
                End If
            Else

            End If
        ElseIf CheckForDrawingState(BoardState) Then 'game is drawn
            GameOverIndicator = True
            If WinMode = "Timed Game" Then 'If it is a timed game, the player with the lowest time wins
                If Game_Form.GetPlayerOneTimerValue > Game_Form.GetPlayerTwoTimerValue Then
                    MsgBox("The game was a draw. " & vbNewLine & PlayerTwoName & " completed their moves quicker, so wins.", MsgBoxStyle.DefaultButton1, "Game Ended")
                    Winner = PlayerOneName
                ElseIf Game_Form.GetPlayerOneTimerValue < Game_Form.GetPlayerTwoTimerValue Then
                    MsgBox("The game was a draw." & vbNewLine & PlayerOneName & " completed their moves quicker, so wins.", MsgBoxStyle.DefaultButton1, "Game Ended")
                    Winner = PlayerOneName
                Else 'if users had same time left (very unlikely)
                    MsgBox("The game was a draw.", MsgBoxStyle.DefaultButton1, "Game Ended")
                    Winner = "Draw"
                End If
            ElseIf WinMode = "Countdown" Then
                If Game_Form.GetPlayerOneTimerValue < Game_Form.GetPlayerTwoTimerValue Then
                    MsgBox("The game was a draw. " & vbNewLine & PlayerTwoName & " had more time left, so wins.", MsgBoxStyle.DefaultButton1, "Game Ended")
                    Winner = PlayerOneName
                ElseIf Game_Form.GetPlayerOneTimerValue > Game_Form.GetPlayerTwoTimerValue Then
                    MsgBox("The game was a draw." & vbNewLine & PlayerOneName & " had more time left, so wins.", MsgBoxStyle.DefaultButton1, "Game Ended")
                    Winner = PlayerOneName
                Else 'if users had same time left (very unlikely)
                    MsgBox("The game was a draw.", MsgBoxStyle.DefaultButton1, "Game Ended")
                    Winner = "Draw"
                End If
            Else
                MsgBox("The game was a draw.", MsgBoxStyle.DefaultButton1, "Game Ended")
                Winner = "Draw"
            End If
        Else 'Game has been won by a 4 in a row
            GameOverIndicator = True
            MsgBox(Game_Form.GetCurrentPlayer() & vbNewLine & "has won the game.", MsgBoxStyle.DefaultButton1, "Game Ended")
            Winner = Game_Form.GetCurrentPlayer()
        End If
        RecordGame(WinMode, Winner)
        Game_Form.Hide() 'hides game form
        Main_Menu.Show() 'returns to main main
    End Sub
    Protected Sub RecordGame(WinMode, WinningPlayer) ' stores game's data into file
        Dim PlayerOneColourString As String
        Dim ThisGameType As String
        Dim PlayerOneTime As String = Game_Form.GetPlayerOneTimerValue()
        Dim PlayerTwoTime As String = Game_Form.GetPlayerTwoTimerValue()
        WinningPlayer = WinningPlayer.Replace(vbCr, "").Replace(vbLf, "") 'removes new line from winningplayer
        If PlayerOneColour = Colours.Red Then 'converts colour to human readable form
            PlayerOneColourString = "Red"
        Else
            PlayerOneColourString = "Yellow"
        End If
        If Game_Form.GetGameType() = "Player" Then
            ThisGameType = "Player vs Player"
        Else
            ThisGameType = "Player vs Computer"
        End If
        Using writer As New StreamWriter("PreviousGames.txt", True) 'open in append. records data about the game.
            writer.WriteLine(DateTime.Now.ToString & "," & ThisGameType & "," & Game_Form.GetSubType & "," & PlayerOneName & "," & PlayerTwoName & "," & PlayerOneColourString & "," & WinningPlayer & "," & PlayerOneTime & "," & PlayerTwoTime)
        End Using
    End Sub
    Public Function GetGameOverIndicator() As Boolean
        Return GameOverIndicator
    End Function
    Protected Sub MakeAnimatedMove(XCoordinate, YCoordinate, ColourChar)
        Game_Form.ToggleBoardInteractivity() 'temporarily disables board
        For Y = 0 To YCoordinate
            If Y > 0 Then
                BoardState(XCoordinate, Y - 1) = Nothing
            End If
            BoardState(XCoordinate, Y) = ColourChar
            Game_Form.Board_Update()
            Game_Form.Board_Refresh()
            Threading.Thread.CurrentThread.Sleep(40) 'pause for 0.4s
        Next
        If Game_Form.GetCurrentPlayer() = PlayerOneName Then 'restarts timers
            Game_Form.SetPlayerOneTimerStatus("Enable")
        Else 'if player two's move
            Game_Form.SetPlayerTwoTimerStatus("Enable")
        End If
        Game_Form.ToggleBoardInteractivity() 're-enables board
    End Sub
    Public Function MakeMove(Location) As Boolean 'makes the move
        Dim Column As Integer = (CInt(Location)) \ 6
        Dim MoveMade As Boolean = False
        For y = 5 To 0 Step -1 'counts backwards (up)
            If GetFilledStatus(Column, y) <> "Y" And GetFilledStatus(Column, y) <> "R" And MoveMade = False Then
                MakeAnimatedMove(Column, y, GetColourChar(GetCurrentColour()))
                MoveMade = True
            End If
        Next
        Return MoveMade
    End Function
    Public Overridable Sub SwitchMove()
        UpdateCurrentMove()
        Game_Form.Board_Update()
        Game_Form.UpdateCurrentMoveVisual()
        If CheckForWinningState(BoardState, CurrentColour) Or CheckForDrawingState(BoardState) Then 'game over
            GameOver(Game_Form.GetSubType())
        Else 'switches colour
            If CurrentColour = Colours.Red Then
                CurrentColour = Colours.Yellow
            Else
                CurrentColour = Colours.Red
            End If
            Game_Form.UpdateCurrentPlayerVisual()
        End If
    End Sub
    Public Function GetCurrentMove() As Integer
        Return CurrentMove
    End Function
    Public Sub UpdateCurrentMove()
        CurrentMove += 1
    End Sub
    Public Function GetPlayerOneName() As String
        Return PlayerOneName
    End Function
    Public Function GetPlayerTwoName() As String
        Return PlayerTwoName
    End Function
    Public Function GetPlayerOneColour() As Colours
        Return PlayerOneColour
    End Function
    Public Overridable Sub SetPlayerTwoColour(Colour)
        PlayerTwoColour = Colour
    End Sub
    Public Sub SetCurrentColour(CurrentColourInput)
        CurrentColour = CurrentColourInput
    End Sub
    Public Function GetCurrentColour() As Colours
        Return CurrentColour
    End Function
    Public Function GetColourChar(ColourChoice) As String
        If ColourChoice = Colours.Red Then
            Return "R"
        Else
            Return "Y"
        End If
    End Function
    Public Overridable Function GetPlayerTwoColour() As Colours
        Return PlayerTwoColour
    End Function
End Class

Public Class Player_Vs_Computer_Game
    Inherits Game ' subclass of game
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
            Randomize()
            If (CInt(Math.Floor(Rnd() * 2))) = 0 Then 'random number, either 1 or 0 resulting in 50% chance to be either player
                SetCurrentColour(PlayerOneColour)
            Else
                SetCurrentColour(ComputerColour)
            End If
        End If
    End Sub
    Private Sub SetDepth(SpecifiedDifficulty) 'sets attributes based on difficulty
        If SpecifiedDifficulty = "Easy Difficulty" Then
            ChanceOfMistake = 50
            InitialDepth = 1
        ElseIf SpecifiedDifficulty = "Medium Difficulty" Then
            ChanceOfMistake = 25
            InitialDepth = 2
        ElseIf SpecifiedDifficulty = "Hard Difficulty" Then
            ChanceOfMistake = 10
            InitialDepth = 3
        Else 'Impossible difficulty
            ChanceOfMistake = 0
            InitialDepth = 5
        End If
    End Sub
    Private Sub ComputerMove()
        Dim Returns
        Dim BestXCoordinate As Integer
        Dim RandomMoveMade As Boolean
        Randomize()
        If ChanceOfMistake <= CInt((100) * Rnd()) Then 'random number from 0 to 100. If it is smaller than the chance of mistake then it makes a mistake and picks a random spot.
            Do
                Returns = Minimax(BoardState, InitialDepth, True)
                BestXCoordinate = Returns.item2
            Loop Until MakeComputerMove(BestXCoordinate) = True 'only allows move if column is not full
        Else
            ' random spot
            Do
                RandomMoveMade = True
                BestXCoordinate = CInt((6) * Rnd())
                If MakeComputerMove(BestXCoordinate) = False Then RandomMoveMade = False 'column was full, so tries again
            Loop Until RandomMoveMade
        End If
        SwitchMove()
    End Sub
    Private Function MakeComputerMove(XCoordinate) As Boolean
        Dim MoveMade As Boolean = False
        For y = 5 To 0 Step -1 'counts backwards (up) as columns fill bottom up
            If GetFilledStatus(XCoordinate, y) <> "Y" And GetFilledStatus(XCoordinate, y) <> "R" And MoveMade = False Then
                MakeAnimatedMove(XCoordinate, y, GetColourChar(ComputerColour))
                MoveMade = True
            End If
        Next
        Return MoveMade
    End Function
    Private Function Minimax(ByVal CurrentBoardState(,) As Char, Depth As Integer, MaximisingPlayer As Boolean)
        Dim Evaluation
        Dim BestEvaluation As Double
        Dim BestColumn As Integer
        Dim Searched As Boolean = False
        Dim YCoordinate As Integer
        If Depth = 0 Then  'if it has reached the end of its search depth
            Return (StaticEvaluation(CurrentBoardState), Nothing)
        Else
            If CheckForDrawingState(CurrentBoardState) Then 'if it is a draw (board is full)
                Return (0, Nothing)
            Else
                If CheckForWinningState(CurrentBoardState, GetPlayerTwoColour) Then 'computer got a 4 in a row
                    Return (10000000, Nothing)
                ElseIf CheckForWinningState(CurrentBoardState, GetPlayerOneColour) Then 'player got a 4 in a row
                    Return (-100000, Nothing)
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
                        Evaluation = Minimax(CurrentBoardState, Depth - 1, False) ' recursively calls itself, calls minimising
                        If (Evaluation.item1 + Depth) > BestEvaluation Then 'if it is less moves in, the score is higher. if two moves get the same score, the shortest win is favoured.
                            BestColumn = x
                            BestEvaluation = Evaluation.item1 + Depth
                        End If
                        CurrentBoardState(x, YCoordinate) = Nothing
                    End If
                    YCoordinate -= 1
                Loop Until YCoordinate = 0 Or Searched = True 'only searches one tile in the column (or 0 if full)
            Next
            Return (BestEvaluation, BestColumn)
        Else 'player move
            BestEvaluation = Double.PositiveInfinity
            For x = 0 To 6
                Searched = False
                YCoordinate = 5
                Do
                    If CurrentBoardState(x, YCoordinate) = Nothing And Searched = False Then
                        Searched = True ' so loop only runs once
                        CurrentBoardState(x, YCoordinate) = GetColourChar(GetPlayerOneColour())
                        Evaluation = Minimax(CurrentBoardState, Depth - 1, True) ' recursively calls itself, calls maximising
                        If (Evaluation.item1 - Depth) < BestEvaluation Then 'if it is less moves in, the score is higher. if two moves get the same score, the shortest win is favoured. 
                            BestColumn = x
                            BestEvaluation = Evaluation.item1 - Depth
                        End If
                        CurrentBoardState(x, YCoordinate) = Nothing
                    End If
                    YCoordinate -= 1
                Loop Until YCoordinate = 0 Or Searched = True 'only searches one tile in the column 
            Next
            Return (BestEvaluation, BestColumn)
        End If
    End Function
    Private Function StaticEvaluation(EvaluatingBoardState(,) As Char) As Decimal 'returns the score of that board state
        Dim EvaluatedScore As Decimal = 0
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
        EvaluatedScore += NextToCentreCount * 0.1
        EvaluatedScore += TwoFromCentreCount * 0.05

        ' Creates sections of 4 adjacent tiles that are analysed. All possible combinations are analysed.
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
    Private Function EvaluateScore(EvaluatingSection() As Char, AnalysisColour As Char) As Integer 'returns score for specific section of 4 adjacent tiles
        Dim NumberOfTilesInSequence = 0
        Dim NumberOfOppositeTilesInSequence = 0
        Dim SectionScore As Integer = 0
        For i = 0 To 3
            If EvaluatingSection(i) <> AnalysisColour And EvaluatingSection(i) <> Nothing Then 'if the sequence has the opponents colour anywhere in it, it is not viable
                NumberOfOppositeTilesInSequence += 1
                ' Return 0
            ElseIf EvaluatingSection(i) = AnalysisColour Then
                NumberOfTilesInSequence += 1
            End If
        Next
        If NumberOfTilesInSequence = 1 And NumberOfOppositeTilesInSequence = 0 Then '1/4 of the tiles is filled rest are empty
            SectionScore += 0
        ElseIf NumberOfTilesInSequence = 2 And NumberOfOppositeTilesInSequence = 0 Then '2/4 of the tiles is filled rest are empty
            SectionScore += 2
        ElseIf NumberOfTilesInSequence = 3 And NumberOfOppositeTilesInSequence = 0 Then '3/4 of the tiles is filled rest are empty
            SectionScore += 10
        ElseIf NumberOfTilesInSequence = 4 And NumberOfOppositeTilesInSequence = 0 Then '4/4 tiles are filled by one colour
            SectionScore += 100
        End If
        If NumberOfOppositeTilesInSequence = 3 And NumberOfTilesInSequence = 0 Then '3 of opposite player, 1 empty
            SectionScore -= 4
        End If
        Return SectionScore
    End Function
    Public Sub StartGame(subtype)
        If GetCurrentColour() = PlayerTwoColour Then 'Computer Move
            ComputerMove()
            If Game_Form.GetBoardEnabled = False Then 'if the board is off
                Game_Form.ToggleBoardInteractivity() 'activate buttons
            End If
        Else 'player move
            If Game_Form.GetBoardEnabled = False Then 'if the board is off
                Game_Form.ToggleBoardInteractivity() 'activate buttons
            End If
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
    Public Overrides Function GetPlayerTwoColour() As Colours
        MyBase.GetPlayerTwoColour()
        Return ComputerColour
    End Function
End Class
Class Player_Vs_Player_Game
    Inherits Game 'sub class
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
    Public Sub StartGame(SubType)
        If SubType = "Countdown" Or SubType = "Timed Game" Then
            Game_Form.MakeTimersVisible() ' start starting players timer
            If Game_Form.GetStartingPlayer() = PlayerOneName Then
                Game_Form.SetPlayerOneTimerStatus("Enable")
            Else
                Game_Form.SetPlayerTwoTimerStatus("Enable")
            End If
            'End If
        End If
        Game_Form.ToggleBoardInteractivity() 'activate buttons
    End Sub
    Public Overrides Sub SwitchMove()
        MyBase.SwitchMove()
        If Game_Form.GetPlayerOneTimerStatus() = "Enabled" Then 'if player one timer is counting up
            Game_Form.SetPlayerOneTimerStatus("Disable")
            Game_Form.SetPlayerTwoTimerStatus("Enable")
        Else 'if player two timer is counting up
            Game_Form.SetPlayerOneTimerStatus("Enable")
            Game_Form.SetPlayerTwoTimerStatus("Disable")
        End If
    End Sub
End Class