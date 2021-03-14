Imports System.IO
Public Class Previous_Game_States ' Shows a specific previous game board, including any relevent info
    ' Control Declarations

    Private Title As New PictureBox
    Private BoardLocations(6, 5) As PictureBox
    Private CloseButton As New Button
    Private GameTypeTitleLabel As New Label
    Private GameTypeLabel As New Label
    Private SubTypeTitleLabel As New Label
    Private SubTypeLabel As New Label
    Private WinnerNameTitleLabel As New Label
    Private WinnerNameLabel As New Label
    Private LoserNameTitleLabel As New Label
    Private LoserNameLabel As New Label
    Private PlayerOneTimeTitleLabel As New Label
    Private PlayerOneTimeLabel As New Label
    Private PlayerTwoTimeTitleLabel As New Label
    Private PlayerTwoTimeLabel As New Label

    ' Variable Declarations

    Protected OneDBoardState(42) As String
    Protected BoardState(6, 5) As String
    Private Painted As Boolean = False
    ' Main Code
    Public Sub This_Form_Setup(DesiredGame) 'formats form
        Me.Show()
        Painted = False 'so the board is reprinted at the start
        Processes.Form_Setup(Me)
        Me.Height = 900
        ' Form Title
        Title.Location = New Point(242, 5)
        Title.Size = New Size(500, 50)
        Title.SizeMode = PictureBoxSizeMode.Zoom
        Title.Image = Image.FromFile("Title.png")
        Controls.Add(Title)
        Board_Setup()
        ' CloseButton
        Processes.RunTimeContructor(CloseButton, 342, 780, 300, 50, "Close", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler CloseButton.Click, AddressOf CloseButton_Click
        Controls.Add(CloseButton)
        Load_Previous_Game_State(DesiredGame)
        UpdateBoard()
        SetUpExtraInfo(DesiredGame)
        AddHandler Me.Paint, AddressOf Board_Paint
    End Sub

    Private Sub Load_Previous_Game_State(DesiredGame) ' reads requested data from file 
        Dim LatestGameState As String
        Using reader As New StreamReader("PreviousGameStates.txt") 'saves requested row in file to variable
            For i = 0 To DesiredGame 'moves down file until user specified line
                LatestGameState = reader.ReadLine 'saves the row 
            Next
        End Using
        Convert1DArrayTo2DArray(LatestGameState)
    End Sub
    Private Function GetExtraInfo(DesiredGame)
        Dim ExtraInfo As String
        Dim ExtraInfoArray() As String
        Using reader As New StreamReader("PreviousGames.txt")
            For i = 0 To DesiredGame 'moves down file until user specified line
                ExtraInfo = reader.ReadLine
            Next
        End Using
        ExtraInfoArray = Split(ExtraInfo, ",")
        Return ExtraInfoArray
    End Function

    Private Sub SetUpExtraInfo(DesiredGame) ' creates controls to display data
        ' Close Button
        Processes.RunTimeContructor(CloseButton, 342, 780, 300, 50, "Close", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        Dim ExtraInfo As String()
        ExtraInfo = GetExtraInfo(DesiredGame)
        ' GameTypeTitleLabel
        Processes.RunTimeContructor(GameTypeTitleLabel, 150, 95, 125, 20, "Game Type:", ContentAlignment.BottomLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        GameTypeTitleLabel.Font = New Font(GameTypeTitleLabel.Font, FontStyle.Bold)
        Controls.Add(GameTypeTitleLabel)
        ' GameTypeLabel
        Processes.RunTimeContructor(GameTypeLabel, 150, 100, 125, 50, ExtraInfo(1), ContentAlignment.MiddleLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        Controls.Add(GameTypeLabel)
        ' SubTypeTitleLabel
        Processes.RunTimeContructor(SubTypeTitleLabel, 275, 95, 130, 20, "Subtype:", ContentAlignment.BottomLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        SubTypeTitleLabel.Font = New Font(SubTypeTitleLabel.Font, FontStyle.Bold)
        Controls.Add(SubTypeTitleLabel)
        ' SubTypeLabel
        Processes.RunTimeContructor(SubTypeLabel, 275, 100, 130, 50, ExtraInfo(2), ContentAlignment.MiddleLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        Controls.Add(SubTypeLabel)
        ' WinnerNameTitleLabel
        Processes.RunTimeContructor(WinnerNameTitleLabel, 405, 95, 75, 20, "Winner:", ContentAlignment.BottomLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        WinnerNameTitleLabel.Font = New Font(WinnerNameTitleLabel.Font, FontStyle.Bold)
        Controls.Add(WinnerNameTitleLabel)
        ' WinnerNameLabel
        Processes.RunTimeContructor(WinnerNameLabel, 405, 100, 115, 50, ExtraInfo(6), ContentAlignment.MiddleLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        Controls.Add(WinnerNameLabel)
        ' LoserNameTitleLabel
        Processes.RunTimeContructor(LoserNameTitleLabel, 520, 95, 115, 20, "Loser:", ContentAlignment.BottomLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        LoserNameTitleLabel.Font = New Font(LoserNameTitleLabel.Font, FontStyle.Bold)
        Controls.Add(LoserNameTitleLabel)
        ' LoserNameLabel
        If ExtraInfo(6) = ExtraInfo(3) Then 'winner is player1
            Processes.RunTimeContructor(LoserNameLabel, 520, 100, 115, 50, ExtraInfo(4), ContentAlignment.MiddleLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        Else
            Processes.RunTimeContructor(LoserNameLabel, 520, 100, 115, 50, ExtraInfo(3), ContentAlignment.MiddleLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        End If
        Controls.Add(LoserNameLabel)
        If ExtraInfo(2) = "Countdown" Or ExtraInfo(2) = "Timed Game" Then 'if the game had a timer
            'PlayerOneTimeTitleLabel
            Processes.RunTimeContructor(PlayerOneTimeTitleLabel, 635, 95, 175, 20, "Player One Time:", ContentAlignment.BottomLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
            PlayerOneTimeTitleLabel.Font = New Font(PlayerOneTimeTitleLabel.Font, FontStyle.Bold)
            Controls.Add(PlayerOneTimeTitleLabel)
            ' PlayerOneTimeLabel
            Processes.RunTimeContructor(PlayerOneTimeLabel, 635, 100, 175, 50, ExtraInfo(7), ContentAlignment.MiddleLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
            Controls.Add(PlayerOneTimeLabel)
            'PlayerTwoTimeTitleLabel
            Processes.RunTimeContructor(PlayerTwoTimeTitleLabel, 810, 95, 175, 20, "Player Two Time:", ContentAlignment.BottomLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
            PlayerTwoTimeTitleLabel.Font = New Font(PlayerTwoTimeTitleLabel.Font, FontStyle.Bold)
            Controls.Add(PlayerTwoTimeTitleLabel)
            ' PlayerTwoTimeLabel
            Processes.RunTimeContructor(PlayerTwoTimeLabel, 810, 100, 175, 50, ExtraInfo(8), ContentAlignment.MiddleLeft, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
            Controls.Add(PlayerTwoTimeLabel)
            If ExtraInfo(2) = "Countdown" Then 'if it is a countdown game, the title is changed
                PlayerOneTimeTitleLabel.Text = "Player One Time Left:"
                PlayerTwoTimeTitleLabel.Text = "Player Two Time Left:"
                PlayerOneTimeTitleLabel.Refresh()
                PlayerTwoTimeTitleLabel.Refresh()
            End If
        End If

    End Sub
    Private Sub Convert1DArrayTo2DArray(LatestGameState) 'Converts grid to 2D array
        OneDBoardState = Split(LatestGameState, ",")
        For x = 0 To 6
            For y = 0 To 5
                If OneDBoardState(x * 6 + y) = "0" Then
                    BoardState(x, y) = "0"
                ElseIf OneDBoardState(x * 6 + y) = "R" Then
                    BoardState(x, y) = "R"
                ElseIf OneDBoardState(x * 6 + y) = "Y" Then
                    BoardState(x, y) = "Y"
                ElseIf OneDBoardState(x * 6 + y) = "WR" Then
                    BoardState(x, y) = "WR"
                ElseIf OneDBoardState(x * 6 + y) = "WY" Then
                    BoardState(x, y) = "WY"
                End If
            Next
        Next
    End Sub
    Private Sub Board_Setup() ' creates the board picturebox
        Dim XPos, YPos As Integer
        ' Board
        XPos = 146
        For x = 0 To 6
            YPos = 170
            For y = 0 To 5
                BoardLocations(x, y) = New PictureBox
                BoardLocations(x, y) = New PictureBox
                BoardLocations(x, y).Location = New Point(XPos, YPos)
                BoardLocations(x, y).Size = New Size(98, 98)
                BoardLocations(x, y).SizeMode = PictureBoxSizeMode.Zoom ' Set to fully display image without warping it
                BoardLocations(x, y).BackColor = Color.Blue 'since I cannot make it transparent, I made the background blue to match the box
                BoardLocations(x, y).Enabled = False
                BoardLocations(x, y).Tag = (x * 6 + y)
                Controls.Add(BoardLocations(x, y))
                YPos += 100
            Next
            XPos += 100
        Next
    End Sub

    Private Sub Board_Paint(sender As Object, e As PaintEventArgs) ' paints the board
        If Painted = False Then 'prevents the board being redrawn. It should only be drawn once.
            Painted = True
            Dim g As Graphics
            g = Me.CreateGraphics

            ' Rectangle Background
            Dim myBrush As New SolidBrush(Color.Blue) 'creates a new blue brush 
            Dim formGraphics As Graphics
            formGraphics = Me.CreateGraphics()
            formGraphics.FillRectangle(myBrush, New Rectangle(145, 170, 700, 600))
            myBrush.Dispose()
            formGraphics.Dispose()
            Dim BlackPen As New Pen(Brushes.Black)
            BlackPen.Width = 6.0F
            ' Rectangle Outline
            g.DrawRectangle(BlackPen, 145, 170, 700, 600) '(pentype, startingx, startingy, xsize, ysize)

            'vertical lines
            Dim x1 As Integer = 245
            For i = 0 To 6
                g.DrawLine(BlackPen, x1, 170, x1, 770) '(pentype, startingx, startingy, endingx, endingy)
                x1 += 100
            Next

            'horizontal lines
            Dim y1 As Integer = 270
            For j = 0 To 5
                g.DrawLine(BlackPen, 145, y1, 845, y1)
                y1 += 100
            Next
        End If
    End Sub

    Private Sub UpdateBoard() ' Changes the board images to that of the requested state
        For x = 0 To 6
            For y = 0 To 5
                If BoardState(x, y) = "Y" Then
                    BoardLocations(x, y).Image = Image.FromFile("YellowTokenTransparent.png")
                ElseIf BoardState(x, y) = "R" Then
                    BoardLocations(x, y).Image = Image.FromFile("RedTokenTransparent.png")
                ElseIf BoardState(x, y) = "WY" Then
                    BoardLocations(x, y).Image = Image.FromFile("YellowToken.png")
                ElseIf BoardState(x, y) = "WR" Then
                    BoardLocations(x, y).Image = Image.FromFile("RedToken.png")
                Else
                    BoardLocations(x, y).Image = Nothing 'if an image is there that shouldn't be there it is cleared (such as when the game is reset)
                End If
            Next
        Next
    End Sub


    Private Sub CloseButton_Click(sender As Object, e As EventArgs) 'if CloseButton is clicked
        Me.Close()
    End Sub
End Class