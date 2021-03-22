Imports System.IO
Public Class Previous_Games ' Shows results of previous 25 games loaded from a file 
    ' Control Declarations

    Private Title As New PictureBox
    Private CreditLabel As New Label
    Private PreviousGamesTable As New DataGridView
    Private BackButton As New Button
    Private DescriptionLabel As New Label

    ' Variable Declarations

    Private CloseButtonClicked As Boolean
    Private ListOfGames As New List(Of String)

    ' Code
    Public Sub This_Form_Setup() 'formats forms
        Me.Show()
        Processes.Form_Setup(Me)
        ' Form Title
        CloseButtonClicked = False
        Title.Location = New Point(242, 5)
        Title.Size = New Size(500, 50)
        Title.SizeMode = PictureBoxSizeMode.Zoom
        Title.Image = Image.FromFile("Title.png")
        Controls.Add(Title)
        Processes.RunTimeContructor(CreditLabel, 0, 748, 984, 13, "James Cracknell - 191673", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 8.25, Cursors.Default, Color.Transparent)
        Controls.Add(CreditLabel)
        Table_Setup()
    End Sub
    Private Sub Table_Setup() ' creates controls
        ' DescriptionLabel
        Processes.RunTimeContructor(DescriptionLabel, 0, 60, 984, 35, "The most recent 25 games are displayed below." & vbNewLine & "Double click on the game type for more information about the game.", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Default, Color.Transparent)
        Controls.Add(DescriptionLabel)
        ' BackButton
        Processes.RunTimeContructor(BackButton, 342, 680, 300, 50, "Back to Menu", ContentAlignment.MiddleCenter, "Microsoft Sans Serif", 9.75, Cursors.Hand, Color.Transparent)
        AddHandler BackButton.Click, AddressOf BackButton_Click
        Controls.Add(BackButton)
        CreateTable()
        AddData()
    End Sub
    Private Sub CreateTable() 'creates and formats the data table 
        With PreviousGamesTable
            .BackgroundColor = Color.LightBlue
            .BorderStyle = BorderStyle.None
            .ColumnCount = 7
            .ColumnHeadersDefaultCellStyle.Font = New Font(PreviousGamesTable.Font, FontStyle.Bold)
            .Location = New Point(150, 100)
            .Size = New Size(800, 600)
            .RowHeadersVisible = False 'removes row headers
            .GridColor = Color.Black
            ' Set column headers
            .Columns(0).Name = "Date and Time"
            .Columns(1).Name = "Game Type"
            .Columns(2).Name = "SubType"
            .Columns(3).Name = "Player 1 Name"
            .Columns(4).Name = "Player 2 Name"
            .Columns(5).Name = "Player 1 Colour"
            .Columns(6).Name = "Winner"
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.AllCells
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
        End With
        For i = 0 To 6
            PreviousGamesTable.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        With PreviousGamesTable 'Prevents the user from manually changing the table
            .ReadOnly = True
            .AllowDrop = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToOrderColumns = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
        End With
        Controls.Add(PreviousGamesTable)
    End Sub
    Private Sub AddData() 'adds data from file to table
        Dim newRow As String()
        Dim StartPosition As Integer
        Using reader As New StreamReader("PreviousGames.txt")
            Do
                ListOfGames.Add(reader.ReadLine)
            Loop Until reader.EndOfStream
        End Using

        If ListOfGames.Count < 25 Then 'if there are less than 25 games recorded, starts at 0
            StartPosition = 0
        Else
            StartPosition = ListOfGames.Count - 25 'if there are more than 25 games, starts at most recent - 25 (to display most recent 25 games)
        End If

        For i = StartPosition To ListOfGames.Count - 1
            newRow = Split(ListOfGames(i), ",")
            PreviousGamesTable.Rows.Add(newRow)
        Next
        AddHandler PreviousGamesTable.CellContentDoubleClick, AddressOf PreviousGamesTable_CellContentDoubleClick
    End Sub
    Private Sub PreviousGamesTable_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) 'if double click
        If ListOfGames.Count < 25 Then
            Previous_Game_States.This_Form_Setup(e.RowIndex) ' opens the previous game state for that data
        Else
            Previous_Game_States.This_Form_Setup(e.RowIndex + ListOfGames.Count - 25) ' if there is more than 25 items in the list, it needs to account for the data not shown
        End If
    End Sub
    Private Sub BackButton_Click(sender As Object, e As EventArgs) 'if back button is clicked
        CloseButtonClicked = True
        Me.Close()
        Main_Menu.Show()
    End Sub
    Private Sub Previous_Games_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing 'when x is pressed
        If CloseButtonClicked = False Then 'only close whole program when the 
            Main_Menu.Close() 'closing first form closes whole program
        End If
    End Sub

End Class
