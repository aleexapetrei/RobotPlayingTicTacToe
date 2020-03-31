Public Class frmTicTacToe
    Inherits System.Windows.Forms.Form
    Private m_TicTacToeGame As New TicTacToe
    Private m_iCurrentPlayer As TicTacToe.GridEntry = TicTacToe.GridEntry.PlayerX
    Private m_TicTacToeAI As New TicTacToeAI(TicTacToe.GridEntry.PlayerO)
    Private m_bIsAIEnabled As Boolean = True
    Private m_bGameEnded As Boolean = False

    Private m_iPlayerXCount As Integer
    Private m_iPlayerOCount As Integer
    Private m_iDrawCount As Integer

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'Grid Clicks
        AddHandler lblGrid00.Click, AddressOf lblGrid_Click
        AddHandler lblGrid01.Click, AddressOf lblGrid_Click
        AddHandler lblGrid02.Click, AddressOf lblGrid_Click
        AddHandler lblGrid10.Click, AddressOf lblGrid_Click
        AddHandler lblGrid11.Click, AddressOf lblGrid_Click
        AddHandler lblGrid12.Click, AddressOf lblGrid_Click
        AddHandler lblGrid20.Click, AddressOf lblGrid_Click
        AddHandler lblGrid21.Click, AddressOf lblGrid_Click
        AddHandler lblGrid22.Click, AddressOf lblGrid_Click

        Me.DrawGrid()

        AddHandler m_TicTacToeGame.TicTacToeWinOccured, AddressOf WeHaveAWinner
        AddHandler m_TicTacToeGame.Cat, AddressOf Cat

        MenuItem2.Checked = m_bIsAIEnabled

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblGrid22 As System.Windows.Forms.Label
    Friend WithEvents lblGrid21 As System.Windows.Forms.Label
    Friend WithEvents lblGrid20 As System.Windows.Forms.Label
    Friend WithEvents lblGrid12 As System.Windows.Forms.Label
    Friend WithEvents lblGrid11 As System.Windows.Forms.Label
    Friend WithEvents lblGrid10 As System.Windows.Forms.Label
    Friend WithEvents lblGrid02 As System.Windows.Forms.Label
    Friend WithEvents lblGrid01 As System.Windows.Forms.Label
    Friend WithEvents lblGrid00 As System.Windows.Forms.Label
    Friend WithEvents lblCurrentPlayer As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblYourWins As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblMyWins As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblDraws As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblGrid22 = New System.Windows.Forms.Label
        Me.lblGrid21 = New System.Windows.Forms.Label
        Me.lblGrid20 = New System.Windows.Forms.Label
        Me.lblGrid12 = New System.Windows.Forms.Label
        Me.lblGrid11 = New System.Windows.Forms.Label
        Me.lblGrid10 = New System.Windows.Forms.Label
        Me.lblGrid02 = New System.Windows.Forms.Label
        Me.lblGrid01 = New System.Windows.Forms.Label
        Me.lblGrid00 = New System.Windows.Forms.Label
        Me.lblCurrentPlayer = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblYourWins = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblMyWins = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblDraws = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblGrid22)
        Me.GroupBox1.Controls.Add(Me.lblGrid21)
        Me.GroupBox1.Controls.Add(Me.lblGrid20)
        Me.GroupBox1.Controls.Add(Me.lblGrid12)
        Me.GroupBox1.Controls.Add(Me.lblGrid11)
        Me.GroupBox1.Controls.Add(Me.lblGrid10)
        Me.GroupBox1.Controls.Add(Me.lblGrid02)
        Me.GroupBox1.Controls.Add(Me.lblGrid01)
        Me.GroupBox1.Controls.Add(Me.lblGrid00)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 144)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Grid"
        '
        'lblGrid22
        '
        Me.lblGrid22.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid22.Location = New System.Drawing.Point(128, 104)
        Me.lblGrid22.Name = "lblGrid22"
        Me.lblGrid22.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid22.TabIndex = 17
        Me.lblGrid22.Text = "X"
        Me.lblGrid22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid21
        '
        Me.lblGrid21.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid21.Location = New System.Drawing.Point(88, 104)
        Me.lblGrid21.Name = "lblGrid21"
        Me.lblGrid21.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid21.TabIndex = 16
        Me.lblGrid21.Text = "X"
        Me.lblGrid21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid20
        '
        Me.lblGrid20.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid20.Location = New System.Drawing.Point(48, 104)
        Me.lblGrid20.Name = "lblGrid20"
        Me.lblGrid20.Size = New System.Drawing.Size(24, 25)
        Me.lblGrid20.TabIndex = 15
        Me.lblGrid20.Text = "X"
        Me.lblGrid20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid12
        '
        Me.lblGrid12.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid12.Location = New System.Drawing.Point(128, 64)
        Me.lblGrid12.Name = "lblGrid12"
        Me.lblGrid12.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid12.TabIndex = 14
        Me.lblGrid12.Text = "X"
        Me.lblGrid12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid11
        '
        Me.lblGrid11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid11.Location = New System.Drawing.Point(88, 64)
        Me.lblGrid11.Name = "lblGrid11"
        Me.lblGrid11.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid11.TabIndex = 13
        Me.lblGrid11.Text = "X"
        Me.lblGrid11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid10
        '
        Me.lblGrid10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid10.Location = New System.Drawing.Point(48, 64)
        Me.lblGrid10.Name = "lblGrid10"
        Me.lblGrid10.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid10.TabIndex = 12
        Me.lblGrid10.Text = "X"
        Me.lblGrid10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid02
        '
        Me.lblGrid02.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid02.Location = New System.Drawing.Point(128, 24)
        Me.lblGrid02.Name = "lblGrid02"
        Me.lblGrid02.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid02.TabIndex = 11
        Me.lblGrid02.Text = "X"
        Me.lblGrid02.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid01
        '
        Me.lblGrid01.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid01.Location = New System.Drawing.Point(88, 24)
        Me.lblGrid01.Name = "lblGrid01"
        Me.lblGrid01.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid01.TabIndex = 10
        Me.lblGrid01.Text = "X"
        Me.lblGrid01.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblGrid00
        '
        Me.lblGrid00.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblGrid00.Location = New System.Drawing.Point(48, 24)
        Me.lblGrid00.Name = "lblGrid00"
        Me.lblGrid00.Size = New System.Drawing.Size(24, 23)
        Me.lblGrid00.TabIndex = 9
        Me.lblGrid00.Text = "X"
        Me.lblGrid00.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblCurrentPlayer
        '
        Me.lblCurrentPlayer.Location = New System.Drawing.Point(16, 176)
        Me.lblCurrentPlayer.Name = "lblCurrentPlayer"
        Me.lblCurrentPlayer.Size = New System.Drawing.Size(200, 23)
        Me.lblCurrentPlayer.TabIndex = 10
        Me.lblCurrentPlayer.Text = "Label1"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(256, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 23)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Your Wins"
        '
        'lblYourWins
        '
        Me.lblYourWins.Location = New System.Drawing.Point(320, 40)
        Me.lblYourWins.Name = "lblYourWins"
        Me.lblYourWins.Size = New System.Drawing.Size(40, 23)
        Me.lblYourWins.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(256, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "My Wins"
        '
        'lblMyWins
        '
        Me.lblMyWins.Location = New System.Drawing.Point(320, 72)
        Me.lblMyWins.Name = "lblMyWins"
        Me.lblMyWins.Size = New System.Drawing.Size(40, 23)
        Me.lblMyWins.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(256, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 23)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Draws"
        '
        'lblDraws
        '
        Me.lblDraws.Location = New System.Drawing.Point(320, 104)
        Me.lblDraws.Name = "lblDraws"
        Me.lblDraws.Size = New System.Drawing.Size(40, 23)
        Me.lblDraws.TabIndex = 16
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2})
        Me.MenuItem1.Text = "&Settings"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "&AI"
        '
        'frmTicTacToe
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 213)
        Me.Controls.Add(Me.lblDraws)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblMyWins)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblYourWins)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblCurrentPlayer)
        Me.Controls.Add(Me.GroupBox1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmTicTacToe"
        Me.Text = "Tic Tac Toe"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Grid Code
    Public Sub lblGrid_Click(ByVal sender As Object, ByVal args As EventArgs)
        Try
            Dim aLabelControl As Label = CType(sender, Label)
            Dim sName As String = aLabelControl.Name
            Dim iRow As Integer = sName.Substring(sName.Length - 2, 1)
            Dim iCol As Integer = sName.Substring(sName.Length - 1, 1)
            If m_TicTacToeGame.Square(iRow, iCol) = TicTacToe.GridEntry.NoEntry Then
                m_TicTacToeGame.Square(iRow, iCol) = m_iCurrentPlayer
                ToggleCurrentPlayer()
                DrawGrid()
            End If
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub ToggleCurrentPlayer()
        If m_iCurrentPlayer = TicTacToe.GridEntry.PlayerX Then
            m_iCurrentPlayer = TicTacToe.GridEntry.PlayerO
            If m_bIsAIEnabled = True And m_bGameEnded = False Then
                PerformAIMove()
            End If
        Else
            m_iCurrentPlayer = TicTacToe.GridEntry.PlayerX
        End If
    End Sub

    Private Sub DrawGrid()
        Dim iRow As Integer
        Dim iCol As Integer
        Dim aLabelControl As Label

        For iRow = 0 To 2
            For iCol = 0 To 2
                Try
                    aLabelControl = GetLabelControl(iRow, iCol)
                    Select Case m_TicTacToeGame.Square(iRow, iCol)
                        Case TicTacToe.GridEntry.NoEntry
                            aLabelControl.Text = ""
                        Case TicTacToe.GridEntry.PlayerX
                            aLabelControl.Text = "X"
                        Case TicTacToe.GridEntry.PlayerO
                            aLabelControl.Text = "O"
                    End Select
                Catch ex As Exception

                End Try
            Next
        Next

        updatelabels()
    End Sub

    Private Sub PerformAIMove()
        Try
            Dim aSC As SquareCoordinate
            Try
                aSC = _
                m_TicTacToeAI.GetAIMove(m_TicTacToeGame)

            Catch ex As NullReferenceException
                Stop
            Catch ex As Exception
                Stop
            End Try
            m_TicTacToeGame.Square(aSC) = TicTacToe.GridEntry.PlayerO
            DrawGrid()
            ToggleCurrentPlayer()
        Catch ex As Exception
                Stop
            End Try

    End Sub

    Private Function GetLabelControl(ByVal iRow As Integer, ByVal iCol As Integer) As Label
        Dim x As Integer
        Dim aControl As Control
        For x = 0 To GroupBox1.Controls.Count - 1
            aControl = GroupBox1.Controls(x)
            If aControl.Name = "lblGrid" & iRow & iCol Then Return aControl
        Next
    End Function

    Private Sub Cat()
        DrawGrid()
        MessageBox.Show("A Cat Has Occured ... There is no winners.")
        m_iDrawCount = m_iDrawCount + 1
        m_bGameEnded = True
        NewGame()
    End Sub

    Private Sub WeHaveAWinner(ByVal iWinner As TicTacToe.GridEntry)
        DrawGrid()
        Dim sMessage As String = "Player "
        Select Case iWinner
            Case TicTacToe.GridEntry.PlayerX
                sMessage = sMessage & " X has won."
                m_iPlayerXCount = m_iPlayerXCount + 1
            Case TicTacToe.GridEntry.PlayerO
                sMessage = sMessage & " O has won."
                m_iPlayerOCount = m_iPlayerOCount + 1
        End Select
        MessageBox.Show(sMessage)
        m_bGameEnded = True
        NewGame()
    End Sub

    Private Sub NewGame()
        Dim rtnDlgResult As DialogResult
        rtnDlgResult = MessageBox.Show("Would you like to play another game?", "New Game", MessageBoxButtons.YesNo)
        If rtnDlgResult = DialogResult.Yes Then
            m_TicTacToeGame.Clear()
            m_bGameEnded = False
        Else
            Me.Close()
        End If
    End Sub

    Private Sub updatelabels()
        Select Case m_iCurrentPlayer
            Case TicTacToe.GridEntry.PlayerX
                lblCurrentPlayer.Text = "It is Player X's turn."
            Case TicTacToe.GridEntry.PlayerO
                lblCurrentPlayer.Text = "It is Player O's turn."
        End Select

        lblYourWins.Text = m_iPlayerXCount
        lblMyWins.Text = m_iPlayerOCount
        lblDraws.Text = m_iDrawCount
    End Sub

    'AI Setting
    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        If m_bIsAIEnabled = True Then
            m_bIsAIEnabled = False
            MenuItem2.Checked = False
            Me.ReSet()
        Else
            m_bIsAIEnabled = True
            Me.ReSet()
            MenuItem2.Checked = True
        End If
    End Sub

    Private Sub ReSet()
        m_TicTacToeGame.Clear()
        m_iPlayerXCount = 0
        m_iPlayerOCount = 0
        m_iDrawCount = 0
        DrawGrid()
    End Sub
End Class
