Public Class TicTacToe
    Private m_iGrid(2, 2) As GridEntry   'A 3x3 Grid
    Private m_bHasWinOccured As Boolean
    Private m_iWinner As GridEntry

    Public Event TicTacToeWinOccured(ByVal aWinner As GridEntry)
    Public Event Cat()

    Public Sub New()
        Me.Clear()
    End Sub

    'Clear the board
    Public Sub Clear()
        Dim iRow As Integer, iCol As Integer
        For iRow = 0 To 2
            For iCol = 0 To 2
                m_iGrid(iRow, iCol) = GridEntry.NoEntry
            Next
        Next
        m_bHasWinOccured = False
        m_iWinner = GridEntry.NoEntry
    End Sub

    'Three possible entries in a grid square
    Public Enum GridEntry
        NoEntry = 0
        PlayerX = 1
        PlayerO = 2
    End Enum

    Public Property Square(ByVal iRow As Integer, ByVal iCol As Integer) As GridEntry
        Get
            Return m_iGrid(iRow, iCol)
        End Get
        Set(ByVal Value As GridEntry)
            If m_iGrid(iRow, iCol) = GridEntry.NoEntry And _
               m_bHasWinOccured = False Then
                m_iGrid(iRow, iCol) = Value
                m_iWinner = CheckForWin()
                If m_iWinner = GridEntry.NoEntry Then
                    If CheckForDraw() = True Then
                        RaiseEvent Cat()
                    End If
                Else
                    m_bHasWinOccured = True
                    RaiseEvent TicTacToeWinOccured(m_iWinner)
                End If
            End If
        End Set
    End Property

    Public Property Square(ByVal aSC As SquareCoordinate) As GridEntry
        Get
            Return m_iGrid(aSC.Row, aSC.Column)
        End Get
        Set(ByVal Value As GridEntry)
            Me.Square(aSC.Row, aSC.Column) = Value
        End Set
    End Property

    Private Function CheckForWin() As GridEntry
        Dim iTemp As GridEntry

        iTemp = CheckForHorizontalWin()
        If iTemp <> GridEntry.NoEntry Then Return iTemp

        iTemp = CheckForVerticalWin()
        If iTemp <> GridEntry.NoEntry Then Return iTemp

        Return CheckForDiagonalWin()
    End Function

    Private Function CheckForHorizontalWin() As GridEntry
        Dim iRow As Integer

        Dim iFirstSquarePlayer As GridEntry

        For iRow = 0 To 2
            iFirstSquarePlayer = m_iGrid(iRow, 0)
            If iFirstSquarePlayer = m_iGrid(iRow, 1) And _
               iFirstSquarePlayer = m_iGrid(iRow, 2) Then
                Return iFirstSquarePlayer
            End If
        Next

        Return GridEntry.NoEntry
    End Function

    Private Function CheckForVerticalWin() As GridEntry
        Dim iCol As Integer
        Dim iFirstSquarePlayer As GridEntry

        For iCol = 0 To 2
            iFirstSquarePlayer = m_iGrid(0, iCol)
            If iFirstSquarePlayer = m_iGrid(1, iCol) And _
               iFirstSquarePlayer = m_iGrid(2, iCol) Then
                Return iFirstSquarePlayer
            End If
        Next

        Return GridEntry.NoEntry
    End Function

    Private Function CheckForDiagonalWin() As GridEntry
        Dim iCol As Integer
        Dim iRow As Integer
        Dim iFirstSquarePlayer As GridEntry

        'Use the middle sqaure as a comparison
        iFirstSquarePlayer = m_iGrid(1, 1)

        '\
        If iFirstSquarePlayer = m_iGrid(0, 0) And _
           iFirstSquarePlayer = m_iGrid(2, 2) Then
            Return iFirstSquarePlayer
        End If

        If iFirstSquarePlayer = m_iGrid(0, 2) And _
           iFirstSquarePlayer = m_iGrid(2, 0) Then
            Return iFirstSquarePlayer
        End If
    End Function

    Private Function CheckForDraw() As Boolean
        Dim iRow As Integer, iCol As Integer
        For iRow = 0 To 2
            For iCol = 0 To 2
                If m_iGrid(iRow, iCol) = GridEntry.NoEntry Then Return False
            Next
        Next

        Return True
    End Function
End Class

Public Class SquareCoordinate
    Private m_iRow As Integer
    Private m_iCol As Integer

    Public Sub New()
        m_iRow = 0 : m_iCol = 0
    End Sub

    Public Sub New(ByVal iRow As Integer, ByVal iCol As Integer)
        m_iRow = iRow : m_iCol = iCol
    End Sub

    Public Property Row()
        Get
            Return m_iRow
        End Get
        Set(ByVal Value)
            m_iRow = Value
        End Set
    End Property

    Public Property Column()
        Get
            Return m_iCol
        End Get
        Set(ByVal Value)
            m_iCol = Value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return " [" & m_iRow & ", " & m_iCol & "] "
    End Function
End Class
