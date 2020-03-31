Public Class TicTacToeAI
    Private m_iAIPlayer As TicTacToe.GridEntry
    Private m_iHumanPlayer As TicTacToe.GridEntry
    Private m_CurrentBoard As TicTacToe


    Public Sub New(ByVal iAIPlayer As TicTacToe.GridEntry)
        m_iAIPlayer = iAIPlayer

        If iAIPlayer = TicTacToe.GridEntry.PlayerX Then
            m_iHumanPlayer = TicTacToe.GridEntry.PlayerO
        Else
            m_iHumanPlayer = TicTacToe.GridEntry.PlayerX
        End If
    End Sub

    Public Function GetAIMove(ByVal aTTT As TicTacToe) As SquareCoordinate
        Dim MoveList As New ArrayList
        m_CurrentBoard = aTTT

        Dim iRow As Integer
        For iRow = 0 To 2
            Dim possibleMoves() As SquareCoordinate
            Try
                possibleMoves = GetPossibleRowMoves(iRow)
            Catch ex As NullReferenceException
                Stop
            Catch ex As Exception
                Stop
            End Try

            If Not possibleMoves Is Nothing Then
                Dim x As Integer
                For x = 0 To possibleMoves.Length - 1
                    MoveList.Add(DetermineMoveEntry(possibleMoves(x)))
                Next
            End If
        Next

        'Return a blocking move or choose a random normal move
        Dim NormalMoveList As New ArrayList
        Dim aMoveEntry As MoveEntry

        'Make sure we check for winning moves
        For Each aMoveEntry In MoveList
            Select Case aMoveEntry.MoveType
                Case MoveType.Winning
                    Console.Write(ControlChars.CrLf & aMoveEntry.ToString)
                    Return aMoveEntry.SquareCoordinate
            End Select
        Next

        'Then blocking moves
        For Each aMoveEntry In MoveList
            Select Case aMoveEntry.MoveType
                Case MoveType.Blocking
                    Console.Write(ControlChars.CrLf & aMoveEntry.ToString)
                    Return aMoveEntry.SquareCoordinate
                Case Else
                    NormalMoveList.Add(aMoveEntry)
            End Select
        Next

        'Determine if there is a best normal move entry
        Try
            Dim bnME As SquareCoordinate = _
            DetermineBestNormalMove(NormalMoveList)
            Console.Write(ControlChars.CrLf & "Best -->" & bnME.ToString)
            Return bnME
        Catch ex As Exception

        End Try

        'If not choose one at random
        Dim iRandomItem As Integer = CInt(Int(((NormalMoveList.Count - 1) - 0 + 1) * Rnd() + 0))
        Dim nME As MoveEntry = NormalMoveList(iRandomItem)
        Console.Write(ControlChars.CrLf & nME.ToString)
        Return nME.SquareCoordinate
    End Function

    Private Function GetPossibleRowMoves(ByVal iRow As Integer) As SquareCoordinate()
        Dim PossibleMoves As New ArrayList
        Dim iCol As Integer
        For iCol = 0 To 2
            If m_CurrentBoard.Square(iRow, iCol) = TicTacToe.GridEntry.NoEntry Then
                PossibleMoves.Add(New SquareCoordinate(iRow, iCol))
            End If
        Next

        If PossibleMoves.Count > 0 Then
            Dim returnSC(PossibleMoves.Count - 1) As SquareCoordinate
            Dim x As Integer
            For x = 0 To PossibleMoves.Count - 1
                returnSC(x) = CType(PossibleMoves(x), SquareCoordinate)
            Next
            Return returnSC
        Else
            Dim returnSC() As SquareCoordinate
            Return returnSC
        End If
    End Function

    Private Function DetermineMoveEntry(ByVal aSC As SquareCoordinate) As MoveEntry
        Dim iRow As Integer = aSC.Row
        Dim iCol As Integer = aSC.Column

        'Horizontal
        Dim iGE1 As TicTacToe.GridEntry
        Dim iGE2 As TicTacToe.GridEntry

        Select Case iCol
            Case 0
                iGE1 = m_CurrentBoard.Square(iRow, 1)
                iGE2 = m_CurrentBoard.Square(iRow, 2)
            Case 1
                iGE1 = m_CurrentBoard.Square(iRow, 0)
                iGE2 = m_CurrentBoard.Square(iRow, 2)
            Case 2
                iGE1 = m_CurrentBoard.Square(iRow, 0)
                iGE2 = m_CurrentBoard.Square(iRow, 1)
        End Select

        If CheckPossibleWin(iGE1, iGE2, m_iAIPlayer) = True Then
            Return New MoveEntry(aSC, MoveType.Winning)
        End If

        If CheckPossibleWin(iGE1, iGE2, m_iHumanPlayer) = True Then
            Return New MoveEntry(aSC, MoveType.Blocking)
        End If

        'Vertical
        Select Case iRow
            Case 0
                iGE1 = m_CurrentBoard.Square(1, iCol)
                iGE2 = m_CurrentBoard.Square(2, iCol)
            Case 1
                iGE1 = m_CurrentBoard.Square(0, iCol)
                iGE2 = m_CurrentBoard.Square(2, iCol)
            Case 2
                iGE1 = m_CurrentBoard.Square(0, iCol)
                iGE2 = m_CurrentBoard.Square(1, iCol)
        End Select

        If CheckPossibleWin(iGE1, iGE2, m_iAIPlayer) = True Then
            Return New MoveEntry(aSC, MoveType.Winning)
        End If

        If CheckPossibleWin(iGE1, iGE2, m_iHumanPlayer) = True Then
            Return New MoveEntry(aSC, MoveType.Blocking)
        End If

        'Diagonal
        Select Case GetSquareType(aSC)
            Case SquareType.UpperLeft
                iGE1 = m_CurrentBoard.Square(1, 1)
                iGE2 = m_CurrentBoard.Square(2, 2)

                If CheckPossibleWin(iGE1, iGE2, m_iAIPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Winning)
                End If

                If CheckPossibleWin(iGE1, iGE2, m_iHumanPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Blocking)
                End If
            Case SquareType.UpperRight
                iGE1 = m_CurrentBoard.Square(1, 1)
                iGE2 = m_CurrentBoard.Square(2, 0)

                If CheckPossibleWin(iGE1, iGE2, m_iAIPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Winning)
                End If

                If CheckPossibleWin(iGE1, iGE2, m_iHumanPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Blocking)
                End If

            Case SquareType.Center
                '/
                iGE1 = m_CurrentBoard.Square(0, 2)
                iGE2 = m_CurrentBoard.Square(2, 0)

                If CheckPossibleWin(iGE1, iGE2, m_iAIPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Winning)
                End If

                If CheckPossibleWin(iGE1, iGE2, m_iHumanPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Blocking)
                End If
                '\
                iGE1 = m_CurrentBoard.Square(0, 0)
                iGE2 = m_CurrentBoard.Square(2, 2)

                If CheckPossibleWin(iGE1, iGE2, m_iAIPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Winning)
                End If

                If CheckPossibleWin(iGE1, iGE2, m_iHumanPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Blocking)
                End If
            Case SquareType.BottomLeft
                iGE1 = m_CurrentBoard.Square(1, 1)
                iGE2 = m_CurrentBoard.Square(0, 2)

                If CheckPossibleWin(iGE1, iGE2, m_iAIPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Winning)

                End If
                If CheckPossibleWin(iGE1, iGE2, m_iHumanPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Blocking)
                End If
            Case SquareType.BottomRight
                iGE1 = m_CurrentBoard.Square(1, 1)
                iGE2 = m_CurrentBoard.Square(0, 0)

                If CheckPossibleWin(iGE1, iGE2, m_iAIPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Winning)
                End If

                If CheckPossibleWin(iGE1, iGE2, m_iHumanPlayer) = True Then
                    Return New MoveEntry(aSC, MoveType.Blocking)
                End If
        End Select

        Return New MoveEntry(aSC, MoveType.Normal)
    End Function

    Private Function DetermineBestNormalMove(ByVal aNormalMoveList As ArrayList) As SquareCoordinate
        Dim aSC As SquareCoordinate
        Dim aME As MoveEntry
        Dim aGE1 As TicTacToe.GridEntry
        Dim aGE2 As TicTacToe.GridEntry


        For Each aME In aNormalMoveList
            aSC = aME.SquareCoordinate
            'Horizontal
            If CheckBestHorizontal(aSC) Then
                Return aSC
            End If

            If CheckBestVertical(aSC) Then
                Return aSC
            End If

            If GetSquareType(aSC) <> SquareType.Other Then
                If CheckBestDiagonal(aSC) Then
                    Return aSC
                End If
            End If
        Next
    End Function

    Private Function CheckBestHorizontal(ByVal aSC As SquareCoordinate) As Boolean
        Dim iCol As Integer = aSC.Column
        Dim iRow As Integer = aSC.Row
        Dim iGE1 As TicTacToe.GridEntry
        Dim iGE2 As TicTacToe.GridEntry

        Select Case iCol
            Case 0
                iGE1 = m_CurrentBoard.Square(iRow, 1)
                iGE2 = m_CurrentBoard.Square(iRow, 2)
            Case 1
                iGE1 = m_CurrentBoard.Square(iRow, 0)
                iGE2 = m_CurrentBoard.Square(iRow, 2)
            Case 2
                iGE1 = m_CurrentBoard.Square(iRow, 0)
                iGE2 = m_CurrentBoard.Square(iRow, 1)
        End Select

        If iGE1 = m_iHumanPlayer Or iGE2 = m_iHumanPlayer Then
            Return False
        End If

        If iGE1 = m_iAIPlayer Or iGE2 = m_iAIPlayer Then
            Return True
        End If

        Return False
    End Function

    Private Function CheckBestVertical(ByVal aSC As SquareCoordinate) As Boolean
        Dim iCol As Integer = aSC.Column
        Dim iRow As Integer = aSC.Row
        Dim iGE1 As TicTacToe.GridEntry
        Dim iGE2 As TicTacToe.GridEntry

        Select Case iRow
            Case 0
                iGE1 = m_CurrentBoard.Square(1, iCol)
                iGE2 = m_CurrentBoard.Square(2, iCol)
            Case 1
                iGE1 = m_CurrentBoard.Square(0, iCol)
                iGE2 = m_CurrentBoard.Square(2, iCol)
            Case 2
                iGE1 = m_CurrentBoard.Square(0, iCol)
                iGE2 = m_CurrentBoard.Square(1, iCol)
        End Select

        If iGE1 = m_iHumanPlayer Or iGE2 = m_iHumanPlayer Then
            Return False
        End If

        If iGE1 = m_iAIPlayer Or iGE2 = m_iAIPlayer Then
            Return True
        End If

        Return False
    End Function

    Private Function CheckBestDiagonal(ByVal aSC As SquareCoordinate) As Boolean
        Return False
    End Function

    Private Enum SquareType
        UpperLeft = 1
        UpperRight = 2
        Center = 3
        BottomLeft = 4
        BottomRight = 5
        Other = 6
    End Enum

    Private Function CheckPossibleWin(ByVal iGE1 As TicTacToe.GridEntry, _
                                      ByVal iGE2 As TicTacToe.GridEntry, _
                                      ByVal iPlayer As TicTacToe.GridEntry) As Boolean

        If (iGE1 = iPlayer) And (iGE2 = iPlayer) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetSquareType(ByVal aSC As SquareCoordinate) As SquareType
        If aSC.Row = 0 And aSC.Column = 0 Then Return SquareType.UpperLeft
        If aSC.Row = 0 And aSC.Column = 2 Then Return SquareType.UpperRight
        If aSC.Row = 1 And aSC.Column = 1 Then Return SquareType.Center
        If aSC.Row = 2 And aSC.Column = 0 Then Return SquareType.BottomLeft
        If aSC.Row = 2 And aSC.Column = 2 Then Return SquareType.BottomRight
        Return SquareType.Other
    End Function

    Enum MoveType
        Normal = 0
        Blocking = 1
        Winning = 2
    End Enum

    Private Class MoveEntry
        Private m_iMoveType As MoveType
        Private m_SquareCoordinate As SquareCoordinate

        Public Sub New(ByVal aSC As SquareCoordinate, ByVal iMoveType As MoveType)
            m_iMoveType = iMoveType
            m_SquareCoordinate = aSC
        End Sub

        Public ReadOnly Property [MoveType]() As MoveType
            Get
                Return m_iMoveType
            End Get
        End Property

        Public ReadOnly Property [SquareCoordinate]() As SquareCoordinate
            Get
                Return m_SquareCoordinate
            End Get
        End Property

        Public Overrides Function ToString() As String
            Dim sOutput As String
            sOutput = Me.MoveType.ToString & " : " & _
                      "[" & Me.m_SquareCoordinate.Row & ", " & _
                     Me.m_SquareCoordinate.Column & "]"
            Return sOutput
        End Function
    End Class
End Class
