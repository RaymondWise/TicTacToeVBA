Attribute VB_Name = "TicTacToe"
Option Explicit
Public Sub TicTacToeMain()
    Const DELIMITER As String = " "
    Dim gameResults As String
    Dim moves As String
    Dim totalGames As Long
    totalGames = Cells(1, 1)
    Dim i As Long
    
    CreateBoard
    
    For i = 2 To totalGames + 1
       moves = Cells(i, 1)
       PlayGame moves, DELIMITER, gameResults
    Next
    Cells(1, 10) = Trim(gameResults)
End Sub
Private Sub PlayGame(ByVal moves As String, ByVal DELIMITER As String, ByRef gameResults As String)
    Const NAMED_RANGE As String = "_"
    Dim isWon As Boolean
    Dim MoveArray() As Long
    ReDim MoveArray(1 To 9)
    MoveArray = GetMoves(moves, DELIMITER)
    
    ClearBoard
    
    Dim move As Long
    For move = 1 To 9
        Select Case move
            Case 1, 3, 5, 7, 9
                ActiveSheet.Range(NAMED_RANGE & MoveArray(move)) = "X"
            Case 2, 4, 6, 8
                ActiveSheet.Range(NAMED_RANGE & MoveArray(move)) = "O"
        End Select
        If move > 4 Then isWon = CheckWin(MoveArray(move))
        If isWon Then
            gameResults = gameResults & move & DELIMITER
            Exit Sub
        End If
    Next
    gameResults = gameResults & "0" & DELIMITER
End Sub
Private Sub CreateBoard()
    With ActiveSheet
        Range("B1").Name = "_1"
        Range("C1").Name = "_2"
        Range("D1").Name = "_3"
        Range("B2").Name = "_4"
        Range("C2").Name = "_5"
        Range("D2").Name = "_6"
        Range("B3").Name = "_7"
        Range("C3").Name = "_8"
        Range("D3").Name = "_9"
    End With
End Sub
Private Sub ClearBoard()
    Range("B1:D3").Clear
End Sub

Private Function CheckWin(ByVal square As Long) As Boolean
    CheckWin = False
    With ActiveSheet
        Select Case square
            Case 1
                If .Range("_1") = .Range("_2") And .Range("_1") = .Range("_3") Then CheckWin = True
                If .Range("_1") = .Range("_4") And .Range("_1") = .Range("_7") Then CheckWin = True
                If .Range("_1") = .Range("_5") And .Range("_1") = .Range("_9") Then CheckWin = True
            Case 2
                If .Range("_2") = .Range("_1") And .Range("_2") = .Range("_3") Then CheckWin = True
                If .Range("_2") = .Range("_5") And .Range("_2") = .Range("_8") Then CheckWin = True
            Case 3
                If .Range("_3") = .Range("_2") And .Range("_3") = .Range("_1") Then CheckWin = True
                If .Range("_3") = .Range("_6") And .Range("_3") = .Range("_9") Then CheckWin = True
                If .Range("_3") = .Range("_5") And .Range("_3") = .Range("_7") Then CheckWin = True
            Case 4
                If .Range("_4") = .Range("_1") And .Range("_4") = .Range("_7") Then CheckWin = True
                If .Range("_4") = .Range("_5") And .Range("_4") = .Range("_6") Then CheckWin = True
            Case 5
                If .Range("_5") = .Range("_1") And .Range("_5") = .Range("_9") Then CheckWin = True
                If .Range("_5") = .Range("_2") And .Range("_5") = .Range("_8") Then CheckWin = True
                If .Range("_5") = .Range("_3") And .Range("_5") = .Range("_7") Then CheckWin = True
                If .Range("_5") = .Range("_4") And .Range("_5") = .Range("_6") Then CheckWin = True
            Case 6
                If .Range("_6") = .Range("_3") And .Range("_6") = .Range("_9") Then CheckWin = True
                If .Range("_6") = .Range("_5") And .Range("_6") = .Range("_4") Then CheckWin = True
            Case 7
                If .Range("_7") = .Range("_8") And .Range("_7") = .Range("_9") Then CheckWin = True
                If .Range("_7") = .Range("_1") And .Range("_7") = .Range("_4") Then CheckWin = True
                If .Range("_7") = .Range("_3") And .Range("_7") = .Range("_5") Then CheckWin = True
            Case 8
                If .Range("_8") = .Range("_2") And .Range("_8") = .Range("_5") Then CheckWin = True
                If .Range("_8") = .Range("_7") And .Range("_8") = .Range("_9") Then CheckWin = True
            Case 9
                If .Range("_9") = .Range("_1") And .Range("_9") = .Range("_5") Then CheckWin = True
                If .Range("_9") = .Range("_3") And .Range("_9") = .Range("_6") Then CheckWin = True
                If .Range("_9") = .Range("_7") And .Range("_9") = .Range("_8") Then CheckWin = True
        End Select
    End With
End Function
Private Function GetMoves(ByVal moves As String, ByVal DELIMITER As String) As Long()
    Dim rawData As Variant
    Dim tempMoves() As Long
    ReDim tempMoves(1 To 9)
    rawData = Split(moves, DELIMITER)
    ReDim GetMoves(1 To 9)
    Dim i As Long
    For i = 1 To 9
        tempMoves(i) = Int(rawData(i - 1))
    Next
    GetMoves = tempMoves
End Function
