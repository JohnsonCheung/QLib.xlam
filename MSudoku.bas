Attribute VB_Name = "MSudoku"
Option Explicit
Private Type RRCC
    R1 As Byte
    R2 As Byte
    C1 As Byte
    C2 As Byte
End Type
Private Type SolveRslt
    Sq() As Variant
    HasSolve As Boolean
End Type
Private Type NineEleRslt
    NineEle() As Variant
    HasSolve As Boolean
End Type

Private Function RRCC(R1 As Byte, R2 As Byte, C1 As Byte, C2 As Byte) As RRCC
With RRCC
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
End Function

Private Function SolveFstRound(Sq()) As Variant()
Dim J%
For J = 1 To 9
    NineElezRow(Sq, J) = SolveNineElezFstRnd(NineElezRow(Sq, J))
Next
SolveFstRound = Sq
End Function

Private Function SolveSudokuzSq(Sq()) As Variant()
Dim O(), HasSolve As Boolean, J%
O = SolveFstRound(Sq)
HasSolve = True
While HasSolve
    J = J + 1: If J > 1000 Then Stop
    HasSolve = False
    With SolveRow(O): O = IIf(.HasSolve, .Sq, O): HasSolve = IIf(.HasSolve, True, HasSolve): End With
    With SolveCol(O): O = IIf(.HasSolve, .Sq, O): HasSolve = IIf(.HasSolve, True, HasSolve): End With
    'With SolveDiag(O): O = IIf(.HasSolve, .Sq, O): HasSolve = IIf(.HasSolve, True, HasSolve): End With
    With SolveSmallSq(O): O = IIf(.HasSolve, .Sq, O): HasSolve = IIf(.HasSolve, True, HasSolve): End With
Wend
SolveSudokuzSq = O
End Function

Private Function SolveRow(Sq()) As SolveRslt
Dim J%, O()
O = Sq
For J = 1 To 9
    With SolveNineEle(NineElezRow(O, J))
        If .HasSolve Then
            SolveRow.HasSolve = True
            NineElezRow(O, J) = .NineEle
        End If
    End With
Next
SolveRow.Sq = O
End Function

Property Get NineElezRow(Sq(), Row%) As Variant()
Dim J%
For J = 1 To 9
    PushI NineElezRow, Sq(Row, J)
Next
End Property

Property Let NineElezRow(Sq(), Row%, NineEle())
Dim J%
For J = 1 To 9
    Sq(Row, J) = NineEle(J - 1)
Next
End Property

Private Function SolveSmallSq(Sq()) As SolveRslt
Dim J%, O()
O = Sq
For J = 1 To 9
    With SolveNineEle(NineElezSmallSq(O, J))
        If .HasSolve Then
            SolveSmallSq.HasSolve = True
            NineElezSmallSq(O, J) = .NineEle
        End If
    End With
Next
SolveSmallSq.Sq = O
End Function

Property Get NineElezSmallSq(Sq(), J%) As Variant()
Dim R As Byte, C As Byte
With RRCCzJ(J)
For R = .R1 To .R2
    For C = .C1 To .C2
        PushI NineElezSmallSq, Sq(R, C)
    Next
Next
End With
End Property

Private Function RRCCzJ(J%) As RRCC
Select Case J
Case 1: RRCCzJ = RRCC(1, 3, 1, 3)
Case 2: RRCCzJ = RRCC(1, 3, 4, 6)
Case 3: RRCCzJ = RRCC(1, 3, 7, 9)
Case 4: RRCCzJ = RRCC(4, 6, 1, 3)
Case 5: RRCCzJ = RRCC(4, 6, 4, 6)
Case 6: RRCCzJ = RRCC(4, 6, 7, 9)
Case 7: RRCCzJ = RRCC(7, 9, 1, 3)
Case 8: RRCCzJ = RRCC(7, 9, 4, 6)
Case 9: RRCCzJ = RRCC(7, 9, 7, 9)
Case Else: Thw CSub, "Invalid J, should be 1 to 9", "J", J
End Select
End Function

Property Let NineElezSmallSq(Sq(), J%, NineEle())
Dim R As Byte, C As Byte
Dim I%
With RRCCzJ(J)
For R = .R1 To .R2
    For C = .C1 To .C2
        Sq(R, C) = NineEle(I)
        I = I + 1
    Next
Next
End With
End Property

Private Function SolveCol(Sq()) As SolveRslt
Dim J%, O()
O = Sq
For J = 1 To 9
    With SolveNineEle(NineElezCol(O, J))
        If .HasSolve Then
            SolveCol.HasSolve = True
            NineElezCol(O, J) = .NineEle
        End If
    End With
Next
SolveCol.Sq = O
End Function

Property Get NineElezCol(Sq(), Col%) As Variant()
Dim J%
For J = 1 To 9
    PushI NineElezCol, Sq(J, Col)
Next
End Property

Property Let NineElezCol(Sq(), Col%, NineEle())
Dim J%
For J = 1 To 9
    Sq(J, Col) = NineEle(J - 1)
Next
End Property

Private Function SolveDiag(Sq()) As SolveRslt
Dim J%, O()
O = Sq
With SolveNineEle(NineElezDiag1(O))
    If .HasSolve Then
        SolveDiag.HasSolve = True
        NineElezDiag1(O) = .NineEle
    End If
End With
With SolveNineEle(NineElezDiag2(O))
    If .HasSolve Then
        SolveDiag.HasSolve = True
        NineElezDiag2(O) = .NineEle
    End If
End With
SolveDiag.Sq = O
End Function

Private Property Get NineElezDiag1(Sq()) As Variant()
Dim J%
For J = 1 To 9
    PushI NineElezDiag1, Sq(J, J)
Next
End Property

Private Property Let NineElezDiag1(Sq(), NineEle())
Dim J%
For J = 1 To 9
    Sq(J, J) = NineEle(J - 1)
Next
End Property

Private Property Get NineElezDiag2(Sq()) As Variant()
Dim J%
For J = 1 To 9
    PushI NineElezDiag2, Sq(10 - J, 10 - J)
Next
End Property

Private Property Let NineElezDiag2(Sq(), NineEle())
Dim J%
For J = 1 To 9
    Sq(10 - J, 10 - J) = NineEle(J - 1)
Next
End Property

Private Function SolveNineElezFstRnd(NineEle()) As Variant()
Dim Should() As Byte: Should = ShouldBe(NineEle)
Dim J%, I
Dim O(): O = NineEle
For Each I In NineEle
    If IsEmpty(I) Then
        O(J) = Should
    End If
    J = J + 1
Next
SolveNineElezFstRnd = O
End Function

Private Function SolveNineEle(NineEle()) As NineEleRslt
Dim Should() As Byte: Should = ShouldBe(NineEle)
Dim O(): O = NineEle
Dim M
Dim I, J%
For Each I In NineEle
    If IsBytAy(I) Then
        M = Intersect(CvBytAy(I), Should)
        If Sz(I) > Sz(M) Then
            SolveNineEle.HasSolve = True
            O(J) = M
        End If
    Else
        If Not IsByt(I) Then Stop
    End If
    J = J + 1
Next
SolveNineEle.NineEle = O
End Function

Private Function Intersect(A() As Byte, B() As Byte)
Dim O: O = AyIntersect(A, B)
Intersect = IIf(Sz(O) = 1, O(0), O)
End Function

Private Function ShouldBe(NineEle()) As Byte()
Dim Certain() As Byte
Dim I
For Each I In NineEle
    If IsByt(I) Then PushI Certain, I
Next
Dim J As Byte
For J = 1 To 9
    If Not HasEle(Certain, J) Then PushI ShouldBe, J
Next
End Function

Sub SolveSudoku(Ws As Worksheet)
PutSudokuSolution Ws, SolveSudokuzSq(SqzSudoku(Ws))
End Sub

Private Function SqzSudoku(Ws As Worksheet) As Variant()
Dim O(): O = RgRCRC(A1zWs(Ws), 1, 1, 9, 9)
Dim I%, J%
For J = 1 To 9
    For I = 1 To 9
        If Not IsEmpty(O(J, I)) Then
            O(J, I) = CByte(O(J, I))
        End If
    Next
Next
SqzSudoku = O
End Function

Private Sub PutSudokuSolution(Ws As Worksheet, Sq())
SolutionRg(Ws).Value = Sq
End Sub

Private Function SolutionRg(Ws As Worksheet) As Range
Set SolutionRg = RgRCRC(A1zWs(Ws), 11, 1, 19, 9)
End Function

Private Property Get SampSudokuSq() As Variant()
Dim E: E = Empty
SampSudokuSq = SqzDry(Av( _
Array(5, E, 7, 6, 9, E, E, E, 2), _
Array(9, 3, E, E, E, 2, 7, 4, 5), _
Array(E, E, E, 3, E, 7, 1, E, E), _
Array(E, 4, 5, E, 6, E, 3, E, 8), _
Array(2, E, E, 4, E, E, E, E, E), _
Array(E, E, E, E, E, 8, 1, E, 2), _
Array(E, E, 9, E, 2, E, E, 1, 3), _
Array(3, E, E, E, E, 6, E, 5, 7), _
Array(7, E, E, 1, 3, E, 9, 8, 4)))
End Property

Private Sub Z_PutSampSudoku()
PutSampSudoku WsRC(ActiveSheet, 1, "L")
End Sub

Sub PutSampSudoku(At As Range)
RgRCRC(At, 1, 1, 9, 9).Value = SampSudokuSq
FmtSudoku At
End Sub

Sub FmtSudoku(At As Range)
BdrRgAround RgRCRC(At, 1, 1, 3, 3)
BdrRgAround RgRCRC(At, 1, 4, 3, 6)
BdrRgAround RgRCRC(At, 1, 7, 3, 9)
BdrRgAround RgRCRC(At, 4, 1, 6, 3)
BdrRgAround RgRCRC(At, 4, 4, 6, 6)
BdrRgAround RgRCRC(At, 4, 7, 6, 9)
BdrRgAround RgRCRC(At, 7, 1, 9, 3)
BdrRgAround RgRCRC(At, 7, 4, 9, 6)
BdrRgAround RgRCRC(At, 7, 7, 9, 9)
RgCC(At, 1, 9).EntireColumn.ColumnWidth = 2
End Sub

Private Sub Z_SolveSudoku()
Dim Ws As Worksheet
GoSub T0
Exit Sub
T0:
    Set Ws = ActiveSheet
    GoTo Tst
Tst:
    SolveSudoku Ws
    Return
End Sub

