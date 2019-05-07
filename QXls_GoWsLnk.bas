Attribute VB_Name = "QXls_GoWsLnk"
Option Explicit
Private Const CMod$ = "MXls_GoWsLnk."
Private Const Asm$ = "QXls"
Private Sub CrtGoLnkForOne(Cell As Range, Wsn$)
Dim A1 As Range: Set A1 = A1zRg(Cell)
With A1.Hyperlinks
    If .Count > 0 Then .Delete
    .Add A1, "", FmtQQ("'?'!A1", Wsn)
End With
End Sub
Private Function CvCellWsnItm(A) As CellWsnItm
Set CvCellWsnItm = A
End Function

Private Function CellWsnItmAy(FstGoCell) As CellWsnItm()
Dim R As Range: Set R = FstGoCell
Dim Wsny$():     Wsny = WsNyzRg(R)
Dim J%, Wsn$
While R.Value = "Go"
    J = J + 1: If J = 1000 Then ThwLoopingTooMuch CSub
    Wsn = CellRight(R).Value
    If HasEle(Wsny, Wsn) Then PushObj CellWsnItmAy, CellWsnItm(R, Wsn)
    Set R = CellBelow(R)
Wend
End Function
Private Function CellWsnItm(Cell As Range, Wsn$) As CellWsnItm
Set CellWsnItm = New CellWsnItm
With CellWsnItm
    Set .Cell = Cell
    .Wsn = Wsn
End With
End Function
Sub CrtGoLnk(FstGoCell As Range)
Dim I
For Each I In Itr(CellWsnItmAy(FstGoCell))
    With CvCellWsnItm(I)
        CrtGoLnkForOne .Cell, .Wsn
    End With
Next
End Sub
Private Function IsOkToFill(A As Range) As Boolean
IsOkToFill = IsEmpty(A.Value) And IsEmpty(CellRight(A))
End Function
Sub FillGoWs(FstGoCell As Range)
Dim R As Range:     Set FstGoCell = R
Dim Wsny$():                 Wsny = WsNyzRg(R)
Dim IsFill As Boolean:     IsFill = IsOkToFill(R)
Dim I%
While IsFill
    R.Value = "Go"
    CellRight(R).Value = Wsny(I)
    I = I + 1
Wend
End Sub
