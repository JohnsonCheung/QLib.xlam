Attribute VB_Name = "QXls_Feat_WsLnk"
Option Explicit
Private Const CMod$ = "MXls_GoWsLnk."
Private Const Asm$ = "QXls"
Private Type Lnkg
    FmCell As Range
    LnkToCell As Range
End Type
Private Type SomLnkg
    Som As Boolean
    Itm As Lnkg
End Type
Private Type Lnkgs: N As Long: Ay() As Lnkg: End Type
Sub CrtHypLnkzWsnA1(A1 As Range) 'WsnA1 is an A1 with cell value is Wsn
CrtHypLnks LnkgszWsnA1(A)
End Sub
Private Function LnkgszWsnA1(A1 As Range) As Lnkgs
Dim FmCell As Range
Dim LnkToCell As Range
Dim I As SomLnkg
While I.Som
    
End With
End Function

Private Function IsWsnCell(Cell As Range, Wsny$()) As Boolean
Dim V: V = Cell.Value
If Not IsStr(V) Then Exit Function
IsWsnCell = HasEle(Wsny, V)
End Function

Private Sub CrtHypLnks(A As Lnkgs)

End Sub
Private Sub CrtHypLnk(A As Lnkg)
With A.FmCell.Hyperlinks
    If .Count > 0 Then .Delete
    .Add , "", AdrzCell(A.JmpToCell)
End With
End Sub
Private Function LnkAdr$(A As Range)
LnkAdr = FmtQQ("'?'!?", Wsn, AdrzCell(A))
End Function
Function AdrzCell$(A As Range)
AdrzCell = A1zRg(A).Address
End Function
Private Function SomLnkg(Itm As Lnkg) As SomLnkg
With SomLnkg
    .Som = True
    .Itm = Itm
End With
End Function
Private Function LnkgszWsnA1_Ver1(A1 As Range) As Lnkgs

Dim R As Range: Set R = FstGoCell
Dim Wsny$():     Wsny = WsnyzRg(R)
Dim J%, Wsn$
While R.Value = "Go"
    J = J + 1: If J = 1000 Then ThwLoopingTooMuch CSub
    Wsn = CellRight(R).Value
    If HasEle(Wsny, Wsn) Then PushObj CellWsnItmAy, CellWsnItm(R, Wsn)
    Set R = CellBelow(R)
Wend
End Function
Private Function Lnkg(FmCell As Range, LnkToCell As Range) As Lnkg
With Lnkg
    .FmCell = FmCell
    .LnkToCell = LnkToCell
End With
End Function
Private Sub THwIf_NoSpace_ToFillWsn(At As Range)

End Sub
Function OthWsny(Ws As Worksheet) As String()

End Function

Sub FillWsn(At As Range)
THwIf_NoSpace_ToFillWsn At
Dim I As Range: Set I = At
Dim Ny$():         Ny = OthWsny(WszAt(At))
For Each Wsn In Ny
    I.Value = Wsn
    Set I = NxtCellBelow(I)
Next
End Sub
D