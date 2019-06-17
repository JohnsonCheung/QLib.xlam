Attribute VB_Name = "QDao_SqlPm"
Option Explicit
Option Compare Text
Type SelIntoPm: Fny() As String: Ey() As String: Into As String: T As String: Bexp As String: End Type
Type SelIntoPms: N As Byte: Ay() As SelIntoPm: End Type
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_SqlPm."
Function SelIntoPm(Fny$(), Ey$(), Into$, T$, Optional Bexp$) As SelIntoPm
With SelIntoPm
    .Fny = Fny
    .Ey = Ey
    .Into = Into
    .T = T
    .Bexp = Bexp
End With
End Function

Sub PushIelIntoPm(O As SelIntoPms, M As SelIntoPm)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function SqyzSelIntoPms(A As SelIntoPms) As String()
Dim J As Byte
For J = 0 To A.N - 1
    PushI SqyzSelIntoPms, SqlzSelIntoPm(A.Ay(J))
Next
End Function

Function SqlzSelIntoPm$(A As SelIntoPm)
With A
'SqlzSelIntoPm = SqlSel_Fny_Extny_Into_T(.Fny, .Extny, .Into, .T, .Bexp)
End With
End Function

