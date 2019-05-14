Attribute VB_Name = "QDao_SqlPm"
Type SelIntoPm: Fny() As String: Ey() As String: Into As String: T As String: Bexpr As String: End Type
Type SelIntoPms: N As Byte: Ay() As SelIntoPm: End Type
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_SqlPm."
Function SelIntoPm(Fny$(), Ey$(), Into$, T$, Optional Bexpr$) As SelIntoPm
With SelIntoPm
    .Fny = Fny
    .Ey = EX
    .Into = Into
    .T = T
    .Bexpr = Bexpr
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
'SqlzSelIntoPm = SqlSel_Fny_ExtNy_Into_T(.Fny, .ExtNy, .Into, .T, .Bexpr)
End With
End Function

