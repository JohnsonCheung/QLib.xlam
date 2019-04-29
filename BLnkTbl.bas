Attribute VB_Name = "BLnkTbl"
Option Explicit
Type LnkTblPm: T As String: S As String: Cn As String: End Type
Type LnkTblPms: N As Integer: Ay() As LnkTblPm: End Type

Sub PushLnkTblPm(O As LnkTblPms, M As LnkTblPm)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function NewLnkTblPm(T$, S$, Cn$) As LnkTblPm
With LnkPm: .T = T: .S = S: .Cn = Cn: End With
End Function

Sub LnkTblzPms(A As Database, B As LnkTblPms)
Dim Ay() As LnkPm: Ay = B.Ay
For J = 0 To B.N
    LnkTblzPm A, Ay(J)
Next
End Sub
Sub LnkTblzPm(A As Database, B As LnkTblPm)
With B: LnkTbl A, .T, .S, .Cn
End Sub

