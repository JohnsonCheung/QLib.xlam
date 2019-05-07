Attribute VB_Name = "QDao_Db_LnkTbl"
Option Explicit
Private Const CMod$ = "BLnkTbl."
Type LnkTblPm: T As String: S As String: Cn As String: End Type
Type LnkTblPms: N As Integer: Ay() As LnkTblPm: End Type

Sub PushLnkTblPm(O As LnkTblPms, M As LnkTblPm)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function LnkTblPm(T$, S$, Cn$) As LnkTblPm
With LnkTblPm: .T = T: .S = S: .Cn = Cn: End With
End Function

Sub LnkTblzPms(A As Database, B As LnkTblPms)
Dim Ay() As LnkTblPm: Ay = B.Ay
Dim J%
For J = 0 To B.N - 1
    LnkTblzPm A, Ay(J)
Next
End Sub
Sub LnkTblzPm(A As Database, B As LnkTblPm)
With B
    LnkTbl A, .T, .S, .Cn
End With
End Sub

