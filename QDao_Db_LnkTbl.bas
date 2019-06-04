Attribute VB_Name = "QDao_Db_LnkTbl"
Option Compare Text
Option Explicit
Private Const CMod$ = "BLnkTbl."
Type LnkTblPm: T As String: S As String: Cn As String: End Type
Type LnkTblPms: N As Long: Ay() As LnkTblPm: End Type
Function TnyzL(A As LnkTblPms) As String()
Dim J&
For J = 0 To A.N - 1
    PushI TnyzL, RmvFstChr(A.Ay(J).T)
Next
End Function
Function AddLnkTblPms(A As LnkTblPms, B As LnkTblPms) As LnkTblPms
AddLnkTblPms = A
PushLnkTblPms AddLnkTblPms, B
End Function
Sub PushLnkTblPms(O As LnkTblPms, M As LnkTblPms)
Dim J&
For J = 0 To M.N - 1
    PushLnkTblPm O, M.Ay(J)
Next
End Sub

Sub PushLnkTblPm(O As LnkTblPms, M As LnkTblPm)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function LnkTblPm(T, S, Cn$) As LnkTblPm
With LnkTblPm: .T = T: .S = S: .Cn = Cn: End With
End Function
Function LnkTblzDrs(A As Database, DrsTSCn As Drs) As Unt
LnkTblzDryTSCn A, DrsTSCn.Dry
End Function

Sub LnkTblzDryTSCn(A As Database, DryTSCn())
Dim Dr, T$, S$, Cn$
For Each Dr In Itr(DryTSCn)
    T = Dr(0)
    S = Dr(1)
    Cn = Dr(2)
    LnkTbl A, T, S, Cn
Next
End Sub
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

