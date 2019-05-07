Attribute VB_Name = "QVb_Ay_FmTo_To_Dic"
Option Explicit
Private Const CMod$ = "MVb_Ay_FmTo_To_Dic."
Private Const Asm$ = "QVb"
Function IxDiczAy(Ay) As Dictionary
Dim O As New Dictionary, J&
For J = 0 To UB(Ay)
    If Not O.Exists(Ay(J)) Then
        O.Add Ay(J), J
    End If
Next
Set IxDiczAy = O
End Function

Function IdCntDiczAy(Ay) As Dictionary
'Type DistIdCntDic = Map Val [Id,Cnt]
Dim X, O As New Dictionary, J&, IdCnt()
For Each X In Itr(Ay)
    If Not O.Exists(X) Then
        O.Add X, Array(J, 1)
        J = J + 1
    Else
        IdCnt = O(X)
        O(X) = Array(IdCnt(0), IdCnt(1) + 1)
    End If
Next
Set IdCntDiczAy = O
End Function
