Attribute VB_Name = "QDao_Rs_Mdy"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Rs_Mdy."
Private Const Asm$ = "QDao"

Sub InsRszDy(A As DAO.Recordset, Dy())
Dim Dr
With A
    For Each Dr In Itr(Dy)
        InsRs A, Dr
    Next
End With
End Sub


Sub SetRs(Rs As DAO.Recordset, Dr)
If Si(Dr) <> Rs.Fields.Count Then
    Thw CSub, "Si of Rs & Dr are diff", _
        "Si-Rs and Si-Dr Rs-Fny Dr", Rs.Fields.Count, Si(Dr), Itn(Rs.Fields), Dr
End If
Dim V, J%
For Each V In Dr
    If IsEmpty(V) Then
        Rs(J).Value = Rs(J).DefaultValue
    Else
        Rs(J).Value = V
    End If
    J = J + 1
Next
End Sub


Sub InsRszAp(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
InsRs Rs, Dr
End Sub

Sub InsRs(Rs As DAO.Recordset, Dr)
Rs.AddNew
SetRs Rs, Dr
Rs.Update
End Sub

Sub UpdRszAp(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
UpdRs Rs, Dr
End Sub


Sub DltRs(A As DAO.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub

