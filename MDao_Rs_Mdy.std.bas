Attribute VB_Name = "MDao_Rs_Mdy"
Option Explicit

Sub InsRszDry(A As Dao.Recordset, Dry())
Dim Dr
With A
    For Each Dr In Itr(Dry)
        UpdRs A, Dr
    Next
End With
End Sub


Property Let ValzRsFld(Rs As Dao.Recordset, Fld, V)
With Rs
    .Edit
    .Fields(Fld).Value = V
    .Update
End With
End Property

Property Get ValzRsFld(Rs As Dao.Recordset, Fld)
With Rs
    If .EOF Then Exit Property
    If .BOF Then Exit Property
    ValzRsFld = .Fields(Fld).Value
End With
End Property
Sub SetRs(Rs As Dao.Recordset, Dr)
If Sz(Dr) = Rs.Fields.Count Then
    Thw CSub, "Sz of Rs & Dr are diff", _
        "Sz-Rs and Sz-Dr Rs-Fny Dr", Rs.Fields.Count, Sz(Dr), Itn(Rs.Fields), Dr
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


Sub InsRszAp(Rs As Dao.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
InsRs Rs, Dr
End Sub

Sub InsRs(Rs As Dao.Recordset, Dr)
Rs.AddNew
SetRs Rs, Dr
Rs.Update
End Sub

Sub UpdRszAp(Rs As Dao.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
UpdRs Rs, Dr
End Sub


Sub DltRs(A As Dao.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub

