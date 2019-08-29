Attribute VB_Name = "QDao_F_DaoVal"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Val."

Sub SetVzQ(D As Database, Q, V)
VzRs(D.OpenRecordset(Q)) = V
End Sub

Sub SetVzRs(A As Dao.Recordset, V)
If NoRec(A) Then
    A.AddNew
Else
    A.Edit
End If
A.Fields(0).Value = V
A.Update
End Sub

Sub SetVzRsFld(Rs As Dao.Recordset, Fld, V)
With Rs
    .Edit
    .Fields(Fld).Value = V
    .Update
End With
End Sub

Sub SetVzSsk(D As Database, T, F$, Sskv(), V)
VzRs(Rs(D, SqlSel_F_T_F_Ev(F, T, SskFld(D, T), Sskv))) = V
End Sub

Function VzQ(D As Database, Q)
VzQ = VzRs(D.OpenRecordset(Q))
End Function

Function VzQQ(D As Database, QQSql$, ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
VzQQ = VzQ(D, FmtQQAv(QQSql, Av))
End Function

Function VzRs(A As Dao.Recordset)
If NoRec(A) Then Exit Function
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Function
VzRs = V
End Function

Function VzRsFld(Rs As Dao.Recordset, Fld)
With Rs
    If .EOF Then Exit Function
    If .BOF Then Exit Function
    VzRsFld = .Fields(Fld).Value
End With
End Function

Function VzSsk(D As Database, T, F$, Sskv())
Dim Ssk$: Ssk = SskFld(D, T)
VzSsk = VzRs(Rs(D, SqlSel_F_T_F_Ev(F, T, Ssk, Sskv)))
End Function

Function VzTF(D As Database, T, F)
VzTF = D.TableDefs(T).OpenRecordset.Fields(F).Value
End Function

Private Sub Z_VzQ()
Dim D As Database
Ept = CByte(18)
Act = VzQ(D, "Select Y from [^YM]")
C
End Sub
