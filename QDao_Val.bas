Attribute VB_Name = "QDao_Val"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Val."

Private Sub Z_ValzQ()
Dim D As Database
Ept = CByte(18)
Act = ValzQ(D, "Select Y from [^YM]")
C
End Sub

Property Get ValzQ(A As Database, Q)
ValzQ = ValzRs(A.OpenRecordset(Q))
End Property
Property Let ValzQ(A As Database, Q, V)
ValzRs(A.OpenRecordset(Q)) = V
End Property

Property Let ValzSsk(A As Database, T, F$, Sskv(), V)
ValzRs(Rs(A, SqlSel_F_T_F_Ev(F, T, SskFld(A, T), Sskv))) = V
End Property

Property Get ValzSsk(A As Database, T, F$, Sskv())
Dim Ssk$: Ssk = SskFld(A, T)
ValzSsk = ValzRs(Rs(A, SqlSel_F_T_F_Ev(F, T, Ssk, Sskv)))
End Property

Function ValzTF(A As Database, T, F)
ValzTF = A.TableDefs(T).OpenRecordset.Fields(F).Value
End Function
Function ValzQQ(A As Database, QQSql$, ParamArray Ap())
Dim Av(): Av = Ap
ValzQQ = ValzQ(A, FmtQQAv(QQSql, Av))
End Function

Property Let ValzRs(A As DAO.Recordset, V)
If NoRec(A) Then
    A.AddNew
Else
    A.Edit
End If
A.Fields(0).Value = V
A.Update
End Property

Property Get ValzRs(A As DAO.Recordset)
If NoRec(A) Then Exit Property
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Property
ValzRs = V
End Property

Property Let ValzRsFld(Rs As DAO.Recordset, Fld, V)
With Rs
    .Edit
    .Fields(Fld).Value = V
    .Update
End With
End Property

Property Get ValzRsFld(Rs As DAO.Recordset, Fld)
With Rs
    If .EOF Then Exit Property
    If .BOF Then Exit Property
    ValzRsFld = .Fields(Fld).Value
End With
End Property
