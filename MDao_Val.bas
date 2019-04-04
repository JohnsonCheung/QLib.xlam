Attribute VB_Name = "MDao_Val"
Option Explicit
Const CMod$ = "MDaoRetrieveVal."

Private Sub Z_ValOfQ()
Dim D As Database
Ept = CByte(18)
Act = ValOfQ(D, "Select Y from [^YM]")
C
End Sub

Property Get ValOfQ(A As Database, Sql)
ValOfQ = ValOfRs(A.OpenRecordset(Sql))
End Property
Property Let ValOfQ(A As Database, Sql, V)
ValOfRs(A.OpenRecordset(Sql)) = V
End Property

Property Let ValOfSsk(Db As Database, T, F, Sskv, V)
ValOfRs(Rs(Db, SqlSel_F_Fm_F_Ev(F, T, SskFld(Db, T), V))) = V
End Property

Property Get ValOfSsk(Db As Database, T, F, Sskv)
Dim Ssk$: Ssk = SskFld(Db, T)
ValOfSsk = ValOfRs(Rs(Db, SqlSel_F_Fm_F_Ev(F, T, Ssk, Sskv)))
End Property
Function ValOfTF(A As Database, T, F)
ValOfTF = A.TableDefs(T).OpenRecordset.Fields(F).Value
End Function
Function ValOfQQ(A As Database, QQSql, ParamArray Ap())
Dim Av(): Av = Ap
ValOfQQ = ValOfQ(A, FmtQQAv(QQSql, Av))
End Function

Property Let ValOfRs(A As Dao.Recordset, V)
If NoRec(A) Then
    A.AddNew
Else
    A.Edit
End If
A.Fields(0).Value = V
A.Update
End Property

Property Get ValOfRs(A As Dao.Recordset)
If NoRec(A) Then Exit Property
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Property
ValOfRs = V
End Property

Property Let ValOfRsFld(Rs As Dao.Recordset, Fld, V)
With Rs
    .Edit
    .Fields(Fld).Value = V
    .Update
End With
End Property

Property Get ValOfRsFld(Rs As Dao.Recordset, Fld)
With Rs
    If .EOF Then Exit Property
    If .BOF Then Exit Property
    ValOfRsFld = .Fields(Fld).Value
End With
End Property
