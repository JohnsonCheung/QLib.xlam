Attribute VB_Name = "MDao_Val"
Option Explicit
Const CMod$ = "MDaoRetrieveVal."

Private Sub Z_ValzQ()
Dim D As Database
Ept = CByte(18)
Act = ValzQ(D, "Select Y from [^YM]")
C
End Sub

Property Get ValzQ(A As Database, Sql)
ValzQ = ValzRs(A.OpenRecordset(Sql))
End Property
Property Let ValzQ(A As Database, Sql, V)
ValzRs(A.OpenRecordset(Sql)) = V
End Property

Property Let ValzSsk(Db As Database, T, F, Sskv, V)
ValzRs(Rs(Db, SqlSel_F_Fm_F_Ev(F, T, SskFld(Db, T), V))) = V
End Property

Property Get ValzSsk(Db As Database, T, F, Sskv)
Dim Ssk$: Ssk = SskFld(Db, T)
ValzSsk = ValzRs(Rs(Db, SqlSel_F_Fm_F_Ev(F, T, Ssk, Sskv)))
End Property
Function ValzTF(A As Database, T, F)
ValzTF = A.TableDefs(T).OpenRecordset.Fields(F).Value
End Function
