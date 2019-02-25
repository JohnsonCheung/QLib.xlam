Attribute VB_Name = "MDao_Val"
Option Explicit
Const CMod$ = "MDaoRetrieveVal."

Private Sub Z_ValzQ()
Ept = CByte(18)
Act = ValzQ("Select Y from [^YM]")
C
End Sub

Function ValzQ(Q)
ValzQ = ValzDbq(CDb, Q)
End Function

Property Get ValzDbq(Db As Database, Sql)
ValzDbq = ValzRs(Db.OpenRecordset(Sql))
End Property
Property Let ValzDbq(Db As Database, Sql, V)
ValzRs(Db.OpenRecordset(Sql)) = V
End Property

Property Let ValzSskDb(Db As Database, T, F, Sskv, V)
Dim Ssk$: Ssk = SskFldz(Db, T)
ValzRs(Rsz(Db, SqlSel_F_Fm_F_Ev(F, T, Ssk, V))) = V
End Property

Property Get ValzSskDb(Db As Database, T, F, Sskv)
Dim Ssk$: Ssk = SskFldz(Db, T)
ValzSskDb = ValzRs(Rsz(Db, SqlSel_F_Fm_F_Ev(F, T, Ssk, Sskv)))
End Property
Function ValzTF(T, F)
ValzTF = ValzDbtf(CDb, T, F)
End Function
Function ValzDbtf(Db As Database, T, F)
ValzDbtf = Db.TableDefs(T).OpenRecordset.Fields(F).Value
End Function
