Attribute VB_Name = "QDao_Ssk"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Ssk."
Public Const C_SkNm$ = "SecondaryKey"
Public Const C_PkNm$ = "PrimaryKey"

Function SkFnyzTd(T As Dao.TableDef) As String()
SkFnyzTd = FnyzIdx(SkIdxzTd(T))
End Function

Function SkSqlQPfxSy(A As Database, T) As String()
Stop '
End Function
Function SkFny(A As Database, T) As String()
SkFny = SkFnyzTd(A.TableDefs(T))
End Function

Function Sskv(A As Database, T) As Aset
'SSskv is [S]ingleFielded [S]econdKey [K]ey [V]alue [Aset], which is always a Value-Aset.
'and Ssk is a field-name from , which assume there is a Unique-Index with name "SecordaryKey" which is unique and and have only one field
'Set Sskv = ColSet(SskFld)
End Function

Function SkIdxzTd(T As Dao.TableDef) As Dao.Index
Set SkIdxzTd = IdxzTd(T, C_SkNm)
End Function

Function SkIdx(A As Database, T) As Dao.Index
Set SkIdx = Idx(A, T, C_SkNm)
End Function

Function SskFld$(A As Database, T)
Dim Sk$(): Sk = SkFny(A, T): If Si(Sk) = 1 Then SskFld = Sk(0): Exit Function
Thw CSub, "SkFny-Sz<>1", "Db T, SkFny-Si SkFny", Dbn(A), T, Si(Sk), Sk
End Function

Sub DltRecNotInSskv(A As Database, SskTbl$, NotInSSskv As Aset) _
'Delete Db-T record for those record's Sk not in NotInSSskv, _
'Assume T has single-fld-sk
Const CSub$ = CMod & "DltRecNotInSskv"
'If NotInSSskv.IsEmp Then Thw CSub, "Given NotInSSskv cannot be empty", "Db SskTbl SskFld", Dbn(A), SskTbl, SskFld_Dbt(Db, SskTbl)
Dim Q$, Excess As Aset
Set Excess = SskVset(A, SskTbl).Minus(NotInSSskv)
If Excess.IsEmp Then Exit Sub
'RunSqy Db, SqyDlt_Fm_WhFld_InAset(SskTbl, SskFld_Dbt(Db, SskTbl), Excess)
End Sub

Function AsetzDbtf(A As Database, T, F$) As Aset
Set AsetzDbtf = AsetzRs(Rs(A, SqlSel_F_T(F, T)), F)
End Function

Function SskVset(A As Database, T) As Aset
Set SskVset = AsetzDbtf(A, T, SskFld(A, T))
End Function

Sub InsReczSskv(A As Database, T, ToInsSSskv As Aset) _
'Insert Single-Field-Secondary-Key-Aset-A into Dbt
'Assume T has single-fld-sk and can be inserted by just giving such SSk-value
Dim ShouldIns As Aset
    Set ShouldIns = ToInsSSskv.Minus(SskVset(A, T))
If ShouldIns.IsEmp Then Exit Sub
Dim I, F$
'F = SskFld_Dbt(A, T)
'With RsDbt(A, T)
'    For Each I In ShouldIns.Itms
'        .AddNew
'        .Fields(F).Value = I
'        .Update
'    Next
'    .Close
'End With
End Sub

Private Sub ZZ()
MDao_DML_SngFldSkTbl_Operation:
End Sub

