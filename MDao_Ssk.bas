Attribute VB_Name = "MDao_Ssk"
Option Explicit
Const CMod$ = "MDao_DML_SngFldSkTbl_Operation."

Sub DltRecNotInSskv(Db As Database, SskTbl, NotInSSskv As Aset) _
'Delete Db-T record for those record's Sk not in NotInSSskv, _
'Assume T has single-fld-sk
Const CSub$ = CMod & "DltRecNotInSskv"
'If NotInSSskv.IsEmp Then Thw CSub, "Given NotInSSskv cannot be empty", "Db SskTbl SskFld", DbNm(Db), SskTbl, SskFld_Dbt(Db, SskTbl)
Dim Q$, Excess As Aset
Set Excess = SskVset(Db, SskTbl).Minus(NotInSSskv)
If Excess.IsEmp Then Exit Sub
'RunSqy Db, SqyDlt_Fm_WhFld_InAset(SskTbl, SskFld_Dbt(Db, SskTbl), Excess)
End Sub
Function AsetzDbtf(A As Database, T, F) As Aset
Set AsetzDbtf = AsetzRs(Rs(A, SqlSel_F_Fm(F, T)))
End Function

Function SskVset(Db As Database, T) As Aset
Set SskVset = AsetzDbtf(Db, T, SskFld(Db, T))
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

Private Sub Z()
MDao_DML_SngFldSkTbl_Operation:
End Sub
