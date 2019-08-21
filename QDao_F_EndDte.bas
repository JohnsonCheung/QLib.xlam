Attribute VB_Name = "QDao_F_EndDte"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Tbl_Upd_EndDteFld."
Private Const Asm$ = "QDao"


Sub UpdEndDte(D As Database, T, EndDteFld$, BegDteFld$, GpFF)
Dim LasBegDte As Date
LasBegDte = DateSerial(2099, 12, 31)
Dim Q$
''Q = SqlSel_FF_Fm_Ordff(Sy(BegDteFld, EndDteFld), T, BegDteFld)
Stop
With Rs(D, Q)
    While Not .EOF
        .Edit
        .Fields(EndDteFld).Value = LasBegDte
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

