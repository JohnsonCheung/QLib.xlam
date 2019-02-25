Attribute VB_Name = "MDao_Tbl_Upd_EndDteFld"
Option Explicit


Sub UpdEndDte(T, EndDteFld$, BegDteFld$, GpFF)
UpdEndDtez CDb, T, EndDteFld, BegDteFld, GpFF
End Sub

Sub UpdEndDtez(A As Database, T, EndDteFld$, BegDteFld$, GpFF)
Dim LasBegDte As Date
LasBegDte = DateSerial(2099, 12, 31)
With Rsz(A, SqlSel_FF_Fm_OrdFF(Sy(BegDteFld, EndDteFld), T, BegDteFld))
    While Not .EOF
        .Edit
        .Fields(EndDteFld).Value = LasBegDte
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

