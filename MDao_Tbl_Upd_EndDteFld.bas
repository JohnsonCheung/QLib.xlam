Attribute VB_Name = "MDao_Tbl_Upd_EndDteFld"
Option Explicit


Sub UpdEndDte(A As Database, T, EndDteFld$, BegDteFld$, GpFF)
Dim LasBegDte As Date
LasBegDte = DateSerial(2099, 12, 31)
With Rs(A, SqlSel_FF_Fm_Ordff(Sy(BegDteFld, EndDteFld), T, BegDteFld))
    While Not .EOF
        .Edit
        .Fields(EndDteFld).Value = LasBegDte
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

