Attribute VB_Name = "MDao_Tbl_Upd_EndDteFld"
Option Explicit


Sub UpdEndDte(A As Database, T, EndDteFld$, BegDteFld$, GpFF)
Dim LasBegDte As Date
LasBegDte = DateSerial(2099, 12, 31)
Dim Q$
''Q = SqlSel_FF_Fm_Ordff(Sy(BegDteFld, EndDteFld), T, BegDteFld)
Stop
With Rs(A, Q)
    While Not .EOF
        .Edit
        .Fields(EndDteFld).Value = LasBegDte
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

