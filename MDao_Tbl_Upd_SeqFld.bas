Attribute VB_Name = "MDao_Tbl_Upd_SeqFld"
Option Explicit

Sub UpdSeqFld(A As Database, T, SeqFld$, GpFF, OrdffMinus$)
Dim Q$: Q = SqlSel_FF_Fm_Ordff(SeqFld & " " & GpFF, T, OrdffMinus)
Dim R As Recordset: Set R = Rs(A, Q)
If NoRec(R) Then Exit Sub
Dim Seq&, Las(), Cur(), N%
With R
    N = .Fields.Count - 1
    .MoveNext
    Las = DrzRs(R)
    While Not .EOF
        Cur = DrzSqR(R, N)
        If Not IsEqAy(Cur, Las) Then
            Cur = Las
            Seq = 0
        End If
        Seq = Seq + 1
        UpdRs R, Array(Seq)
        .MoveNext
    Wend
End With
End Sub


Private Sub ZZ_UpdSeqFld()
Dim Db As Database, T
Set Db = TmpDb
RunQ Db, "Select * into [#A] from [T] order by Sku,PermitDate"
RunQ Db, "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
UpdSeqFld Db, T, "BchRateSeq", "Sku", "Sku Rate"
Stop
DrpT Db, "#A"
End Sub

