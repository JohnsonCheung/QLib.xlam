Attribute VB_Name = "MDao_Tbl_Upd_SeqFld"
Option Explicit

Sub UpdSeqFldzRs(Rs As DAO.Recordset, SelFld$, KK)
Rs.MoveFirst
Dim Fny$(): Fny = FnyzFF(KK)
Dim Las(): 'Las = DrzFnyzRs(Rs, Fny)
While Not Rs.EOF
'    Cur = DrzFnyzRs(Rs, Fny)
    Rs.Edit
'    Rs.Fields(SelFld).Value = Seq
    Rs.Update
'    If IsEqAy(Cur, Las) Then
'        Seq = Seq + 1
'    Else
'        Las = Cur
'        Seq = 0
'    End If
    Rs.MoveNext
Wend
End Sub


Sub UpdSeqFldvRs(Rs As DAO.Recordset)
If NoRec(Rs) Then Exit Sub
Dim Seq&, Las(), Cur(), N%
With Rs
    N = .Fields.Count - 1
    .MoveNext
    Las = DrzRs(Rs)
    While Not .EOF
        Cur = DrzSqr(Rs, N)
        If Not IsEqAy(Cur, Las) Then
            Cur = Las
            Seq = 0
        End If
        Seq = Seq + 1
        UpdRs Rs, Array(Seq)
        .MoveNext
    Wend
End With
End Sub
Function SqlvUpdSeqFld$(T, SeqFld$, GpFF, OrdFFMinus$)
SqlvUpdSeqFld = SqlSel_FF_Fm_Ord(SeqFld & " " & GpFF, T, OrdFFMinus)
End Function
Sub UpdSeqFld(T, SeqFld$, GpFF, OrdFFMinus$)
UpdSeqFldz CDb, T, SeqFld, GpFF, OrdFFMinus
End Sub
Sub UpdSeqFldz(Db As Database, T, SeqFld$, GpFF, OrdFFMinus$)
UpdSeqFldvRs Rsz(Db, SqlvUpdSeqFld(T, SeqFld, GpFF, OrdFFMinus))
End Sub


Private Sub ZZ_UpdSeqFld()
Dim Db As Database, T
RunQ "Select * into [#A] from ZZ_UpdSeqFld order by Sku,PermitDate"
RunQ "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
UpdSeqFld T, "BchRateSeq", "Sku", "Sku Rate"
Stop
Drp "#A"
End Sub

