Attribute VB_Name = "QDao_Tbl_ReseqFld"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Tbl_ReseqFld."
Private Const Asm$ = "QDao"

Public Const SampReSeqLin$ = _
"Flg RecTy Amt Key Uom MovTy Qty BchRateUX RateTy Bch Las GL |" & _
" Flg IsAlert IsWithSku |" & _
" Key Sku PstMth PstDte |" & _
" Bch BchNo BchPermitDate BchPermit |" & _
" Las LasBchNo LasPermitDate LasPermit |" & _
" GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
" Uom Des StkUom Ac_U"
Public Const LgSchmNm$ = "LgSchm" ' The LgSchm-Spnm

Function FnyzReseqSpec(ReSeqSpec$) As String()
Dim L1$
Dim D As Dictionary
    Dim Ay$()
    Ay = SplitVBar(ReSeqSpec)
    L1 = Shf(Ay)
    Set D = Dic(Ay)
Dim L1Itm$, I, O$()
For Each I In SyzSS(L1)
    L1Itm = I
    If FstChr(L1Itm) = "*" Then
        PushIAy O, SyzSS(D(L1Itm))
    Else
        PushI O, L1Itm
    End If
Next
FnyzReseqSpec = O
End Function

Function ReSeqAy(Ay, ByAy)
Dim O
O = AyIntersect(ByAy, Ay)
ReSeqAy = MinusAy(Ay, O)
End Function

Sub ReSeqFldzFny(D As Database, T, ByFny$())
Dim F, J%
For Each F In ReSeqAy(Fny(D, T), ByFny)
    J = J + 1
    D.TableDefs(T).Fields(F).OrdinalPosition = J
Next
End Sub

Sub ReSeqFldzSpec(D As Database, T, ReSeqSpec$)
ReSeqFldzFny D, T, FnyzReseqSpec(ReSeqSpec)
End Sub

Sub UpdSeqFld(D As Database, T, SeqFld$, GpFF$, OrdffMinus$)
Dim Q$: Q = SqlSel_FF_T_Ordff(SeqFld & " " & GpFF, T, OrdffMinus)
Dim R As Recordset: Set R = Rs(D, Q)
If NoRec(R) Then Exit Sub
Dim Seq&, Las(), Cur(), N%
With R
    N = .Fields.Count - 1
    .MoveNext
    Las = DrzRs(R)
    While Not .EOF
        Stop
'        Cur = DrzSqR(R, N)
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

Private Sub Y_FnyzReseqSpec()
BrwAy FnyzReseqSpec("*Flg RecTy Amt *Key *Uom MovTy Qty BchRateUX RateTy *Bch *Las *GL |" & _
" *Flg IsAlert IsWithSku |" & _
" *Key Sku PstMth PstDte |" & _
" *Bch BchNo BchPermitDate BchPermit |" & _
" *Las LasBchNo LasPermitDate LasPermit |" & _
" *GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
" *Uom Des StkUom Ac_U")
End Sub

Private Sub Y_UpdSeqFld()
Dim Db As Database, T$
Set Db = TmpDb
Rq Db, "Select * into [#A] from [T] order by Sku,PermitDate"
Rq Db, "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
UpdSeqFld Db, T, "BchRateSeq", "Sku", "Sku Rate"
Stop
DrpT Db, "#A"
End Sub

