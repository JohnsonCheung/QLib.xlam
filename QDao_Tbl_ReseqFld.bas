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

Sub ReSeqFldzSpec(D As Database, T, ReSeqSpec$)
ReSeqFldzFny D, T, FnyzReseqSpec(ReSeqSpec)
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

Sub ReSeqFldzFny(D As Database, T, ByFny$())
Dim F, J%
For Each F In ReSeqAy(Fny(D, T), ByFny)
    J = J + 1
    D.TableDefs(T).Fields(F).OrdinalPosition = J
Next
End Sub

Function ReSeqAy(Ay, ByAy)
Dim O
O = IntersectAy(ByAy, Ay)
ReSeqAy = MinusAy(Ay, O)
End Function

