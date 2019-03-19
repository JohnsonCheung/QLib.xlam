Attribute VB_Name = "MDao_Tbl_ReseqFld"
Option Explicit

Sub ReSeqFldzSpec(A As Database, T, ReSeqSpec$)
ReSeqFldzFny A, T, FnyzReseqSpec(ReSeqSpec)
End Sub

Private Sub ZZ_FnyzReseqSpec()
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
    Ay = SplitVbar(ReSeqSpec)
    L1 = AyShf(Ay)
    Set D = Dic(Ay)
Dim L1Itm, O$()
For Each L1Itm In SySsl(L1)
    If FstChr(L1Itm) = "*" Then
        PushIAy O, SySsl(D(L1Itm))
    Else
        Push O, L1Itm
    End If
Next
FnyzReseqSpec = O
End Function

Sub ReSeqFldzFny(A As Database, T, ByFny$())
Dim F, J%
For Each F In AyReSeq(Fny(A, T), ByFny)
    J = J + 1
    A.TableDefs(T).Fields(F).OrdinalPosition = J
Next
End Sub

Function AyReSeq(Ay, ByAy)
Dim O
O = AyIntersect(ByAy, Ay)
AyReSeq = AyMinus(Ay, O)
End Function

