Attribute VB_Name = "QDta_Dta_Dry"
Option Explicit
Private Const Asm$ = "QDta"
Private Const CMod$ = "MDta_Dry."

Function IntAyzIIStr(IIStr$) As Integer()
Dim I
For Each I In Itr(SyzSsLin(IIStr))
    PushI IntAyzIIStr, I
Next
End Function
Function CntDryWhGt1(CntDry()) As Variant()
Dim Dr
For Each Dr In CntDry
    If Dr(1) > 1 Then PushI CntDryWhGt1, Dr
Next
End Function

Function DrywColInAy(Dry(), ColIx%, InAy) As Variant()
Const CSub$ = CMod & "DrywColInAy"
If Not IsArray(InAy) Then Thw CSub, "[InAy] is not Array, but [TypeName]", InAy, TypeName(InAy)
If Si(InAy) = 0 Then DrywColInAy = Dry: Exit Function
Dim Dr
For Each Dr In Itr(Dry)
    If HasEle(InAy, Dr(ColIx)) Then PushI DrywColInAy, Dr
Next
End Function
Function DotSyzDry(Dry()) As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DotSyzDry, JnDot(Dr)
Next
End Function
Function DryzDotSy(DotNy$()) As Variant()
If Si(DotSy) = 0 Then Exit Function
Dim O(), I, S$
For Each I In DotSy
    S = I
    With Brk1(S, ".")
       Push O, Sy(.S1, .S2)
   End With
Next
DryzDotSy = O
End Function

Private Sub ZZ_FmtA()
DmpAy FmtDry(SampDry1)
End Sub

Function DrywDup(Dry(), C&) As Variant()
DrywDup = AywIxAy(Dry, IxAyzDup(ColzDry(Dry, C)))
End Function
Private Sub Z_DryzJnFldKK()
Dim Dry(), KKIxAy&(), JnFldIx&, Sep$
Sep = " "
Dry = Array(Array(1, 2, 3, 4, "Dry"), Array(1, 2, 3, 6, "B"), Array(1, 2, 2, 8, "C"), Array(1, 2, 2, 12, "DD"), _
Array(2, 3, 1, 1, "x"), Array(12, 3), Array(12, 3, 1, 2, "XX"))
Ept = Array()
KKIxAy = Array(0, 1, 2)
JnFldIx = 4
GoSub Tst
Exit Sub
Tst:
    Act = DryzJnFldKK(Dry, KKIxAy, JnFldIx, Sep)
    BrwDry CvAv(Act)
    StopNE
    Return
End Sub

Function DryzJnFldKK(Dry(), KKIxAy&(), JnFldIx&, Optional Sep$ = " ") As Variant()
Dim IxAy&(): IxAy = KKIxAy: PushI IxAy, JnFldIx
Dim N%: N = Si(IxAy)
DryzJnFldKK = DryJnFldNFld(SelCol(Dry, IxAy), N)
End Function

Function RowIxOptzDryDr&(Dry(), Dr)
Dim N%: N = Si(Dr)
Dim Ix&, D
For Each D In Itr(Dry)
    If IsEqAy(AywFstNEle(D, N), Dr) Then
        RowIxOptzDryDr = Ix
        Exit Function
    End If
    Ix = Ix + 1
Next
RowIxOptzDryDr = -1
End Function
Function DryJnFldNFld(Dry(), FstNFld%, Optional Sep$ = " ") As Variant()
Dim U%: U = FstNFld - 1
Dim UK%: UK = U - 1
Dim O(), Dr
For Each Dr In Itr(Dry)
    If U <> UB(Dr) Then
        ReDim Preserve Dr(U)
    End If
    Dim Ix: Ix = RowIxOptzDryDr(O, AywFstNEle(Dr, UK))
    If Ix = -1 Then
        PushI O, Dr
    Else
        Stop
'        O(Ix)(U) = ApdIf(O(Ix)(U), Sep) & Dr(U)
    End If
Next
DryJnFldNFld = O
End Function

Function DryzSslAy(SslAy$()) As Variant()
Dim Ssl$, I
For Each I In Itr(SslAy)
    Ssl = I
    PushI DryzSslAy, SyzSsLin(Ssl)
Next
End Function

Function CntDiczDryC(Dry(), C&) As Dictionary
Set CntDiczDryC = CntDic(ColzDry(Dry, C))
End Function

Function SeqCntDiczDry(Dry(), C&) As Dictionary
Set SeqCntDiczDry = SeqCntDic(ColzDry(Dry, C&))
End Function


Function IsRowBrk(Dry(), R&, BrkColIx&) As Boolean
If Si(Dry) = 0 Then Exit Function
If R& = 0 Then Exit Function
If R = UB(Dry) Then Exit Function
If Dry(R)(BrkColIx) = Dry(R - 1)(BrkColIx) Then Exit Function
IsRowBrk = True
End Function

Function NColzDry%(Dry())
Dim O%, Dr
For Each Dr In Itr(Dry)
    O = Max(O, Si(Dr))
Next
NColzDry = O
End Function

Function NRowzInDryzColEv&(Dry(), C&, Ev)
Dim J&, O&, Dr
For Each Dr In Itr(Dry)
   If Dr(C) = Ev Then O = O + 1
Next
NRowzInDryzColEv = O
End Function

Function DrywCEv(Dry(), C&, Ev) As Variant()
Dim O()
Dim Dr
For Each Dr In Itr(Dry)
   If Dr(C) = Ev Then PushI DrywCEv, Dr
Next
End Function

Sub ThwIfNEDry(Dry(), B())
If Not IsEqDry(Dry, B) Then Stop
End Sub

Function DrywCCNe(Dry(), C1&, C2&) As Variant()
Dim Dr
For Each Dr In Dry
    If Dr(C1) <> Dr(C2) Then PushI DrywCCNe, Dr
Next
End Function

Function DrywColEq(Dry(), C%, V) As Variant()
Dim Dr
For Each Dr In Dry
    If Dr(C) = V Then PushI DrywColEq, Dr
Next
End Function

Function DrywCGt(Dry(), C%, GtV) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    If Dr(C) > GtV Then PushI DrywCGt, Dr
Next
End Function

Function DrywDupCC(Dry(), CCIxAy&()) As Variant()
Dim Dup$(), Dr
Dup = AywDup(LyzDryCC(Dry, CCIxAy))
For Each Dr In Itr(Dry)
    If HasEle(Dup, Jn(AywIxAy(Dr, CCIxAy), vbFldSep)) Then Push DrywDupCC, Dr
Next
End Function

Private Function DrywDup1(Dry(), C&) As Variant()
Dim Dup$(), Dr, O()
Dup = CvSy(AywDup(StrColzDry(Dry, C)))
For Each Dr In Itr(Dry)
    If HasEle(Dup, Dr(C)) Then PushI DrywDup1, Dr
Next
End Function

Function DrywIxAyzy(Dry(), IxAy&(), EqVy()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    If IsEqAy(AywIxAy(Drv, IxAy), EqVy) Then PushI DrywIxAyzy, Drv
Next
End Function
Function ColzDryC(Dry(), C&) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    If UB(Drv) < C Then
        PushI ColzDryC, Empty
    Else
        PushI ColzDryC, Drv(C)
    End If
Next
End Function
Function DistColzDryC(Dry(), C&) As Variant()
DistColzDryC = AywDist(ColzDryC(Dry, C))
End Function

Function DryzSqCol(Sq(), ColIxAy) As Variant()
Dim R&
For R = 1 To UBound(Sq(), 1)
    PushI DryzSqCol, DrzSqr(Sq(), R)
Next
End Function

Function DryzSq(Sq()) As Variant()
Dim R&
For R = 1 To UBound(Sq(), 1)
    PushI DryzSq, DrzSqr(Sq, R)
Next
End Function


Function HasDr(Dry(), Dr) As Boolean
Dim IDr
For Each IDr In Itr(Dry)
    If IsEqAy(IDr, Dr) Then HasDr = True: Exit Function
Next
End Function

