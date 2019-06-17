Attribute VB_Name = "QDta_Dta_Dry"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDta"
Private Const CMod$ = "MDta_Dry."

Function IntAyzIIStr(IIStr$) As Integer()
Dim I
For Each I In Itr(SyzSS(IIStr))
    PushI IntAyzIIStr, I
Next
End Function
Function CntDryWhGt1(CntDry()) As Variant()
Dim Dr
For Each Dr In CntDry
    If Dr(1) > 1 Then PushI CntDryWhGt1, Dr
Next
End Function

Function DrywCoLiny(Dry(), ColIx%, InAy) As Variant()
Const CSub$ = CMod & "DrywCoLiny"
If Not IsArray(InAy) Then Thw CSub, "[InAy] is not Array, but [TypeName]", InAy, TypeName(InAy)
If Si(InAy) = 0 Then DrywCoLiny = Dry: Exit Function
Dim Dr
For Each Dr In Itr(Dry)
    If HasEle(InAy, Dr(ColIx)) Then PushI DrywCoLiny, Dr
Next
End Function
Function DotSyzDry(Dry()) As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DotSyzDry, JnDot(Dr)
Next
End Function

Private Sub Z_FmtA()
DmpAy FmtDry(SampDry1)
End Sub

Function DrywDup(Dry(), C&) As Variant()
DrywDup = AywIxy(Dry, IxyzDup(ColzDry(Dry, C)))
End Function
Private Sub Z_DryzJnFldKK()
Dim Dry(), KKIxy&(), JnFldIx&, Sep$
Sep = " "
Dry = Array(Array(1, 2, 3, 4, "Dry"), Array(1, 2, 3, 6, "B"), Array(1, 2, 2, 8, "C"), Array(1, 2, 2, 12, "DD"), _
Array(2, 3, 1, 1, "x"), Array(12, 3), Array(12, 3, 1, 2, "XX"))
Ept = Array()
KKIxy = Array(0, 1, 2)
JnFldIx = 4
GoSub Tst
Exit Sub
Tst:
    Act = DryzJnFldKK(Dry, KKIxy, JnFldIx, Sep)
    BrwDry CvAv(Act)
    StopNE
    Return
End Sub

Function DryzJnFldKK(Dry(), KKIxy&(), JnFldIx&, Optional Sep$ = " ") As Variant()
Dim Ixy&(): Ixy = KKIxy: PushI Ixy, JnFldIx
Dim N%: N = Si(Ixy)
DryzJnFldKK = DryJnFldNFld(SelCol(Dry, Ixy), N)
End Function

Function RowIxOptzDryDr&(Dry(), Dr)
Dim N%: N = Si(Dr)
Dim Ix&, D
For Each D In Itr(Dry)
    If IsEqAy(FstNEle(D, N), Dr) Then
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
    Dim Ix: Ix = RowIxOptzDryDr(O, FstNEle(Dr, UK))
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
    PushI DryzSslAy, SyzSS(Ssl)
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
Function KeepFstNColzDrs(A As Drs, N%) As Drs
KeepFstNColzDrs = Drs(CvSy(FstNEle(A.Fny, N)), KeepFstNCol(A.Dry, N))
End Function

Function KeepFstNCol(Dry(), N%) As Variant()
Dim Dr, U%
U = N - 1
For Each Dr In Itr(Dry)
    ReDim Preserve Dr(U)
    PushI KeepFstNCol, Dr
Next
End Function
Function DrywColPfx(Dry(), C&, Pfx, Optional Cmp As VbCompareMethod = vbTextCompare) As Variant()
Dim Dr: For Each Dr In Itr(Dry)
   If HasPfx(Dr(C), Pfx, Cmp) Then PushI DrywColPfx, Dr
Next
End Function
Function AlignzTRst(Ly$()) As String()
Dim TAy$(), RstAy$()
Dim L, T$, Rst$: For Each L In Itr(Ly)
    AsgTRst L, T, Rst
    PushI TAy, T
    PushI RstAy, Rst
Next
TAy = AlignLzAy(TAy)
Dim J&: For J = 0 To UB(TAy)
    PushI AlignzTRst, TAy(J) & " " & RstAy(J)
Next
End Function
Function AlignRzDryC(Dry(), C) As Variant()
Dim Ay$(): Ay = AlignRzAy(ColzDry(Dry, C))
Dim O(): O = Dry
Dim J&
For J = 0 To UB(O)
    O(J)(C) = Ay(J)
Next
AlignRzDryC = O
End Function
Sub ThwIf_NEDry(Dry(), B())
If Not IsEqDry(Dry, B) Then Stop
End Sub
Function DrywDist(Dry()) As Variant()
Dim O(), Dr
For Each Dr In Itr(Dry)
    PushNoDupDr O, Dr
Next
DrywDist = O
End Function

Function DrywCCNe(Dry(), C1&, C2&) As Variant()
Dim Dr
For Each Dr In Dry
    If Dr(C1) <> Dr(C2) Then PushI DrywCCNe, Dr
Next
End Function

Function DrywColGt(Dry(), C%, GtV) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    If Dr(C) > GtV Then PushI DrywColGt, Dr
Next
End Function

Function DrywColNe(Dry(), C, Ne) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    If Dr(C) <> Ne Then PushI DrywColNe, Dr
Next
End Function

Function DrywColIn(Dry(), C, InVy) As Variant()
If Not IsArray(InVy) Then Thw CSub, "Given InVy is not an array", "Ty-InVy", TypeName(InVy)
Dim Dr
For Each Dr In Itr(Dry)
    If HasEle(InVy, Dr(C)) Then
        PushI DrywColIn, Dr
    End If
Next
End Function

Function DrywColEq(Dry(), C&, Eq) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    If Dr(C) = Eq Then PushI DrywColEq, Dr
Next
End Function
Function HasColEqzDry(Dry(), C&, Eq) As Boolean
Dim Dr
For Each Dr In Itr(Dry)
    If Dr(C) = Eq Then HasColEqzDry = True: Exit Function
Next
End Function
Function FstRecEqzDry(Dry(), C, Eq, SelIxy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    If Dr(C) = Eq Then FstRecEqzDry = Array(AywIxy(Dr, SelIxy)): Exit Function
Next
Thw CSub, "No first rec in Dry of Col-A eq to Val-B", "Col-A Val-B Dry", C, Eq, FmtDry(Dry)
End Function
Function FstDrEqzDry(Dry(), C, Eq, SelIxy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    If Dr(C) = Eq Then FstDrEqzDry = AywIxy(Dr, SelIxy): Exit Function
Next
Thw CSub, "No first Dr in Dry of Col-A eq to Val-B", "Col-A Val-B Dry", C, Eq, FmtDry(Dry)
End Function

Function DrywDupCC(Dry(), CCIxy&()) As Variant()
Dim Dup$(), Dr
Dup = AywDup(LyzDryCC(Dry, CCIxy))
For Each Dr In Itr(Dry)
    If HasEle(Dup, Jn(AywIxy(Dr, CCIxy), vbFldSep)) Then Push DrywDupCC, Dr
Next
End Function

Private Function DrywDup1(Dry(), C&) As Variant()
Dim Dup$(), Dr, O()
Dup = CvSy(AywDup(StrColzDry(Dry, C)))
For Each Dr In Itr(Dry)
    If HasEle(Dup, Dr(C)) Then PushI DrywDup1, Dr
Next
End Function

Function DrywIxyzy(Dry(), Ixy&(), EqVy()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    If IsEqAy(AywIxy(Drv, Ixy), EqVy) Then PushI DrywIxyzy, Drv
Next
End Function
Function DistColzDry(Dry(), C&) As Variant()
DistColzDry = AywDist(ColzDry(Dry, C))
End Function

Function DistCol(A As Drs, C$)
DistCol = AywDist(ColzDry(A.Dry, IxzAy(A.Fny, C)))
End Function

Function DistColzStr(A As Drs, C$) As String()
Dim I%: I = IxzAy(A.Fny, C)
Dim Col$(): Col = StrColzDry(A.Dry, I)
DistColzStr = AywDist(Col)
End Function

Function DryzSqCol(Sq(), ColIxy) As Variant()
Dim R&
For R = 1 To UBound(Sq(), 1)
    PushI DryzSqCol, DrzSqr(Sq(), R)
Next
End Function

Function DryzSq(Sq()) As Variant()
If Si(Sq) = 0 Then Exit Function
Dim R&
For R = 1 To UBound(Sq(), 1)
    PushI DryzSq, DrzSqr(Sq, R)
Next
End Function

Function HasIxyDr(Dry(), Ixy&(), Dr) As Boolean
Dim IDr, A()
For Each IDr In Itr(Dry)
    A = AywIxy(Dr, Ixy)
    If IsEqAy(A, Dr) Then HasIxyDr = True: Exit Function
Next
End Function
Function HasDr(Dry(), Dr) As Boolean
Dim IDr
For Each IDr In Itr(Dry)
    If IsEqAy(IDr, Dr) Then HasDr = True: Exit Function
Next
End Function

