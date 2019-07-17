Attribute VB_Name = "QDta_B_Dry"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDta"
Private Const CMod$ = "MDta_Dy."

Function CntDyWhGt1(CntDy()) As Variant()
Dim Dr
For Each Dr In CntDy
    If Dr(1) > 1 Then PushI CntDyWhGt1, Dr
Next
End Function

Function DywCoLiny(Dy(), ColIx%, InAy) As Variant()
Const CSub$ = CMod & "DywCoLiny"
If Not IsArray(InAy) Then Thw CSub, "[InAy] is not Array, but [TypeName]", InAy, TypeName(InAy)
If Si(InAy) = 0 Then DywCoLiny = Dy: Exit Function
Dim Dr
For Each Dr In Itr(Dy)
    If HasEle(InAy, Dr(ColIx)) Then PushI DywCoLiny, Dr
Next
End Function
Function DotSyzDy(Dy()) As String()
Dim Dr
For Each Dr In Itr(Dy)
    PushI DotSyzDy, JnDot(Dr)
Next
End Function

Private Sub Z_FmtA()
DmpAy JnDy(SampDy1)
End Sub

Function DywDup(Dy(), C&) As Variant()
DywDup = AwIxy(Dy, IxyzDup(ColzDy(Dy, C)))
End Function
Private Sub Z_DyoJnFldKK()
Dim Dy(), KKIxy&(), JnFldIx&, Sep$
Sep = " "
Dy = Array(Array(1, 2, 3, 4, "Dy"), Array(1, 2, 3, 6, "B"), Array(1, 2, 2, 8, "C"), Array(1, 2, 2, 12, "DD"), _
Array(2, 3, 1, 1, "x"), Array(12, 3), Array(12, 3, 1, 2, "XX"))
Ept = Array()
KKIxy = Array(0, 1, 2)
JnFldIx = 4
GoSub Tst
Exit Sub
Tst:
    Act = DyoJnFldKK(Dy, KKIxy, JnFldIx, Sep)
    BrwDy CvAv(Act)
    StopNE
    Return
End Sub

Function DyoJnFldKK(Dy(), KKIxy&(), JnFldIx&, Optional Sep$ = " ") As Variant()
Dim Ixy&(): Ixy = KKIxy: PushI Ixy, JnFldIx
Dim N%: N = Si(Ixy)
DyoJnFldKK = DyJnFldNFld(SelCol(Dy, Ixy), N)
End Function

Function RowIxOptzDyDr&(Dy(), Dr)
Dim N%: N = Si(Dr)
Dim Ix&, D
For Each D In Itr(Dy)
    If IsEqAy(FstNEle(D, N), Dr) Then
        RowIxOptzDyDr = Ix
        Exit Function
    End If
    Ix = Ix + 1
Next
RowIxOptzDyDr = -1
End Function
Function DyJnFldNFld(Dy(), FstNFld%, Optional Sep$ = " ") As Variant()
Dim U%: U = FstNFld - 1
Dim UK%: UK = U - 1
Dim O(), Dr
For Each Dr In Itr(Dy)
    If U <> UB(Dr) Then
        ReDim Preserve Dr(U)
    End If
    Dim Ix: Ix = RowIxOptzDyDr(O, FstNEle(Dr, UK))
    If Ix = -1 Then
        PushI O, Dr
    Else
        Stop
'        O(Ix)(U) = ApdIf(O(Ix)(U), Sep) & Dr(U)
    End If
Next
DyJnFldNFld = O
End Function

Function DyoSslAy(SslAy$()) As Variant()
Dim Ssl$, I
For Each I In Itr(SslAy)
    Ssl = I
    PushI DyoSslAy, SyzSS(Ssl)
Next
End Function

Function DiKqCntzDyC(Dy(), C&) As Dictionary
Set DiKqCntzDyC = DiKqCnt(ColzDy(Dy, C))
End Function

Function SeqDiKqCntzDy(Dy(), C&) As Dictionary
Set SeqDiKqCntzDy = SeqDiKqCnt(ColzDy(Dy, C&))
End Function


Function IsRowBrk(Dy(), R&, BrkColIx&) As Boolean
If Si(Dy) = 0 Then Exit Function
If R& = 0 Then Exit Function
If R = UB(Dy) Then Exit Function
If Dy(R)(BrkColIx) = Dy(R - 1)(BrkColIx) Then Exit Function
IsRowBrk = True
End Function

Function NColzDy%(Dy())
Dim O%, Dr
For Each Dr In Itr(Dy)
    O = Max(O, Si(Dr))
Next
NColzDy = O
End Function

Function NRowzInDyoColEv&(Dy(), C&, Ev)
Dim J&, O&, Dr
For Each Dr In Itr(Dy)
   If Dr(C) = Ev Then O = O + 1
Next
NRowzInDyoColEv = O
End Function

Function KeepFstNColzDrs(A As Drs, N%) As Drs
KeepFstNColzDrs = Drs(CvSy(FstNEle(A.Fny, N)), KeepFstNCol(A.Dy, N))
End Function

Function KeepFstNCol(Dy(), N%) As Variant()
Dim Dr, U%
U = N - 1
For Each Dr In Itr(Dy)
    ReDim Preserve Dr(U)
    PushI KeepFstNCol, Dr
Next
End Function

Function DywPfx(Dy(), C&, Pfx, Optional Cmp As VbCompareMethod = vbTextCompare) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
   If HasPfx(Dr(C), Pfx, Cmp) Then PushI DywPfx, Dr
Next
End Function

Function AlignTRst(Ly$()) As String()
Dim TAy$(), RstAy$()
Dim L, T$, Rst$: For Each L In Itr(Ly)
    AsgTRst L, T, Rst
    PushI TAy, T
    PushI RstAy, Rst
Next
TAy = AlignAy(TAy)
Dim J&: For J = 0 To UB(TAy)
    PushI AlignTRst, TAy(J) & " " & RstAy(J)
Next
End Function
Function AlignRzDyC(Dy(), C) As Variant()
Dim Ay$(): Ay = AlignRzAy(ColzDy(Dy, C))
Dim O(): O = Dy
Dim J&
For J = 0 To UB(O)
    O(J)(C) = Ay(J)
Next
AlignRzDyC = O
End Function
Sub ThwIf_NEDy(Dy(), B())
If Not IsEqDy(Dy, B) Then Stop
End Sub

Function DywDist(Dy()) As Variant()
Dim O(), Dr
For Each Dr In Itr(Dy)
    PushNDupDr O, Dr
Next
DywDist = O
End Function

Function DywCCNe(Dy(), C1&, C2&) As Variant()
Dim Dr
For Each Dr In Dy
    If Dr(C1) <> Dr(C2) Then PushI DywCCNe, Dr
Next
End Function

Function DywColGt(Dy(), C%, GtV) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) > GtV Then PushI DywColGt, Dr
Next
End Function

Function DywColNe(Dy(), C, Ne) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) <> Ne Then PushI DywColNe, Dr
Next
End Function

Function DywIn(Dy(), C, InVy) As Variant()
If Not IsArray(InVy) Then Thw CSub, "Given InVy is not an array", "Ty-InVy", TypeName(InVy)
Dim Dr
For Each Dr In Itr(Dy)
    If HasEle(InVy, Dr(C)) Then
        PushI DywIn, Dr
    End If
Next
End Function

Function DywEq(Dy(), C&, Eq) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then PushI DywEq, Dr
Next
End Function

Function HasColEqzDy(Dy(), C&, Eq) As Boolean
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then HasColEqzDy = True: Exit Function
Next
End Function

Function FstRecEqzDy(Dy(), C, Eq, SelIxy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then FstRecEqzDy = Array(AwIxy(Dr, SelIxy)): Exit Function
Next
Thw CSub, "No first rec in Dy of Col-A eq to Val-B", "Col-A Val-B Dy", C, Eq, JnDy(Dy)
End Function

Function FstDrzDy(Dy(), C, V) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) = V Then FstDrzDy = Dr: Exit Function
Next
Thw CSub, "No first Dr in Dy of Cix eq to V", "Cix V Dy", C, V, JnDy(Dy)
End Function

Function DywDupCC(Dy(), CCIxy&()) As Variant()
Dim Dup$(), Dr
Dup = AwDup(JnDyCC(Dy, CCIxy))
For Each Dr In Itr(Dy)
    If HasEle(Dup, Jn(AwIxy(Dr, CCIxy), vbFldSep)) Then Push DywDupCC, Dr
Next
End Function

Private Function DywDup1(Dy(), C&) As Variant()
Dim Dup$(), Dr, O()
Dup = CvSy(AwDup(StrColzDy(Dy, C)))
For Each Dr In Itr(Dy)
    If HasEle(Dup, Dr(C)) Then PushI DywDup1, Dr
Next
End Function

Function DywEqVy(Dy(), Ixy&(), Vy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If IsEqAy(AwIxy(Dr, Ixy), Vy) Then PushI DywEqVy, Dr
Next
End Function
Function DistColzDy(Dy(), C&) As Variant()
DistColzDy = AwDist(ColzDy(Dy, C))
End Function

Function DistCol(A As Drs, C$)
DistCol = AwDist(ColzDy(A.Dy, IxzAy(A.Fny, C)))
End Function

Function DistColoS(A As Drs, C$) As String()
Dim I%: I = IxzAy(A.Fny, C)
Dim Col$(): Col = StrColzDy(A.Dy, I)
DistColoS = AwDist(Col)
End Function
Function IxyzCny(Cny) As Long()
If Si(Cny) = 0 Then Exit Function
Dim Cno: For Each Cno In Cny
    PushI IxyzCny, Cno - 1
Next
End Function

Function DyzVbl(Vbl$) As Variant()
Dim L: For Each L In Itr(SplitVBar(Vbl))
    PushI DyzVbl, SyzSS(L)
Next
End Function

Function DyzSqc(Sq(), Cny) As Variant()
'Fm Cny : :Cny ! selecting which col of @Sq
Dim Ixy&(): Ixy = IxyzCny(Cny)
Dim R&: For R = 1 To UBound(Sq(), 1)
    PushI DyzSqc, AwIxy(DrzSqr(Sq(), R), Ixy)
Next
End Function

Function DyoSq(Sq()) As Variant()
If Si(Sq) = 0 Then Exit Function
Dim R&: For R = 1 To UBound(Sq(), 1)
    PushI DyoSq, DrzSqr(Sq, R)
Next
End Function

Function HasDrzIxy(Dy(), Dr, Ixy&()) As Boolean
Dim IDr, A()
For Each IDr In Itr(Dy)
    A = AwIxy(IDr, Ixy)
    If IsEqAy(A, IDr) Then HasDrzIxy = True: Exit Function
Next
End Function

Function HasDr(Dy(), Dr) As Boolean
Dim IDr
For Each IDr In Itr(Dy)
    If IsEqAy(IDr, Dr) Then HasDr = True: Exit Function
Next
End Function

