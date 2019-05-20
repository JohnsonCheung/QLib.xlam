Attribute VB_Name = "QVb_Dta_KeyLin"
Option Explicit
Option Compare Text
Type KL: K As String: Lno As Long: End Type:            Type KLs: N As Integer: Ay() As KL: End Type
Type KLx: KL As KL: Lnxs As Lnxs: End Type:             Type KLxs: N As Long: Ay() As KLx: End Type
Type KLy: K As String: Ly() As String: End Type:        Type KLys: N As Long: Ay() As KLy: End Type
Type KLss: K As String: Lnoss As String: End Type:      Type KLsses: N As Integer: Ay() As KLss: End Type
Type KLssOpt: Som As Boolean: KLss As KLss: End Type
Sub PushKLsses(O As KLsses, M As KLsses)
Dim J%
For J = 0 To M.N - 1
    PushKLss O, M.Ay(J)
Next
End Sub
Function PushKL(O As KLs, M As KL)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Function

Function KLy(K$, Ly$()) As KLy
KLy.K = K
KLy.Ly = Ly
End Function

Sub PushKLss(O As KLsses, M As KLss)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function KLss(K$, Lnoss$) As KLss
KLss.K = K
KLss.Lnoss = Lnoss
End Function

Function SomKLss(KLss As KLss) As KLssOpt
SomKLss.Som = True
SomKLss.KLss = KLss
End Function

Function RmvEle_ForLTrimPfx_Is2Hypron_FmLnxs(A As Lnxs) As Lnxs
Dim J%
For J = 0 To A.N - 1
    If Not HasPfx(LTrim(A.Ay(J).Lin), "--") Then
        PushLnx RmvEle_ForLTrimPfx_Is2Hypron_FmLnxs, A.Ay(J)
    End If
Next
End Function

Function RmvEle_ForLTrimPfx_Is2Hypron(Ly$()) As String()
Dim L
For Each L In Itr(Ly)
    If Not HasPfx(LTrim(L), "--") Then
        PushI RmvEle_ForLTrimPfx_Is2Hypron, L
    End If
Next
End Function
Function RplEle_OfNonBlankFstChr_ByT1(Ly$()) As String()
Dim J%, L
For Each L In Itr(Ly)
    If FstChr(L) = " " Then
        PushI RplEle_OfNonBlankFstChr_ByT1, L
    Else
        PushI RplEle_OfNonBlankFstChr_ByT1, T1(L)
    End If
Next
End Function

Function KLys_ForTakT1_AsK(Ly$()) As KLys
Dim A$(): A = RplEle_OfNonBlankFstChr_ByT1(Ly)
KLys_ForTakT1_AsK = KLys(A)
End Function

Function HasKPfx_InKLys(KPfx$, A As KLys) As Boolean
Dim J%
For J = 0 To A.N - 1
    If HasPfx(A.Ay(J).K, KPfx) Then HasKPfx_InKLys = True: Exit Function
Next
End Function

Function KLyseFstNEle(A As KLys, Optional N& = 1) As KLys
Dim J&
For J = N To A.N - 1
    PushKLy KLyseFstNEle, A.Ay(J)
Next
End Function

Function KLys(Ly$()) As KLys
Dim L$(): L = RmvEle_ForLTrimPfx_Is2Hypron(Ly)
If Si(L) = 0 Then Exit Function
If FstChr(L(0)) = " " Then Thw CSub, "FstLin fstChr must be non blank", "L", L
Dim M As KLy
Dim I
For Each I In L
    If FstChr(I) = " " Then
        PushI M.Ly, Trim(I)
    Else
        If Not IsEmpKLy(M) Then PushKLy KLys, M: M = EmpKLy
        M = KLy(BefOrAll(I, " "), EmpSy)
    End If
Next
If Not IsEmpKLy(M) Then PushKLy KLys, M: M = EmpKLy
End Function
Function IsEmpKLx(A As KLx) As Boolean
With A
    If Not IsEmpKL(A.KL) Then Exit Function
    If Not IsEmpLnxs(A.Lnxs) Then Exit Function
End With
IsEmpKLx = True
End Function

Function IsEmpKL(A As KL) As Boolean
If A.K <> "" Then Exit Function
If A.Lno > 0 Then Exit Function
IsEmpKL = True
End Function

Function IsEmpLnxs(A As Lnxs) As Boolean
IsEmpLnxs = A.N <= 0
End Function
Function IsEmpKLy(A As KLy) As Boolean
If A.K <> "" Then Exit Function
If Si(A.Ly) <> 0 Then Exit Function
IsEmpKLy = True
End Function

Function EmpKLx() As KLx
End Function

Function EmpKLy() As KLy
End Function
Private Sub ZZ_FmtKLys()
GoSub ZZ
Exit Sub
ZZ:
    BrwKLys KLys(Src(CMd))
    Return
End Sub
Sub BrwKLys(A As KLys)
B FmtKLys(A)
End Sub
Function FmtKLys(A As KLys) As String()
Dim J&
For J = 0 To A.N - 1
    PushIAy FmtKLys, FmtKLy(A.Ay(J))
Next
End Function
Function FmtKLy(A As KLy) As String()
Dim K$, Ly$(): K = A.K: Ly = A.Ly
Select Case Si(Ly)
Case 0: PushI FmtKLy, K
Case 1: PushI FmtKLy, K & " " & Ly(0)
Case Else
    PushI FmtKLy, K & " " & Ly(0)
    Dim S$: S = Space(Len(K) + 1)
    Dim J&
    For J = 1 To UB(Ly)
        PushI FmtKLy, S & Ly(J)
    Next
End Select
End Function

Function ShfOptLy_FmKLys(O As KLys, K$) As String() 'Shift-KLy-Opt-Ly
If O.N = 0 Then Exit Function
If O.Ay(0).K <> K Then Exit Function
ShfOptLy_FmKLys = O.Ay(0).Ly
O = KLyseFstNEle(O)
End Function

Function ShfOptLin_FmKLys$(O As KLys, K$) ' Shift-KLy-Must-Lin
ShfOptLin_FmKLys = JnSpc(ShfOptLy_FmKLys(O, K))
End Function

Function ShfMusLy_FmKLys_ByKK(O As KLys, KK$) As String()
End Function

Function ShfMusLin_FmKLys$(O As KLys, K$) ' Shift-KLy-Must-Lin
ShfMusLin_FmKLys = JnSpc(ShfMusLy_FmKLys(O, K))
End Function

Function ShfMusLy_FmKLys(O As KLys, K$) As String() ' Shift-KLy-Must-Ly
Dim X As KLys: X = O
Dim Ly$(): Ly = ShfOptLy_FmKLys(O, K)
If Si(Ly) Then Thw CSub, "Fst ele of given KLys is expected to have given K and have non-0-Ly.", "Given-K Given-KLys", K, FmtKLys(O)
ShfMusLy_FmKLys = Ly
End Function

Function RmvFstElezKLys(A As KLys) As KLys
Dim J%
For J = 1 To A.N - 1
    PushKLy RmvFstElezKLys, A.Ay(J)
Next
End Function

Sub PushKLy(O As KLys, M As KLy)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function KLs_FmKLxs_WhNoLin(A As KLxs) As KLs
Dim J%
For J = 0 To A.N - 1
    With A.Ay(J)
        If .Lnxs.N = 0 Then
            PushKL KLs_FmKLxs_WhNoLin, .KL
        End If
    End With
Next
End Function
Function KLzLnx(A As Lnx) As KL
Dim K$:       K = BefOrAll(A.Lin, " ")
Dim Lno&:   Lno = A.Ix + 1
         KLzLnx = KL(K, Lno)
End Function
Function KL(K, Lno) As KL
KL.K = K
KL.Lno = Lno
End Function
Function KLx(KL As KL, Lnxs As Lnxs) As KLx
KLx.KL = KL
KLx.Lnxs = Lnxs
End Function

Private Sub PushKLxs(O As KLxs, M As KLxs)
Dim J&
For J = 0 To M.N - 1
    PushKLx O, M.Ay(J)
Next
End Sub

Private Sub PushKLx(O As KLxs, M As KLx)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function RmvKeyPfxzKLxs(A As KLxs, Pfx$, Optional C As VbCompareMethod = vbTextCompare) As KLxs
Dim J%, O As KLxs
O = A
For J = 0 To O.N - 1
    O.Ay(J).KL.K = RmvPfx(O.Ay(J).KL.K, Pfx, C)
Next
RmvKeyPfxzKLxs = O
End Function

Function KLxs_WhPfx(A As KLxs, Pfx$, Optional C As VbCompareMethod = vbTextCompare) As KLxs
Dim J%, M As KLx
For J = 0 To A.N - 1
    M = A.Ay(J)
    If HasPfx(M.KL.K, Pfx, C) Then
        PushKLx KLxs_WhPfx, M
    End If
Next
End Function
Function LnoAyzKLsK(A As KLs, K) As Long()
Dim J%
For J = 0 To A.N - 1
    With A.Ay(J)
        If .K = K Then
            PushI LnoAyzKLsK, .Lno
        End If
    End With
Next
End Function
Function KyzKLs(A As KLs) As String()
Dim J%
For J = 0 To A.N - 1
    PushI KyzKLs, A.Ay(J).K
Next
End Function


Function KLszKLxs(A As KLxs) As KLs
Dim J%
For J = 0 To A.N - 1
    PushKL KLszKLxs, A.Ay(J).KL
Next
End Function

Function KLxs_WhPfx_RmvPfx(A As KLxs, Pfx$, Optional C As VbCompareMethod = vbTextCompare) As KLxs
Dim B As KLxs: B = KLxs_WhPfx(A, Pfx, C)
KLxs_WhPfx_RmvPfx = RmvKeyPfxzKLxs(B, Pfx, C)
End Function

Function KLxszLy(Ly$()) As KLxs
Dim A As Lnxs: A = Lnxs(Ly)
Dim B As Lnxs: B = RmvEle_ForLTrimPfx_Is2Hypron_FmLnxs(A)
KLxszLy = KLxszLnxs(B)
End Function

Function KLxszLnxs(A As Lnxs) As KLxs
If A.N = 0 Then Exit Function
Dim L As Lnx: L = A.Ay(0)
If FstChr(L.Lin) = " " Then Thw CSub, "FstLin fstChr must be non blank", "L", FmtLnx(L)
Dim M As KLx
Dim J%
For J = 0 To A.N - 1
    L = A.Ay(J)
    If FstChr(L.Lin) = " " Then
        PushLnx M.Lnxs, Lnx(Trim(L.Lin), L.Ix)
    Else
        If Not IsEmpKLx(M) Then PushKLx KLxszLnxs, M: M = EmpKLx
        M = KLx(KLzLnx(L), EmpLnxs)
    End If
Next
If Not IsEmpKLx(M) Then PushKLx KLxszLnxs, M
End Function

Function T1zKLx$(A As KLx)
T1zKLx = T1(A.KL.K)
End Function
Function KLxs_WhT1nn(A As KLxs, T1nn$) As KLxs
Dim J&, T1Ay$(), M As KLx
T1Ay = SyzSS(T1nn)
For J = 0 To A.N - 1
    M = A.Ay(J)
    If HasEle(T1Ay, T1zKLx(M)) Then PushKLx KLxs_WhT1nn, M
Next
End Function
Function KLxszLyT1nn(Ly$(), T1nn$) As KLxs
'KLxszLyT1nn = KLxszLyT1nn(KLxs(Ly), T1nn)
End Function

Private Sub SetLin(O As KLxs, K, L As Lnx)
Dim Ix&: ' Ix = IxzKKlnxes(K, O)
If Ix >= 0 Then
    PushLnx O.Ay(Ix).Lnxs, L
Else
'    PushKLxs O, KLxs(K, SngLnx(L))
End If
End Sub

Function LnxszKLxsK(A As KLxs, K) As Lnxs
Dim Ix&: Ix = IxzKLxsK(A, K)
'If Ix >= 0 Then LnxszKLxsK = A.Ay(Ix)
End Function
Private Function IxzKLxsK&(A As KLxs, K)
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
        If .KL.K = K Then
            IxzKLxsK = J
            Exit Function
        End If
    End With
Next
IxzKLxsK = -1
End Function
Function FmtKLx(A As KLx, Optional Ix& = -1) As String()
Dim P$: If Ix >= 0 Then P = "Itm-Ix:" & Ix & " "
PushI FmtKLx, P & "KLno:" & A.KL.K & " LinCnt:" & A.Lnxs.N
'PushIAy FmtKLx, TabAy(FmtLnxs(A.L))
End Function
Function FmtKLxs(A As KLxs) As String()
PushI FmtKLxs, "KLxsItmCnt=" & A.N
Dim J&
For J = 0 To A.N - 1
'    PushIAy FmtKLxs, TabAy(FmtKLx(A.Ay(J), J))
Next
End Function

Sub BrwKLxs(A As KLxs)
B FmtKLxs(A)
End Sub
Private Sub ZZ_FmtKLxs()
BrwKLxs Y_KLxs
End Sub
Private Sub Z_LnxszT1()
Dim T1$, KLxs As KLxs, Act As Lnxs, Ept As Lnxs
GoSub ZZ
Exit Sub
ZZ:
'    BrwLnxs LnxszT1(Y_KLxs, "Wdt")
    Return
Tst:
'    Act = LnxszT1("Wdt", KLxs)
    C
    Return
End Sub

Private Sub Z_KLxs()
Dim Ly$(), T1nn$
GoSub ZZ
Exit Sub
ZZ:
    BrwKLxs Y_KLxs
    Return
End Sub

Private Function Y_KLxs() As KLxs
'Y_KLxs = KLxs(Y_Lof, Y_LofT1nn)
End Function
Function FstLyOrDie_FmKLys_ByK(A As KLys, K) As String()
Dim O$(): O = FstLy_FmKLys_ByK(A, K)
If Si(O) = 0 Then Thw CSub, "Given K is not fnd in given KLys or fnd but zero-ly is return", "Given-K Given-KLys", K, FmtKLys(A)
End Function

Function FstLy_FmKLys_ByK(A As KLys, K) As String()
Dim J%
For J = 0 To A.N - 1
    If A.Ay(J).K = K Then FstLy_FmKLys_ByK = A.Ay(J).Ly: Exit Function
Next
End Function
Private Function IndentedLy_VerUseKLys(IndentedSrc$(), K$) As String()
Dim KLys As KLys: KLys = KLys_ForTakT1_AsK(IndentedSrc)
IndentedLy_VerUseKLys = FstLy_FmKLys_ByK(KLys, K)
End Function

Private Property Get Y_LofT1nn$()
Y_LofT1nn = SampLofT1nn
End Property
Private Property Get Y_Lof() As String()
Y_Lof = SampLof
End Property

Private Sub ZZZ()
QVb_Tp_KLxs:
End Sub
Private Sub ZZ_IndentedLy()
Dim IndentedSrc$(), K$
'GoSub ZZ
GoSub T0
Exit Sub
T0:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A 2"
    IndentedSrc = XX
    Erase XX
    Ept = Sy("1", "2")
    GoTo Tst
Tst:
    Act = IndentedLy(IndentedSrc, K)
    C
    Return
ZZ:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    IndentedSrc = XX
    Erase XX
    D IndentedLy(IndentedSrc, K)
    Return
End Sub

Function IndentedLy(IndentedSrc$(), Key$) As String()
'IndentedLy = IndentedLy_VerUseKLys(IndentedSrc, Key)
'IndentedLy = IndentedLy_VerNorm(IndentedSrc, Key)

Dim A$(), B$()
A = IndentedLy_VerUseKLys(IndentedSrc, Key)
B = IndentedLy_VerNorm(IndentedSrc, Key)
If Not IsEqAy(A, B) Then Stop
IndentedLy = A
End Function

Function IndentedLyOrDie(IndentedSrc$(), Key$) As String()
Dim Ly$(): Ly = IndentedLy(IndentedSrc, Key)
If Si(Ly) = 0 Then
    Thw CSub, "No given-K in given IndentedSrc", "K IndentedSrc", Key, IndentedSrc
Else
    IndentedLyOrDie = Ly
End If
End Function
Private Function IndentedLy_VerNorm(IndentedSrc$(), Key$) As String()
Dim O$()
Dim J%, I, L$, Fnd As Boolean, IsNewSection As Boolean, IsFstChrSpc As Boolean, FstA%, Hit As Boolean
Const SpcAsc% = 32
For Each I In Itr(IndentedSrc)
    L = I
    FstA = FstAsc(L)
    IsNewSection = IsAscUCas(FstA)
    If IsNewSection Then
        Hit = T1(L) = Key
    End If
    IsFstChrSpc = FstA = SpcAsc
    Select Case True
    Case IsNewSection And Not Fnd And Hit: Fnd = True
    Case IsNewSection And Fnd:             IndentedLy_VerNorm = O: Exit Function
    Case Fnd And IsFstChrSpc:              PushI O, Trim(L)
    End Select
Next
If Fnd Then IndentedLy_VerNorm = O: Exit Function
End Function

Private Sub ZZ()
ZZ_FmtKLxs
ZZ_FmtKLys
ZZ_IndentedLy
End Sub
Function DupKey(A As KLs) As KLsses
Dim D$(): D = AywDup(KyzKLs(A))
Dim ILnoss$, IStru$, I%
For I = 0 To UB(D)
    IStru = D(I)
    ILnoss = JnSpc(LnoAyzKLsK(A, IStru))
    PushKLss DupKey, KLss(IStru, ILnoss)
Next
End Function

