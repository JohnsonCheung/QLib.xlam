Attribute VB_Name = "QVb_Dta_KeyLin"
Option Compare Text
Type KLx: K As String: KLno As Long: L As Lnxs: End Type
Type KLxs: N As Integer: Ay() As KLx: End Type
Type KLy: K As String: Ly() As String: End Type
Type KLys: N As Integer: Ay() As KLy: End Type
Function KLy(K$, Ly$()) As KLy
KLy.K = K
KLy.Ly = Ly
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
Dim J%
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
Function KAy_FmKLys_WhKPfx(A As KLys, KPfx$) As String()
Dim J%, K$
For J = 0 To A.N - 1
    K = A.Ay(J).K
    If HasPfx(K, KPfx) Then PushI KAy_FmKLys_WhKPfx, K
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
Function IsEmpKLy(A As KLy) As Boolean
If A.K <> "" Then Exit Function
If Si(A.Ly) <> 0 Then Exit Function
IsEmpKLy = True
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

Function ShfKLyOLy(O As KLys, K$) As String() 'Shift-KLy-Opt-Ly
If O.N = 0 Then Exit Function
If O.Ay(0).K <> K Then Exit Function
ShfKLyOLy = O.Ay(0).Ly
O = KLyseFstNEle(O)
End Function

Function ShfKLyOLin$(O As KLys, K$) ' Shift-KLy-Must-Lin
ShfKLyOLin = JnSpc(ShftKLyOLy(O, K))
End Function

Function ShfKLyMLyzKK(O As KLys, KK$) As String()

End Function
Function ShfKLyMLin$(O As KLys, K$) ' Shift-KLy-Must-Lin
ShfKLyMLin = JnSpc(ShfKLyMLy(O, K))
End Function

Function ShfKLyMLy(O As KLys, K$) As String() ' Shift-KLy-Must-Ly
Dim X As KLys: X = O
Dim Ly$(): Ly = ShftKLyOLy(O, K)
If Si(Ly) Then Thw CSub, "Fst ele of given KLys is expected to have given K and have non-0-Ly.", "Given-K Given-KLys", K, FmtKLys(O)
ShfKLyMLy = Ly
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

Function KLxs(K, KLno, L As Lnxs) As KLxs
KLxs.K = K
KLxs.KLno = KLno
KLxs.L = L
End Function
Private Sub PushKLxs(O As KLxs, M As KLxs)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function KLxszLy(Ly$()) As KLxs
Dim T1Ay$(): T1Ay = SyzSS(T1nn)
Dim L, T1$, Rst$, Ix&
For Each L In Itr(Ly)
    AsgTRst L, T1, Rst
    If HasEle(T1Ay, T1) Then
        SetLin KLxszLyT1nn, T1, Lnx(Rst, Ix)
    End If
    Ix = Ix + 1
Next
End Function
Function T1zKLx$(A As KLx)
T1zKLx = T1(A.K)
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
KLxszLyT1nn = KLxszLyT1nn(KLxs(Ly), T1nn)
End Function

Private Sub SetLin(O As KLxs, K, L As Lnx)
Dim Ix&: Ix = IxzKKlnxes(K, O)
If Ix >= 0 Then
    PushLnx O.Ay(Ix).L, L
Else
    PushKLxs O, KLxs(K, SngLnx(L))
End If
End Sub

Function LnxszKKlnxes(K, A As KLxs) As Lnxs
Dim Ix&: Ix = IxzKKlnxes(K, A)
If Ix >= 0 Then LnxszKKLxs = A.Ay(Ix)
End Function
Private Function IxzKKlnxes&(K, A As KLxs)
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
        If .K = K Then
            IxzKKlnxes = J
            Exit Function
        End If
    End With
Next
IxzKKlnxes = -1
End Function
Function FmtKLx(A As KLx, Optional Ix& = -1) As String()
If Ix >= 0 Then P = "Itm-Ix:" & Ix & " "
PushI FmtKLx, P & "KLno:" & A.KLno & " LinCnt:" & A.L.N
PushIAy FmtKLx, TabAy(FmtLnxs(A.L))
End Function
Function FmtKLxs(A As KLxs) As String()
PushI FmtKLxs, "KLxsItmCnt=" & A.N
Dim J&
For J = 0 To A.N - 1
    PushIAy FmtKLxs, TabAy(FmtKLx(A.Ay(J), J))
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
    BrwLnxs LnxszT1("Wdt", Y_KLxs)
    Return
Tst:
    Act = LnxszT1("Wdt", KLxs)
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
Y_KLxs = KLxs(Y_Lof, Y_LofT1nn)
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
