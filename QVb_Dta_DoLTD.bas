Attribute VB_Name = "QVb_Dta_DoLTD"
Option Explicit
Option Compare Text
Public Const FFoLTDH$ = "L T1 Dta IsHdr"
Type DoLTTT: D As Drs: End Type 'Drs-L-T1-T2-T3
Type DoLTT:  D As Drs: End Type 'Drs-L-T1-T2
Type DoLDta: D As Drs: End Type 'Drs-L-Dta
Type DoLTD:  D As Drs: End Type 'Drs-L-T1-Dta
Type DoLTDH: D As Drs: End Type 'Drs-L-T1-Dta-IsHdr

Private Property Get Y_LofT1nn$()
Y_LofT1nn = LofT1nn
End Property

Private Property Get Y_Lof() As String()
Y_Lof = SampLof
End Property

Private Sub Z()
QVb_Dta_IndentSrc:
End Sub

Private Sub Z_IndentSrcDy()
Dim IndentSrc$()
GoSub Z
GoSub T0
Exit Sub
T0:
    Erase XX
    X "A Bc"
    X " 1"
    X " --"
    X " 2"
    X "A 2"
    IndentSrc = XX
    Erase XX
    Ept = Array( _
        Array(0&, "A", True, "Bc"), _
        Array(1&, "A", False, "1"), _
        Array(3&, "A", False, "2"), _
        Array(4&, "A", True, "2"))
    GoTo Tst
Tst:
    Act = IndentSrcDy(IndentSrc)
    C
    Return
Z:
    Erase XX
    X "A Bc"
    X " 1"
    X " --"
    X " 2"
    X "A 2"
    IndentSrc = XX
    Erase XX
    DmpDy IndentSrcDy(IndentSrc)
    Return
End Sub
Private Sub Z_IndentedLy()
Dim IndentSrc$(), K$
GoSub Z
GoSub T0
Exit Sub
T0:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A 2"
    IndentSrc = XX
    Erase XX
    Ept = Sy("1", "2")
    GoTo Tst
Tst:
    Act = IndentedLy(IndentSrc, K)
    C
    Return
Z:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A Bc"
    X " 1 2"
    X " 2 3"
    IndentSrc = XX
    Erase XX
    D IndentedLy(IndentSrc, K)
    Return
End Sub

Function IndentedLyOrDie(IndentSrc$(), Key$) As String()
Dim Ly$(): Ly = IndentedLy(IndentSrc, Key)
If Si(Ly) = 0 Then
    Thw CSub, "No given-K in given IndentSrc", "K IndentSrc", Key, IndentSrc
Else
    IndentedLyOrDie = Ly
End If
End Function
Function IndentSrcDrs(IndentSrc$()) As Drs
IndentSrcDrs = DrszFF("L T1 IsHdr Dta", IndentSrcDy(IndentSrc))
End Function

Function IndentSrcDy(IndentSrc$()) As Variant()
'Ret:: Dy{L T1 IsHdr Dta}
Dim Lin, L&, IsHdr As Boolean, T1$, Dta$
Const SpcAsc% = 32
For Each Lin In Itr(IndentSrc)
    L = L + 1
    If Fst2Chr(LTrim(L)) = "--" Then GoTo Nxt
    IsHdr = FstChr(Lin) <> " "
    If IsHdr Then
        T1 = T1zS(Lin)
        Dta = RmvT1(Lin)
    Else
        Dta = LTrim(Lin)
    End If
    PushI IndentSrcDy, Array(L, T1, IsHdr, Dta)
Nxt:
Next
End Function
Function IndentedLy(IndentSrc$(), Key$) As String()
Dim O$()
Dim L, Fnd As Boolean, IsNewSection As Boolean, IsFstChrSpc As Boolean, FstA%, Hit As Boolean
Const SpcAsc% = 32
For Each L In Itr(IndentSrc)
    If Fst2Chr(LTrim(L)) = "--" Then GoTo Nxt
    FstA = FstAsc(L)
    IsNewSection = IsAscUCas(FstA)
    If IsNewSection Then
        Hit = T1(L) = Key
    End If
    
    IsFstChrSpc = FstA = SpcAsc
    Select Case True
    Case IsNewSection And Not Fnd And Hit: Fnd = True
    Case IsNewSection And Fnd:             IndentedLy = O: Exit Function
    Case Fnd And IsFstChrSpc:              PushI O, Trim(L)
    End Select
Nxt:
Next
If Fnd Then IndentedLy = O: Exit Function
End Function

Function DoLTT(A As DoLTDH, T1$, TT$) As DoLTT
DoLTT = DoLTTzLDta(DoLDta(A, T1), TT)
End Function

Function DoLTTT(A As DoLTDH, T1$, TTT$) As DoLTTT
DoLTTT = DoLTTTzLDta(DoLDta(A, T1), TTT)
End Function

Function DoLTTzLDta(A As DoLDta, TT$) As DoLTT
Dim Dr, L&, Dta$, T1$, T2$, Dy()
For Each Dr In Itr(A.D.Dy)
    L = Dr(0)
    Dta = Dr(1)
    T1 = ShfT1(Dta)
    T2 = Dta
    PushI Dy, Array(L, T1, T2)
Next
DoLTTzLDta.D = DrszFF("L " & TT, Dy)
End Function
Function DoLTTTzLDta(A As DoLDta, TTT$) As DoLTTT
Dim Dr, L&, Dta$, T1$, T2$, T3$, Dy()
For Each Dr In Itr(A.D.Dy)
    L = Dr(0)
    Dta = Dr(1)
    T1 = ShfT1(Dta)
    T2 = ShfT1(Dta)
    T3 = Dta
    PushI Dy, Array(L, T1, T2, T3)
Next
DoLTTTzLDta.D = DrszFF("L " & TTT, Dy)
End Function
Function DoLDtazT1Pfx(A As DoLTDH, T1Pfx$) As DoLDta
Dim B As Drs, C As Drs
B = ColPfx(A.D, "T1", T1Pfx)
C = RmvPfxzDrs(B, "T1", T1Pfx)
DoLDtazT1Pfx.D = DwEqE(C, "IsHdr", False)
'BrwDrs2 A.D, DoLDta.D, NN:="LTDH LDta": Stop

End Function

Function DoLDta(A As DoLTDH, T1$) As DoLDta
Dim B As Drs
B = DwEqE(A.D, "T1", T1)
DoLDta.D = DwEqE(B, "IsHdr", False)
'BrwDrs2 A.D, DoLDta.D, NN:="LTDH LDta": Stop
End Function

Private Function DyoLTD(Src$()) As Variant()
'Ret:: Dy{L T1 Dta}
Dim L&, Dta$, T1$, Lin
For Each Lin In Itr(Src)
    L = L + 1
    If Fst2Chr(LTrim(L)) = "--" Then GoTo X
    T1 = T1zS(Lin)
    Dta = RmvT1(Lin)
    PushI DyoLTD, Array(L, T1, Dta)
X:
Next
End Function
Private Function DyoLTDH(IndentedSrc$()) As Variant()
Dim L&, Dta$, T1$, IsHdr As Boolean, Lin
For Each Lin In Itr(IndentedSrc)
    L = L + 1
    If Fst2Chr(LTrim(Lin)) = "--" Then GoTo X
    IsHdr = FstChr(Lin) <> " "
    If IsHdr Then
        Dta = RmvT1(Lin)
        T1 = T1zS(Lin)
    Else
        Dta = LTrim(Lin)
    End If
    PushI DyoLTDH, Array(L, T1, Dta, IsHdr)
X:
Next
End Function

Function DoLTD(Src$()) As Drs
'Ret : :DoLTD @@
':DoLTD: :Drs-L-T1-Dta
DoLTD = DrszFF("L T1 Dta", DyoLTD(Src))
End Function

Function FmtDoLTDH(A As DoLTDH, T1$) As String()
Dim B As Drs: B = Dw2Eq(A.D, "IsHdr T1", False, T1)
FmtDoLTDH = StrCol(B, "Lin")
End Function

Function DoLTDH(IndentedSrc$()) As DoLTDH
'Ret: L T1 Dta IsHdr: @@
DoLTDH.D = DrszFF(FFoLTDH, DyoLTDH(IndentedSrc))
End Function


'

