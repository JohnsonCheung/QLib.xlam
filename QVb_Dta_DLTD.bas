Attribute VB_Name = "QVb_Dta_DLTD"
Option Explicit
Option Compare Text
Type DLTTT: D As Drs: End Type 'Drs::{L T1 T2 T3}
Type DLTT: D As Drs: End Type 'Drs::{L T1 T2}
Type DLTD: D As Drs: End Type 'Drs::{L T1 Dta}
Type DLTDH: D As Drs: End Type 'Drs::{L T1 Dta IsHdr}
Type DLDta: D As Drs: End Type 'Drs::{L Dta}
Private Property Get Y_LofT1nn$()
Y_LofT1nn = SampLofT1nn
End Property

Private Property Get Y_Lof() As String()
Y_Lof = SampLof
End Property

Private Sub ZZZ()
QVb_Dta_IndentSrc:
End Sub
Private Sub ZZ_IndentSrcDry()
Dim IndentSrc$()
GoSub ZZ
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
    Act = IndentSrcDry(IndentSrc)
    C
    Return
ZZ:
    Erase XX
    X "A Bc"
    X " 1"
    X " --"
    X " 2"
    X "A 2"
    IndentSrc = XX
    Erase XX
    DmpDry IndentSrcDry(IndentSrc)
    Return
End Sub
Private Sub ZZ_IndentedLy()
Dim IndentSrc$(), K$
GoSub ZZ
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
ZZ:
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
IndentSrcDrs = DrszFF("L T1 IsHdr Dta", IndentSrcDry(IndentSrc))
End Function

Function IndentSrcDry(IndentSrc$()) As Variant()
'Ret:: Dry{L T1 IsHdr Dta}
Dim Lin, L&, IsHdr As Boolean, T1$, Dta$
Const SpcAsc% = 32
For Each Lin In Itr(IndentSrc)
    L = L + 1
    If FstTwoChr(LTrim(L)) = "--" Then GoTo Nxt
    IsHdr = FstChr(Lin) <> " "
    If IsHdr Then
        T1 = T1zS(Lin)
        Dta = RmvT1(Lin)
    Else
        Dta = LTrim(Lin)
    End If
    PushI IndentSrcDry, Array(L, T1, IsHdr, Dta)
Nxt:
Next
End Function
Function IndentedLy(IndentSrc$(), Key$) As String()
Dim O$()
Dim L, Fnd As Boolean, IsNewSection As Boolean, IsFstChrSpc As Boolean, FstA%, Hit As Boolean
Const SpcAsc% = 32
For Each L In Itr(IndentSrc)
    If FstTwoChr(LTrim(L)) = "--" Then GoTo Nxt
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

Private Sub ZZ()
End Sub
Function DLTT(A As DLTDH, T1$, TT$) As DLTT
DLTT = DLTTzLDta(DLDta(A, T1), TT)
End Function

Function DLTTT(A As DLTDH, T1$, TTT$) As DLTTT
DLTTT = DLTTTzLDta(DLDta(A, T1), TTT)
End Function

Function DLTTzLDta(A As DLDta, TT$) As DLTT
Dim Dr, L&, Dta$, T1$, T2$, Dry()
For Each Dr In Itr(A.D.Dry)
    L = Dr(0)
    Dta = Dr(1)
    T1 = ShfT1(Dta)
    T2 = Dta
    PushI Dry, Array(L, T1, T2)
Next
DLTTzLDta.D = DrszFF("L " & TT, Dry)
End Function
Function DLTTTzLDta(A As DLDta, TTT$) As DLTTT
Dim Dr, L&, Dta$, T1$, T2$, T3$, Dry()
For Each Dr In Itr(A.D.Dry)
    L = Dr(0)
    Dta = Dr(1)
    T1 = ShfT1(Dta)
    T2 = ShfT1(Dta)
    T3 = Dta
    PushI Dry, Array(L, T1, T2, T3)
Next
DLTTTzLDta.D = DrszFF("L " & TTT, Dry)
End Function
Function DLDtazT1Pfx(A As DLTDH, T1Pfx$) As DLDta
Dim B As Drs, C As Drs
B = DrswColPfx(A.D, "T1", T1Pfx)
C = RmvPfxzDrs(B, "T1", T1Pfx)
DLDtazT1Pfx.D = ColEqExlEqCol(C, "IsHdr", False)
'BrwDrs2 A.D, DLDta.D, NN:="LTDH LDta": Stop

End Function

Function DLDta(A As DLTDH, T1$) As DLDta
Dim B As Drs
B = ColEqExlEqCol(A.D, "T1", T1)
DLDta.D = ColEqExlEqCol(B, "IsHdr", False)
'BrwDrs2 A.D, DLDta.D, NN:="LTDH LDta": Stop
End Function

Private Function DryOfLTD(Src$()) As Variant()
'Ret:: Dry{L T1 Dta}
Dim L&, Dta$, T1$, Lin
For Each Lin In Itr(Src)
    L = L + 1
    If FstTwoChr(LTrim(L)) = "--" Then GoTo X
    T1 = T1zS(Lin)
    Dta = RmvT1(Lin)
    PushI DryOfLTD, Array(L, T1, Dta)
X:
Next
End Function
Private Function DryOfLTDH(IndentedSrc$()) As Variant()
Dim L&, Dta$, T1$, IsHdr As Boolean, Lin
For Each Lin In Itr(IndentedSrc)
    L = L + 1
    If FstTwoChr(LTrim(Lin)) = "--" Then GoTo X
    IsHdr = FstChr(Lin) <> " "
    If IsHdr Then
        Dta = RmvT1(Lin)
        T1 = T1zS(Lin)
    Else
        Dta = LTrim(Lin)
    End If
    PushI DryOfLTDH, Array(L, T1, Dta, IsHdr)
X:
Next
End Function
Function DLTD(Src$()) As DLTD
'Ret:: *DLTD::Drs{Ix T1 Dta}
DLTD.D = DrszFF("L T1 Dta", DryOfLTD(Src))
End Function
Function DLTDH(IndentedSrc$()) As DLTDH
'Ret:: Drs{Ix T1 IsHdr Dta}
DLTDH.D = DrszFF("L T1 Dta IsHdr", DryOfLTDH(IndentedSrc))
End Function

