Attribute VB_Name = "QIde_ContLin"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_ContLin."
Function ContLin$(Src$(), Ix)
Dim O$, I&
O = Src(Ix)
For I = Ix + 1 To UB(Src)
    If LasChr(O) <> "_" Then ContLin = O: Exit Function
    O = RmvLasChr(O) & LTrim(Src(Ix))
Next
ThwImpossible CSub
End Function

Function ContLinzML$(M As CodeModule, Lno)
Dim L&, O$
O = M.Lines(Lno, 1)
For L = Lno + 1 To M.CountOfLines
    If LasChr(O) <> "_" Then ContLinzML = O: Exit Function
    O = RmvLasChr(O) & LTrim(M.Lines(Lno, 1))
Next
ThwImpossible CSub
End Function
Function NxtSrcIx&(Src$(), Optional FmIx&)
Dim J&
For J = FmIx + 1 To UB(Src)
    If LasChr(Src(J - 1)) <> "_" Then
        NxtSrcIx = J
        Exit Function
    End If
Next
'No need to throw error, just exit it returns -1
'Thw CSub, "Cannot find Lno where to insert CSub of a given method", "Mthn MthLy", A.Mthn, AywFT(Src, A.FmIx, A.EIx)
NxtSrcIx = -1
End Function

Private Sub Z_ContLin()
Dim Src$(), MthIx
MthIx = 0
Dim O$(3)
O(0) = "A _"
O(1) = "  B _"
O(2) = "C"
O(3) = "D"
Src = O
Ept = "ABC"
GoSub Tst
Exit Sub
Tst:
    Act = ContLin(Src, MthIx)
    C
    Return
End Sub
Function ContLinCntzM%(A As CodeModule, Lno)
Dim J&, O%
For J = Lno To A.CountOfLines
    O = O + 1
    If LasChr(A.Lines(J, 1)) <> "_" Then
        ContLinCntzM = O
        Exit Function
    End If
Next
Thw CSub, "LasLin of Md cannot be end of [_]", "LasLin-Of-Md Md", A.Lines(A.CountOfLines, 1), Mdn(A)
End Function

Function ContLinCnt%(Src$(), Ix)
If Si(Src) = 0 Then Exit Function
Dim J&, O%
For J = Ix To UB(Src)
    O = O + 1
    If LasChr(Src(J)) <> "_" Then
        ContLinCnt = O
        Exit Function
    End If
Next
Thw CSub, "LasLin of Src cannot be end of [_]", "LasLin-Of-Src Src", LasEle(Src), Src
End Function
Function JnContLin$(ContLy)
Dim J%, L$, O$()
PushI O, ContLy(0)
For J = 1 To UB(ContLy) - 1

    PushI O, ContLy(J)
Next
End Function

Private Function ContToLno(A As CodeModule, Lno)
Dim J&
For J = Lno To A.CountOfLines
   If Not HasSfx(A.Lines(J, 1), " _") Then
        ContToLno = J
        Exit Function
   End If
Next
ThwImpossible CSub ' CSub, "each lines ends with sfx _ started from Lno, which is impossible", "Md Started-Fm-Lno", Mdn(A), Lno
End Function

