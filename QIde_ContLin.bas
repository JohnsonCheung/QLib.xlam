Attribute VB_Name = "QIde_ContLin"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_ContLin."
Function ContLinzMdLno$(A As CodeModule, Lno)
Dim J&, L&
L = Lno
Dim O$: O = A.Lines(L, 1)
While LasChr(O) = "_"
    L = L + 1
    O = RmvLasChr(O) & A.Lines(L, 1)
Wend
ContLinzMdLno = O
End Function
Function NxtSrcIx&(Src$(), Optional Ix& = 0)
Const CSub$ = CMod & "NxtSrcIx"
Dim J&
For J = Ix To UB(Src)
    If LasChr(Src(J)) <> "_" Then
        NxtSrcIx = J + 1
        Exit Function
    End If
Next
Thw CSub, "All line From Ix is Src has _ as LasChr", "Ix Src", Ix, AddIxPfx(Src, 1)
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

Function ContLin$(Src$(), Ix)
ContLin = JnContLin(CvSy(AywIxCnt(Src, Ix, ContLinCnt(Src, Ix))))
Else
    ContLin = JnCrLf(AywIxCnt(Src, Ix, ContLinCnt(Src, Ix)))
End If
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

Function ContLinzML$(A As CodeModule, Lno)
Dim J&, O$()
For J = Lno To ContToLno(A, Lno)
    PushI O, RmvSfx(A.Lines(J, 1), "_")
Next
ContLinzML = JnSpc(O)
End Function


