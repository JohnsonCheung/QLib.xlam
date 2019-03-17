Attribute VB_Name = "MIde_ContLin"
Option Explicit
Const CMod$ = "MIde__ContLin."
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
Function NxtSrcIx&(Src$(), Ix&)
Dim O&
For O = Ix + 1 To UB(Src)
    If LasChr(Src(Ix)) <> "_" Then
        NxtSrcIx = O
        Exit Function
    End If
Next
NxtSrcIx = -1
End Function
Private Sub Z_ContLin()
Dim Src$(), MthFmIx%
MthFmIx = 0
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
    Act = ContLin(Src, MthFmIx)
    C
    Return
End Sub
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

Function ContLin$(A$(), Ix)
ContLin = JnCrLf(AywIxCnt(A, Ix, ContLinCnt(A, Ix)))
End Function

Function ContFTIxzSrc(Src$(), Ix) As FTIx
Set ContFTIxzSrc = FTIxzIxCnt(Ix, ContLinCnt(Src, Ix))
End Function

Function ContFTIxzMd(A As CodeModule, Lno&) As FTIx
Set ContFTIxzMd = FTIx(Lno - 1, ContToLno(A, Lno) - 1)
End Function

Private Function ContToLno&(A As CodeModule, Lno&)
Dim J&
For J = Lno To A.CountOfLines
   If Not HasSfx(A.Lines(J, 1), " _") Then
        ContToLno = J
        Exit Function
   End If
Next
If Lno <> A.CountOfLines Then Thw CSub, "each lines ends with sfx _ started from Lno, which is impossible", "Md Started-Fm-Lno", MdNm(A), Lno
End Function
Function ContLinzMd$(A As CodeModule, Lno&)
Dim J&, O$()
For J = Lno To ContToLno(A, Lno)
    PushI O, RmvSfx(A.Lines(J, 1), "_")
Next
ContLinzMd = JnSpc(O)
End Function


