Attribute VB_Name = "MIde_ContLin"
Option Explicit
Const CMod$ = "MIde__ContLin."
Function ContLinMdLno$(A As CodeModule, Lno)
Dim J&, L&
L = Lno
Dim O$: O = A.Lines(L, 1)
While LasChr(O) = "_"
    L = L + 1
    O = RmvLasChr(O) & A.Lines(L, 1)
Wend
ContLinMdLno = O
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
Function ContLin$(A$(), Ix)
Const CSub$ = CMod & "ContLin"
If Ix <= -1 Then Exit Function
Dim J&, I$
Dim O$, IsCont As Boolean
For J = Ix To UB(A)
   I = A(J)
   O = O & LTrim(I)
   IsCont = HasSfx(I, " _")
   If IsCont Then O = RmvSfx(RmvSfx(O, "_"), " ")
   If Not IsCont Then Exit For
Next
If IsCont Then Thw CSub, "each lines {Src} ends with sfx _, which is impossible"
ContLin = O
End Function
Function ContLinMd$(A As CodeModule, Lno&)
Const CSub$ = CMod & "ContLinMd"
Dim Ix&
If Ix <= -1 Then Exit Function
Dim J&, I$
Dim O$, IsCont As Boolean
For J = Lno To A.CountOfLines
   I = A.Lines(J, 1)
   O = O & LTrim(I)
   If Not HasSfx(I, " _") Then
        ContLinMd = O & LTrim(I)
   End If
   If IsCont Then O = RmvSfx(RmvSfx(O, "_"), " ")
   If Not IsCont Then Exit For
Next
If IsCont Then Thw CSub, "each lines {Src} ends with sfx _, which is impossible"
ContLinMd = O
End Function


