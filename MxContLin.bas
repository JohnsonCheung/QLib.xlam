Attribute VB_Name = "MxContLin"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxContLin."

Sub RmvContLin(M As CodeModule, Lno&)
M.DeleteLines Lno, NContLin(Src(M), Lno)
End Sub

Function ContLin$(Src$(), Ix)
Dim O$, I&
O = Src(Ix)
For I = Ix + 1 To UB(Src)
    If LasChr(O) <> "_" Then ContLin = O: Exit Function
    O = RmvLasChr(O) & LTrim(Src(I))
Next
ThwImpossible CSub
End Function

Function ContLy(Src$()) As String()
Dim PushPrv As Boolean
Dim O$()
    Dim L, J&: For Each L In Itr(Src)
        If PushPrv Then
            O(J) = O(J) & LTrim(L)
        Else
            PushI O, L
        End If
        PushPrv = LasChr(L) = "_"
        If PushPrv Then
            O(J) = RmvLasChr(O(J))
        End If
        J = J + 1
    Next
ContLy = O
End Function

Function ContLyzM(M As CodeModule, Lno&) As String()
Dim O$()
Dim J&: For J = Lno To M.CountOfLines
    Dim L$: L = M.Lines(J, 1)
    PushI ContLyzM, L
    If LasChr(L) <> "_" Then Exit Function
Next
ThwImpossible CSub
End Function
Function JnContLy$(ContLy$())
Dim O$(): O = ContLy
Dim J%, U%
    U = UB(O)
    For J = 0 To U - 1
        O(J) = RmvLasChr(O(J))
    Next
    For J = 1 To U
        O(J) = LTrim(O(J))
    Next
JnContLy = Jn(O)
End Function

Function ContLinzM$(M As CodeModule, Lno&)
ContLinzM = JnContLy(ContLyzM(M, Lno))
End Function

Function SrcLinzNxt$(M As CodeModule, Lno&)
SrcLinzNxt = M.Lines(NxtLnozML(M, Lno), 1)
End Function

Function NxtLnozML&(M As CodeModule, Lno&)
Dim J&
For J = Lno + 1 To M.CountOfLines
    If LasChr(M.Lines(J - 1, 1)) <> "_" Then
        NxtLnozML = J
        Exit Function
    End If
Next
'No need to throw error, just exit it returns -1
'Thw CSub, "Cannot find Lno where to insert CSub of a given method", "Mthn MthLy", A.Mthn, AwFT(Src, A.FmIx, A.EIx)
End Function

Function NxtIxzSrc&(Src$(), Optional FmIx&)
Dim J&
For J = FmIx + 1 To UB(Src)
    If LasChr(Src(J - 1)) <> "_" Then
        NxtIxzSrc = J
        Exit Function
    End If
Next
'No need to throw error, just exit it returns -1
'Thw CSub, "Cannot find Lno where to insert CSub of a given method", "Mthn MthLy", A.Mthn, AwFT(Src, A.FmIx, A.EIx)
NxtIxzSrc = -1
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
Function ContLinCntzM%(M As CodeModule, Lno)
Dim J&, O%
For J = Lno To M.CountOfLines
    O = O + 1
    If LasChr(M.Lines(J, 1)) <> "_" Then
        ContLinCntzM = O
        Exit Function
    End If
Next
Thw CSub, "LasLin of Md cannot be end of [_]", "LasLin-Of-Md Md", M.Lines(M.CountOfLines, 1), Mdn(M)
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

Private Function ContToLno(M As CodeModule, Lno)
Dim J&
For J = Lno To M.CountOfLines
   If Not HasSfx(M.Lines(J, 1), " _") Then
        ContToLno = J
        Exit Function
   End If
Next
ThwImpossible CSub ' CSub, "each lines ends with sfx _ started from Lno, which is impossible", "Md Started-Fm-Lno", Mdn(A), Lno
End Function

Function ContEIx&(Src$(), Ix&)
Dim O&: For O = Ix To UB(Src)
    If LasChr(Src(O)) <> "_" Then ContEIx = O: Exit Function
Next
Thw CSub, "las lin of @Src has LasChr = '_'", "Las-Src-Ele", LasEle(Src)
End Function