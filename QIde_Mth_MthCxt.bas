Attribute VB_Name = "QIde_Mth_MthCxt"
Option Compare Text
Option Explicit
Type Fc
    FmLno As Long
    Cnt As Long
End Type
Type Fcs: N As Long: Ay() As Fc: End Type
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Rmk."

Sub UnRmkMth(M As CodeModule, Mthn)
'UnRmkMdzFes A, MthCxtFes(Src(A), Mthn)
End Sub

Function Fc(FmLno, Cnt) As Fc
If Cnt <= 0 Then Exit Function
If FmLno <= 0 Then Exit Function
Fc.FmLno = FmLno
Fc.Cnt = Cnt
End Function

Sub RmkLineszFc(M As CodeModule, Fc As Fc)
RmkLines M, Fc.FmLno, Fc.Cnt
End Sub

Function MthCxtFcs(Src$(), Mthn) As Fcs
MthCxtFcs = MthCxtFcszzSM(Src, MthCxtFcs(Src, Mthn))
End Function

Function MthCxtFcszzSM(Src$(), Mth As Fcs) As Fcs
Dim J&
For J = 0 To Mth.N - 1
    PushFc MthCxtFcszzSM, MthCxtFczzSM(Src, Mth.Ay(J))
Next
End Function

Function NContLin(Src$(), MthIx) As Byte
Dim J&, O&
For J = MthIx To UB(Src)
    O = O + 1
    If LasChr(Src(J)) <> "_" Then NContLin = O: Exit Function
Next
Thw CSub, "LasEle of Src has LasChr = _", "Src", Src
End Function

Function AddFc(A As Fc, B As Fc) As Fcs
PushFc AddFc, A
PushFc AddFc, B
End Function

Function FmtFcs$(A As Fcs)
Dim O$(), J&
For J = 0 To A.N - 1
    PushI O, FmtFc(A.Ay(J))
Next
FmtFcs = JnCrLf(O)
End Function

Function FmtFc$(Fc As Fc)
With Fc
FmtFc = "Fc " & .FmLno & " " & .Cnt
End With
End Function

Function MthCxtFczzSM(Src$(), Mth As Fc) As Fc
With Mth
Dim N%: N = NContLin(Src, .FmLno)
MthCxtFczzSM = Fc(.FmLno - N, .Cnt - N - 1)
End With
End Function
Function MthCxtFc(M As CodeModule, Mthn) As Fc

End Function
Sub RmkMthzN(M As CodeModule, Mthn)
RmkLineszFc M, MthCxtFc(M, Mthn)
End Sub

Function MthFcs(M As CodeModule, Mthn) As Fcs
Dim Ix, S$()
S = Src(M)
For Each Ix In Itr(MthIxyzSN(S, Mthn))
    PushFc MthFcs, Fc(Ix + 1, ContLinCnt(S, Ix))
Next
End Function

Sub PushFc(O As Fcs, M As Fc)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Sub RmkLines(M As CodeModule, Lno&, N&)
Dim J&
For J = Lno To Lno + N - 1
    RmkLin M, J
Next
End Sub

Sub RmkLin(M As CodeModule, Lno&)
M.ReplaceLine M, "'" & M.Lines(Lno, 1)
End Sub

Sub RmkMth()
RmkMthzN CMd, CMthn
End Sub

Private Sub ZZ_RmkMth()
Dim Md As CodeModule, Mthn
'            Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
'RmkMth M:   Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|Stop '|End Property||Property Let YYA(V)|Stop '|'|End Property"
'UnRmkMth M: Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
End Sub
Function NxtMdLno(M As CodeModule, Lno)
Const CSub$ = CMod & "NxtMdLno"
Dim J&
For J = Lno To M.CountOfLines
    If LasChr(M.Lines(Lno, 1)) <> "_" Then
        NxtMdLno = J
        Exit Function
    End If
Next
Thw CSub, "All line From Lno has _ as LasChr", "Lno Md Src", Lno, Mdn(M), AddIxPfx(Src(M), 1)
End Function

Sub UnRmkMdzFes(M As CodeModule, B As Feis)
Dim J&
For J = 0 To B.N - 1
    UnRmkMdzFei M, B.Ay(J)
Next
End Sub

Sub UnRmkMdzFei(M As CodeModule, B As Fei)
'If Not IsRmkedzS(LyzMdFei(A, B)) Then Exit Sub
Stop
Dim J%, L$
'For J = NxtMdLno(A, B.FmNo) To B.ToNo - 1
    L = M.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    M.ReplaceLine J, Mid(L, 2)
'Next
End Sub

Sub RmkMdzFes(M As CodeModule, B As Feis)
Dim J%
For J = 0 To B.N - 1
    RmkMdzFe M, B.Ay(J)
Next
End Sub

Sub RmkMdzFe(M As CodeModule, B As Fei)
If IsRmkedzMFe(M, B) Then Exit Sub
Dim J%
'For J = 0 To UB(B)
    M.ReplaceLine J, "'" & M.Lines(J, 1)
'Next
End Sub

Function IsRmkedzMFe(M As CodeModule, B As Fei) As Boolean
'IsRmkedzMFe = IsRmkedzS(LyzMdFei(A, B))
End Function
Function CDMthCxt() As Drs
CDMthCxt = DMthCxt(CMd, CMthLno)
End Function
Function DMthCxt(M As CodeModule, MthLno&) As Drs
'*DMthCxt::Drs{L Lin}
Dim Dry(), L&, ELin$, MthLin$, Lin$
MthLin = M.Lines(MthLno, 1)
ELin = MthELin(MthLin)
For L = SrcLnozNxt(M, MthLno) To M.CountOfLines
    Lin = M.Lines(L, 1)
    If Lin = ELin Then
        GoTo X
    End If
    Lin = M.Lines(L, 1)
    PushI Dry, Array(L, Lin)
Next
ThwImpossible CSub
X:
DMthCxt = DrszFF("L MthLin", Dry)
End Function
Function IsRmkedzMthLy(MthLy$()) As Boolean
If Si(MthLy) = 0 Then Exit Function
If Not HasPfx(MthLy(0), "Stop '") Then Exit Function
Dim L
For Each L In MthLy
    If Left(L, 1) <> "'" Then Exit Function
Next
IsRmkedzMthLy = True
End Function
Function MthCxtFe(MthLy$(), Fe As Fei) As Fei
MthCxtFe = Fei(SrcIxzNxt(MthLy, Fe.FmIx), Fe.EIx - 1)
End Function

Function MthCxt$(MthLy$())
MthCxt = JnCrLf(MthCxtLy(MthLy))
End Function

Function MthCxtLy(MthLy$()) As String()
If Si(MthLy) = 0 Then Exit Function
Dim L&: L = FstMthIx(MthLy): If L = -1 Then Thw CSub, "Given MthLy is not MthLy", "MthLy", MthLy
Dim J%
For J = SrcIxzNxt(MthLy, L) To UB(MthLy) - 1
    PushI MthCxtLy, MthLy(J)
Next
End Function

Function MthCxtRgs(Src$(), Mthn) As MthRgs
Dim A As Feis, J&
'A = MthFeszSN(Src, Mthn)
For J = 0 To A.N - 1
    'PushFei MthCxtFeis, MthCxtFei(Src, A.Ay(J))
Next
End Function

Private Sub ZZ_MthCxtFeis _
 _
()
Stop
Dim I
'For Each I In MthCxtFeis(CSrc, CurMthn)
    'With CvFei(I)
'        Debug.Print .FmNo, .ToNo
    'End With
'Next
End Sub



