Attribute VB_Name = "QIde_Mth_Rmk"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Rmk."
Sub UnRmkMth(A As CodeModule, Mthn)
'UnRmkMdzFes A, MthCxtFes(Src(A), Mthn)
End Sub

Sub RmkMth(A As CodeModule, Mthn)
'RmkMdzFes A, MthCxtFes(Src(A), Mthn)
End Sub

Private Sub ZZ_RmkMth()
Dim Md As CodeModule, Mthn
'            Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
'RmkMth M:   Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|Stop '|End Property||Property Let YYA(V)|Stop '|'|End Property"
'UnRmkMth M: Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
End Sub
Function NxtMdLno(A As CodeModule, Lno)
Const CSub$ = CMod & "NxtMdLno"
Dim J&
For J = Lno To A.CountOfLines
    If LasChr(A.Lines(Lno, 1)) <> "_" Then
        NxtMdLno = J
        Exit Function
    End If
Next
Thw CSub, "All line From Lno has _ as LasChr", "Lno Md Src", Lno, Mdn(A), AddIxPfx(Src(A), 1)
End Function

Sub UnRmkMdzFes(A As CodeModule, B As Feis)
Dim J&
For J = 0 To B.N - 1
    UnRmkMdzFei A, B.Ay(J)
Next
End Sub

Sub UnRmkMdzFei(A As CodeModule, B As Fei)
'If Not IsRmkedzS(LyzMdFei(A, B)) Then Exit Sub
Stop
Dim J%, L$
'For J = NxtMdLno(A, B.FmNo) To B.ToNo - 1
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
'Next
End Sub

Sub RmkMdzFes(A As CodeModule, B As Feis)
Dim J%
For J = 0 To B.N - 1
    RmkMdzFe A, B.Ay(J)
Next
End Sub

Sub RmkMdzFe(A As CodeModule, B As Fei)
If IsRmkedzMFe(A, B) Then Exit Sub
Dim J%
'For J = 0 To UB(B)
    A.ReplaceLine J, "'" & A.Lines(J, 1)
'Next
End Sub

Function IsRmkedzMFe(A As CodeModule, B As Fei) As Boolean
'IsRmkedzMFe = IsRmkedzS(LyzMdFei(A, B))
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
MthCxtFe = Fei(NxtSrcIx(MthLy, Fe.FmIx), Fe.EIx - 1)
End Function

Function MthCxtLy(MthLy$()) As String()
MthCxtLy = CvSy(AywFei(MthLy, Fei(1, Si(MthLy))))
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



