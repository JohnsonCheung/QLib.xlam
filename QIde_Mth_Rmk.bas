Attribute VB_Name = "QIde_Mth_Rmk"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Rmk."
Sub UnRmkMth(A As CodeModule, Mthn)
UnRmkMdzFEIxs A, MthCxtFEIxs(Src(A), Mthn)
End Sub

Sub RmkMth(A As CodeModule, Mthn)
RmkMdzFEIxs A, MthCxtFEIxs(Src(A), Mthn)
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

Sub UnRmkMdzFEIxs(A As CodeModule, B As FEIxs)
Dim J&
For J = 0 To B.N - 1
    UnRmkMdzFEIx A, B.Ay(J)
Next
End Sub

Sub UnRmkMdzFEIx(A As CodeModule, B As FEIx)
'If Not IsRmkedzSrc(LyzMdFEIx(A, B)) Then Exit Sub
Stop
Dim J%, L$
'For J = NxtMdLno(A, B.FmNo) To B.ToNo - 1
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
'Next
End Sub

Sub RmkMdzFEIxs(A As CodeModule, B As FEIxs)
Dim J%
For J = 0 To B.N - 1
    RmkMdzFEIx A, B.Ay(J)
Next
End Sub

Sub RmkMdzFEIx(A As CodeModule, B As FEIx)
If IsRmkedzMdFEIx(A, B) Then Exit Sub
Dim J%
'For J = 0 To UB(B)
    A.ReplaceLine J, "'" & A.Lines(J, 1)
'Next
End Sub

Function IsRmkedzMdFEIx(A As CodeModule, B As FEIx) As Boolean
'IsRmkedzMdFEIx = IsRmkedzSrc(LyzMdFEIx(A, B))
End Function

Function IsRmkedzSrc(A$()) As Boolean
If Si(A) = 0 Then Exit Function
If Not HasPfx(A(0), "Stop '") Then Exit Function
Dim L
For Each L In Itr(A)
    If Left(L, 1) <> "'" Then Exit Function
Next
IsRmkedzSrc = True
End Function

Function MthCxtFEIx(Src$(), MthFEIx As FEIx) As FEIx
MthCxtFEIx = FEIx(NxtSrcIx(Src, MthFEIx.FmIx), MthFEIx.EIx - 1)
End Function

Function MthCxtLy(MthLy$()) As String()
MthCxtLy = CvSy(AywFEIx(MthLy, FEIx(1, Si(MthLy))))
End Function

Function MthCxtFEIxs(Src$(), Mthn) As FEIxs
Dim A As FEIxs, J&
A = MthFEIxszSN(Src, Mthn)
For J = 0 To A.N - 1
    PushFEIx MthCxtFEIxs, MthCxtFEIx(Src, A.Ay(J))
Next
End Function

Private Sub ZZ_MthCxtFEIxs _
 _
()
Stop
Dim I
'For Each I In MthCxtFEIxs(CurSrc, CurMthn)
    'With CvFEIx(I)
'        Debug.Print .FmNo, .ToNo
    'End With
'Next
End Sub



