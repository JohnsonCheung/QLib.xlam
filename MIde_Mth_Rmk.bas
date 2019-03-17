Attribute VB_Name = "MIde_Mth_Rmk"
Option Explicit
Const CMod$ = "MIde_Mth_Rmk."
Sub UnRmkMth(A As CodeModule, MthNm$)
UnRmkMdzFTIxAy A, MthCxtFTIxAy(Src(A), MthNm)
End Sub

Sub RmkMth(A As CodeModule, MthNm$)
RmkMdzFTIxAy A, MthCxtFTIxAy(Src(A), MthNm)
End Sub

Private Sub ZZ_RmkMth()
Dim Md As CodeModule, MthNm$
'            Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
'RmkMth M:   Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|Stop '|End Property||Property Let YYA(V)|Stop '|'|End Property"
'UnRmkMth M: Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
End Sub
Function NxtSrcIx&(Src$(), Ix&)
Const CSub$ = CMod & "NxtSrcIx"
Dim J&
For J = Ix To UB(Src)
    If LasChr(Src(J)) <> "_" Then
        NxtSrcIx = J
        Exit Function
    End If
Next
Thw CSub, "All line From Ix is Src has _ as LasChr", "Ix Src", Ix, AyAddIxPfx(Src, 1)
End Function
Function NxtMdLno&(A As CodeModule, Lno&)
Const CSub$ = CMod & "NxtMdLno"
Dim J&
For J = Lno To A.CountOfLines
    If LasChr(A.Lines(Lno, 1)) <> "_" Then
        NxtMdLno = J
        Exit Function
    End If
Next
Thw CSub, "All line From Lno has _ as LasChr", "Lno Md Src", Lno, MdNm(A), AyAddIxPfx(Src(A), 1)
End Function

Sub UnRmkMdzFTIxAy(A As CodeModule, B() As FTIx)
Dim I
For Each I In Itr(B)
    UnRmkMdzFTIx A, CvFTIx(I)
Next
End Sub

Sub UnRmkMdzFTIx(A As CodeModule, B As FTIx)
If Not IsRmkedzSrc(LyzMdFTIx(A, B)) Then Exit Sub
Dim J%, L$
For J = NxtMdLno(A, B.FmNo) To B.ToNo - 1
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
Next
End Sub

Sub RmkMdzFTIxAy(A As CodeModule, B() As FTIx)
Dim J%
For J = 0 To UB(B)
    RmkMdzFTIx A, B(J)
Next
End Sub

Sub RmkMdzFTIx(A As CodeModule, B As FTIx)
If IsRmkedzMdFTIx(A, B) Then Exit Sub
Dim J%
For J = 0 To UB(B)
    A.ReplaceLine J, "'" & A.Lines(J, 1)
Next
End Sub

Function IsRmkedzMdFTIx(A As CodeModule, B As FTIx) As Boolean
IsRmkedzMdFTIx = IsRmkedzSrc(LyzMdFTIx(A, B))
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

Function MthCxtFTIx(Src$(), MthFTIx As FTIx) As FTIx
Set MthCxtFTIx = FTIx(NxtSrcIx(Src, MthFTIx.FmIx), MthFTIx.ToIx - 1)
End Function

Function MthCxtLy(MthLy$()) As String()
MthCxtLy = CvSy(AywFTIx(MthLy, FTIx(1, Si(MthLy))))
End Function

Function MthCxtFTIxAy(Src$(), MthNm$) As FTIx()
Dim FTIx
For Each FTIx In Itr(MthFTIxAyzSrcMth(Src, MthNm))
    PushObj MthCxtFTIxAy, MthCxtFTIx(Src, CvFTIx(FTIx))
Next
End Function

Private Sub ZZ_MthCxtFTIxAy _
 _
()

Dim I
For Each I In MthCxtFTIxAy(CurSrc, CurMthNm)
    With CvFTIx(I)
        Debug.Print .FmNo, .ToNo
    End With
Next
End Sub



