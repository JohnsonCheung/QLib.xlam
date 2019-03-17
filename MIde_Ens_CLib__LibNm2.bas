Attribute VB_Name = "MIde_Ens_CLib__LibNm2"
Option Explicit
Sub EnsClsLibNmPj(Pj As VBProject)
Dim C, L$
For Each C In ClsAyPj(Pj)
    EnsClsLibNmCls CvMd(C)
Next
End Sub
Sub EnsClsLibNm()
EnsClsLibNmPj CurPj
End Sub
Function HasClsLibLin(A As CodeModule) As Boolean

End Function
Function ClsAyPjLibNm(Pj As VBProject, LibNm) As CodeModule()

End Function
Sub EnsClsLibNmCls(A As CodeModule)
If HasClsLibLin(A) Then Exit Sub
Dim L$
L = ClsLibNmLin(PjNmzMd(A))
A.InsertLines FstLnozAftOptMd(A), L
End Sub
Function FstIxzAftOpt&(Src$())
Dim J&, L$
For J = 0 To UBound(Src)
    If Not HasPfx(Src(J), "Option") Then
        If IsCdLin(L) Then
            FstIxzAftOpt = J
            Exit Function
        End If
    End If
    If IsMthLin(L) Then
        FstIxzAftOpt = J - 1
        Exit Function
    End If
Next
FstIxzAftOpt = 0
End Function

Function FstLnozAftOptMd%(A As CodeModule)
Dim J%, L$
For J = 1 To A.CountOfDeclarationLines
    L = A.Lines(J, 1)
    If Not HasPfx(L, "Option") Then
        If IsCdLin(L) Then
            FstLnozAftOptMd = J
            Exit Function
        End If
    End If
Next
FstLnozAftOptMd = J
End Function
Private Function ClsLibNmLin$(LibNm$)
ClsLibNmLin = "Private Const ClsLibNm$=" & QuoteDbl(LibNm)
End Function
Private Function HasClsLibNmLinMd(A As CodeModule) As Boolean

End Function
