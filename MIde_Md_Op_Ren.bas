Attribute VB_Name = "MIde_Md_Op_Ren"
Option Explicit
Sub RenTo(FmCmpNm$, ToNm$)
Cmp(FmCmpNm).Name = ToNm
End Sub
Sub Ren(NewCmpNm$)
CurCmp.Name = NewCmpNm
End Sub
Function DftPj(A As VBProject) As VBProject
If IsNothing(A) Then
    Set DftPj = CurPj
Else
    Set DftPj = A
End If
End Function
Sub RenMdPfx(FmPfx$, ToPfx$, Optional Pj As VBProject)
Dim A As VBProject: Set A = DftPj(Pj)

End Sub
Sub RenMd(A As CodeModule, NewNm$)

A.Parent.Name = NewNm
End Sub

Sub MthKeyDrFny()

End Sub
