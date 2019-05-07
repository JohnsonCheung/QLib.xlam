Attribute VB_Name = "QIde_Md_Op_Ren"
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Ren."
Private Const Asm$ = "QIde"
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
Sub RenMdzPfx(FmPfx$, ToPfx$, Optional Pj As VBProject)
Dim A As VBProject: Set A = DftPj(Pj)
Dim C As VBComponent
For Each C In A.VBComponents
    If HasPfx(C.Name, FmPfx) Then
        RenMd C.CodeModule, RplPfx(C.Name, FmPfx, ToPfx)
    End If
Next
End Sub
Sub RenMd(A As CodeModule, NewNm$)
If HasMd(PjzMd(A), NewNm) Then
    Debug.Print FmtQQ("NewMdNm[?] exists, cannot rename Md[?]", NewNm, MdNm(A))
    Exit Sub
End If
A.Parent.Name = NewNm
End Sub

Sub MthKeyDrFny()

End Sub
