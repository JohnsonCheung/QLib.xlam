Attribute VB_Name = "QIde_Md_Op_RenMd"
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Ren."
Private Const Asm$ = "QIde"

Sub RenTo(FmCmpn, ToNm)
If HasCmpzPN(CPj, ToNm) Then Inf CSub, "CmpToNm exist", "ToNm", ToNm: Exit Sub
Cmp(FmCmpn).Name = ToNm
End Sub
Sub Ren(NewCmpn)
CCmp.Name = NewCmpn
End Sub
Function DftPj(P As VBProject) As VBProject
If IsNothing(P) Then
    Set DftPj = CPj
Else
    Set DftPj = P
End If
End Function
Sub RenMdzPfx(FmPfx$, ToPfx$, Optional Pj As VBProject)
Dim P As VBProject: Set P = DftPj(Pj)
Dim C As VBComponent
For Each C In P.VBComponents
    If HasPfx(C.Name, FmPfx) Then
        RenMd C.CodeModule, RplPfx(C.Name, FmPfx, ToPfx)
    End If
Next
End Sub
Sub RenMd(A As CodeModule, NewNm$)
If HasMd(PjzM(A), NewNm) Then
    Debug.Print FmtQQ("NewMdn[?] exists, cannot rename Md[?]", NewNm, Mdn(A))
    Exit Sub
End If
A.Parent.Name = NewNm
End Sub

Sub MthKeyDrFny()

End Sub