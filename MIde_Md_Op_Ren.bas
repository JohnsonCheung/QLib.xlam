Attribute VB_Name = "MIde_Md_Op_Ren"
Option Explicit
Sub RenTo(FmCmpNm$, ToNm$)
Cmp(FmCmpNm).Name = ToNm
End Sub
Sub Ren(NewCmpNm$)
CurCmp.Name = NewCmpNm
End Sub

Sub RenMd(A As CodeModule, NewNm$)
A.Parent.Name = NewNm
End Sub

Sub MthKeyDrFny()

End Sub
