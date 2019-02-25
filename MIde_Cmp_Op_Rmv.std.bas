Attribute VB_Name = "MIde_Cmp_Op_Rmv"
Option Explicit
Sub DltCmpz(A As VBProject, MdNm$)
If Not HasCmpz(A, MdNm) Then Exit Sub
A.VBComponents.Remove A.VBComponents(MdNm)
End Sub

Sub RmvMd(A As CodeModule)
Dim M$, P$, Pj As VBProject
    M = MdNm(A)
    'Set Pj = MdzPj(A)
    P = Pj.Name
Debug.Print FmtQQ("RmvMd: Before Md(?) is deleted from Pj(?)", M, P)
A.Parent.Collection.Remove A.Parent
Debug.Print FmtQQ("RmvMd: After Md(?) is deleted from Pj(?)", M, P)
End Sub

Sub RmvCmp(A As VBComponent)
A.Collection.Remove A
End Sub

