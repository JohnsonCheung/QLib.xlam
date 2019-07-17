Attribute VB_Name = "QIde_B_Op_CpyCmp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Cmp_Op_Cpy."
Sub ThwNotCls(M As CodeModule, Fun$)
If M.Parent.Type = vbext_ct_ClassModule Then Thw Fun, "Should be a Cls", "ShtCmpTy", ShtCmpTy(M.Parent.Type)
End Sub
Private Sub CpyCls(M As CodeModule, ToPj As VBProject)
Const CSub$ = CMod & "CpyCls"
ThwNotCls M, CSub
ThwEqObj ToPj, PjzM(M), CSub, "From Md's Pj cannot eq to ToPj"
Dim T$: T = TmpFt(Fnn:=M.Name)
M.Parent.Export T
ToPj.VBComponents.Import T
Kill T
End Sub
Sub CpyModAyToPj(ModAy() As CodeModule, ToPj As VBProject)
Dim I
For Each I In Itr(ModAy)
    CpyMod CvMd(I), ToPj
Next
End Sub

Sub CpyClsAyToPj(ClsAy() As CodeModule, ToPj As VBProject)
Dim I
For Each I In Itr(ClsAy)
    CpyCls CvMd(I), ToPj
Next
End Sub

Sub CpyCmp(A As VBComponent, ToPj As VBProject)
If IsCmpzCls(A) Then
    CpyCls A.CodeModule, ToPj 'If ClassModule need to export and import due to the Public/Private class property can only the set by Export/Import
Else
    CpyMod A.CodeModule, ToPj
End If
End Sub
Sub ThwNotMod(M As CodeModule, Fun$)
If M.Parent.Type <> vbext_ct_StdModule Then Thw Fun, "Should be Mod", "Type", ShtCmpTy(M.Parent.Type)
End Sub

Sub CpyMod(M As CodeModule, ToPj As VBProject)
AddCmpzSrc ToPj, M.Name, SrcLzM(M)
End Sub

Private Sub Z()
Dim A As VBComponent
Dim B As VBProject
Dim D As CodeModule
CpyCmp A, B
CpyMod D, B
End Sub

