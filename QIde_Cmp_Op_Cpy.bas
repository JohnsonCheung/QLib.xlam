Attribute VB_Name = "QIde_Cmp_Op_Cpy"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Cmp_Op_Cpy."
Sub ThwNotCls(A As CodeModule, Fun$)
If A.Parent.Type = vbext_ct_ClassModule Then Thw Fun, "Should be a Cls", "ShtCmpTy", ShtCmpTy(A.Parent.Type)
End Sub
Private Sub CpyCls(A As CodeModule, ToPj As VBProject)
Const CSub$ = CMod & "CpyCls"
ThwNotCls A, CSub
ThwEqObj ToPj, PjzM(A), CSub, "From Md's Pj cannot eq to ToPj"
Dim T$: T = TmpFt(Fnn:=A.Name)
A.Parent.Export T
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
Sub ThwNotMod(A As CodeModule, Fun$)
If A.Parent.Type <> vbext_ct_StdModule Then Thw Fun, "Should be Mod", "Type", ShtCmpTy(A.Parent.Type)
End Sub

Sub CpyMod(A As CodeModule, ToPj As VBProject)
AddCmpzPNL ToPj, A.Name, SrcLineszM(A)
End Sub

Private Sub ZZ()
Dim A As VBComponent
Dim B As VBProject
Dim D As CodeModule
CpyCmp A, B
CpyMod D, B
End Sub

