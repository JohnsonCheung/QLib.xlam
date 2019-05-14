Attribute VB_Name = "QIde_Md_Op_MdOp"
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Add_Lines."
Private Const Asm$ = "QIde"
Function InsDcl(A As CodeModule, Dcl$) As CodeModule
A.InsertLines FstMthLnozM(A), Dcl
Debug.Print FmtQQ("MdInsDcl: Module(?) a DclLin is inserted", Mdn(A))
End Function

Sub ApdLy(A As CodeModule, Ly$())
ApdLines A, JnCrLf(Ly)
End Sub

Function TmpMod() As CodeModule
Dim T$: T = TmpNm("TmpMod")
AddModzPN CPj, T
Set TmpMod = Md(T)
End Function
Function TmpModNyzP(P As VBProject) As String()
TmpModNyzP = AywPfx(ModNyzP(P), "TmpMod")
End Function

Sub ClrTmpMod()
Dim N
For Each N In TmpModNyzP(CPj)
    If HasPfx(Md(N), "TmpMod") Then RmvCmpzN N
Next
End Sub

