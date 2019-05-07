Attribute VB_Name = "QIde_Md_Op_Add_Lines"
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Add_Lines."
Private Const Asm$ = "QIde"
Function InsDcl(A As CodeModule, Dcl$) As CodeModule
A.InsertLines FstMthLnozMd(A), Dcl
Debug.Print FmtQQ("MdInsDcl: Module(?) a DclLin is inserted", MdNm(A))
End Function

Sub ApdLy(A As CodeModule, Ly$())
ApdLines A, JnCrLf(Ly)
End Sub

