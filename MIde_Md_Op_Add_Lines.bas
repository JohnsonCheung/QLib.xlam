Attribute VB_Name = "MIde_Md_Op_Add_Lines"
Option Explicit
Function MdInsDcl(A As CodeModule, Dcl$) As CodeModule
A.InsertLines FstMthLnoMd(A), Dcl
Debug.Print FmtQQ("MdInsDcl: Module(?) a DclLin is inserted", MdNm(A))
End Function

Function MdApdLy(A As CodeModule, Ly$()) As CodeModule
Set MdApdLy = MdApdLines(A, JnCrLf(Ly))
End Function

