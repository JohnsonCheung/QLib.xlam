Attribute VB_Name = "MIde_Md_Op_Add_Lines"
Option Explicit
Function MdInsDclLines(A As CodeModule, DclLines$) As CodeModule
A.InsertLines FstMthLnoMd(A), DclLines
Debug.Print FmtQQ("MdInsDclLines: Module(?) a DclLin is inserted", MdNm(A))
End Function

Function MdAppLy(A As CodeModule, Ly$()) As CodeModule
Set MdAppLy = MdAppLines(A, JnCrLf(Ly))
End Function

