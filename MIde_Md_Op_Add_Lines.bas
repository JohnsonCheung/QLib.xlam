Attribute VB_Name = "MIde_Md_Op_Add_Lines"
Option Explicit
Sub MdIns_DclLin(A As CodeModule, DclLines$)
A.InsertLines A.CountOfDeclarationLines + 1, DclLines
Debug.Print FmtQQ("MdIns_DclLin: Module(?) a DclLin is inserted", MdNm(A))
End Sub

Sub MdAppy_Ly(A As CodeModule, Ly$())
AppMdLines A, JnCrLf(Ly)
End Sub

