Attribute VB_Name = "QIde_Vbe_Cur"
Option Explicit
Private Const CMod$ = "MIde_Vbe_Cur."
Private Const Asm$ = "QIde"

Property Get CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Property

Function HasBar(BarNm$) As Boolean
HasBar = HasBarzVbe(CurVbe, BarNm)
End Function

Function HasPjf(Pjf$) As Boolean
HasPjf = HasPjfzVbe(CurVbe, Pjf)
End Function

Function PjzPjfC(Pjf$) As VBProject
Set PjzPjfC = PjzPjf(CurVbe, Pjf)
End Function

Function MdDrszVbe(A As Vbe, Optional WhStr$) As Drs
MdDrszVbe = Drs(MdTblFny, MdDryzVbe(A, WhStr))
End Function
Function MdTblFny() As String()

End Function

Sub SavCurVbe()
SavVbe CurVbe
End Sub
