Attribute VB_Name = "MIde_Vbe_Cur"
Option Explicit

Property Get CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Property

Function HasBar(BarNm$) As Boolean
HasBar = HasBarzVbe(CurVbe, BarNm)
End Function

Function HasPjf(Pjf$) As Boolean
HasPjf = HasPjfzVbe(CurVbe, Pjf)
End Function

Function PjzPjf(A) As VBProject
Set PjzPjf = PjzPjfVbe(CurVbe, A)
End Function

Function MdDrszVbe(A As Vbe, Optional WhStr$) As Drs
MdDrszVbe = Drs(MdTblFny, MdDryzVbe(A, WhStr))
End Function
Function MdTblFny() As String()

End Function

Sub SavCurVbe()
SavVbe CurVbe
End Sub
