Attribute VB_Name = "QIde_Vbe_Cur"
Option Explicit
Private Const CMod$ = "MIde_Vbe_Cur."
Private Const Asm$ = "QIde"

Function HasBar(BarNm$) As Boolean
HasBar = HasBarzV(CVbe, BarNm)
End Function

Function HasPjf(Pjf) As Boolean
HasPjf = HasPjfzV(CVbe, Pjf)
End Function

Function PjzPjfC(Pjf) As VBProject
Set PjzPjfC = PjzPjf(CVbe, Pjf)
End Function

Function MdDrszV(A As Vbe) As Drs
MdDrszV = Drs(MdTblFny, MdDryzV(A))
End Function
Function MdTblFny() As String()

End Function

Sub SavCurVbe()
SavVbe CVbe
End Sub
