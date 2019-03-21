Attribute VB_Name = "MIde_Vbe_Cur"
Option Explicit

Property Get CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Property

Function HasBar(Nm$) As Boolean
HasBar = HasVbeBar(CurVbe, Nm)
End Function

Function HasPjf(Pjf) As Boolean
HasPjf = HasPjfVbe(CurVbe, Pjf)
End Function

Function PjzPjf(A) As VBProject
Set PjzPjf = PjzPjfVbe(CurVbe, A)
End Function

Function MdDRszbe(A As Vbe, Optional WhStr$) As Drs
Set MdDRszbe = Drs(MdTblFny, MdDryzbe(A, WhStr))
End Function
Function MdTblFny() As String()

End Function

Sub CurVbeSav()
SavVbe CurVbe
End Sub
