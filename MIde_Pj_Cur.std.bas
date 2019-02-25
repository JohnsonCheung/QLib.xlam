Attribute VB_Name = "MIde_Pj_Cur"
Option Explicit

Property Get CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Property

Sub DltCmp(A As VBComponent)
A.Collection.Remove A
End Sub

Function EnsMd(MdNm$) As CodeModule
Set EnsMd = EnsMdvNm(CurPj, MdNm)
End Function
Function EnsMdvNm(A As VBProject, MdNm$) As CodeModule

End Function

Property Get PjNm$()
PjNm = CurPj.Name
End Property
Function PthPj$()
PthPj = PjPth(CurPj)
End Function
Sub BrwPjPth()
BrwPth PthPj
End Sub
