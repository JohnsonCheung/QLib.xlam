Attribute VB_Name = "MIde_Pj_Cur"
Option Explicit

Property Get CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Property

Function EnsMd(MdNm$) As CodeModule
Set EnsMd = EnsMdzPj(CurPj, MdNm)
End Function

Function EnsModzPj(A As VBProject, ModNm$) As CodeModule
If Not HasMd(A, ModNm) Then AddModzPj A, ModNm
Set EnsModzPj = MdzPj(A, ModNm)
End Function

Function HasMd(A As VBProject, MdNm) As Boolean
Dim C As VBComponent
For Each C In A.VBComponents
    If C.Name = MdNm Then HasMd = True: Exit Function
Next
End Function

Sub ThwIfNotMod(A As CodeModule, Fun$)
If Not IsMod(A) Then Thw Fun, "Should be a Mod", "MdNm MdTy", MdNm(A), ShtCmpTy(CmpTyzMd(A))
End Sub

Function HasMod(A As VBProject, ModNm) As Boolean
If Not HasMd(A, ModNm) Then Exit Function
ThwIfNotMod MdzPj(A, ModNm), CSub
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
