Attribute VB_Name = "MIde_Cnt_Cmp"
Option Explicit
Property Get NCls%()
NCls = NClszPj(CurPj)
End Property
Property Get NCmpPj%()
NCmpPj = CurPj.VBComponents.Count
End Property


Property Get NModPj%()
NModPj = NModzPj(CurPj)
End Property
'===============================================
Function LockedCmpCnt() As CmpCnt
Static X As New CmpCnt
X.Locked = True
Set LockedCmpCnt = X
End Function
Function CmpCnt(A As VBProject) As CmpCnt
If A.Protection = vbext_pp_locked Then Set CmpCnt = LockedCmpCnt: Exit Function
Set CmpCnt = New CmpCnt
CmpCnt.Init NModzPj(A), NClszPj(A), NDocOfPj(A), NOthCmpzPj(A)
End Function
'----------------------------------------------
Property Get CmpCntPj() As CmpCnt
Set CmpCntPj = CmpCnt(CurPj)
End Property
'==============================================
Function NCmpzPj%(A As VBProject)
If A.Protection = vbext_pp_locked Then Exit Function
NCmpzPj = A.VBComponents.Count
End Function

Function NModzPj%(Pj As VBProject)
NModzPj = NCmpzTy(Pj, vbext_ct_StdModule)
End Function

Function NDocOfPj%(A As VBProject)
NDocOfPj = NCmpzTy(A, vbext_ct_Document)
End Function

Function NClszPj%(A As VBProject)
NClszPj = NCmpzTy(A, vbext_ct_ClassModule)
End Function


Function NCmpzTy%(A As VBProject, Ty As vbext_ComponentType)
If A.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
Dim O%
For Each C In A.VBComponents
    If C.Type = Ty Then O = O + 1
Next
NCmpzTy = O
End Function

Function NOthCmpzPj%(A As VBProject)
NOthCmpzPj = NCmpzPj(A) - NClszPj(A) - NModzPj(A) - NDocOfPj(A)
End Function

