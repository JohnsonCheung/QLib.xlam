Attribute VB_Name = "MIde_Cmp"
Option Explicit
Function Cmp(CmpNm$) As VBComponent
Set Cmp = CurPj.VBComponents(CmpNm)
End Function

Function PjzCmp(A As VBComponent) As VBProject
Set PjzCmp = A.Collection.Parent
End Function

Function HasCmpzPj(A As VBProject, CmpNm) As Boolean
If IsProtect(A) Then Exit Function
HasCmpzPj = HasItn(A.VBComponents, CmpNm)
End Function
Function PjNmzCmp$(A As VBComponent)
PjNmzCmp = PjzCmp(A).Name
End Function

Property Get CurCmp() As VBComponent
Set CurCmp = CurMd.Parent
End Property

Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function

Private Function HasCmpzPjTy(A As VBProject, Nm, Ty As vbext_ComponentType) As Boolean
Dim T As vbext_ComponentType
If Not HasItn(A.VBComponents, Nm) Then Exit Function
T = A.VBComponents(Nm).Type
If T = Ty Then HasCmpzPjTy = True: Exit Function
Thw CSub, "Pj has Cmp not as expected type", "PjNmzCmp EptTy ActTy", A.Name, Nm, ShtCmpTy(Ty), ShtCmpTy(T)
End Function


Function MdAyzCmp(A() As VBComponent) As CodeModule()
Dim I
For Each I In Itr(A)
    PushObj MdAyzCmp, CvCmp(I).CodeModule
Next
End Function

