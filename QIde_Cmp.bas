Attribute VB_Name = "QIde_Cmp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmp."
Private Const Asm$ = "QIde"
Function PjzC(A As VBComponent) As VBProject
Set PjzC = A.Collection.Parent
End Function
Function HasCmp(Cmpn) As Boolean
HasCmp = HasCmpzP(CPj, Cmpn)
End Function
Function HasCmpzP(P As VBProject, Cmpn) As Boolean
If IsProtectzvInf(P) Then Exit Function
HasCmpzP = HasItn(P.VBComponents, Cmpn)
End Function
Function PjnzC$(A As VBComponent)
PjnzC = PjzC(A).Name
End Function

Property Get CCmp() As VBComponent
Set CCmp = CMd.Parent
End Property

Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function

Private Function HasCmpzPTN(P As VBProject, Ty As vbext_ComponentType, Cmpn) As Boolean
Dim T As vbext_ComponentType
If Not HasCmpzP(P, Cmpn) Then Exit Function
T = CmpTyzPN(P, Cmpn)
If T = Ty Then HasCmpzPTN = True: Exit Function
Thw CSub, "Pj has Cmp not as expected type", "PjnzC EptTy ActTy", P.Name, Cmpn, ShtCmpTy(Ty), ShtCmpTy(T)
End Function

Function MdAyzC(CmpAy() As VBComponent) As CodeModule()
Dim I
For Each I In Itr(CmpAy)
    PushObj MdAyzC, CvCmp(I).CodeModule
Next
End Function

