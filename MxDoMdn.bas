Attribute VB_Name = "MxDoMdn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMdn."
Public Const FFoMdn$ = "Pjn MdTy Mdn"
Private Function FoMdn() As String(): FoMdn = SyzSS(FFoMdn): End Function
Function DroMdn(M As CodeModule) As Variant()
DroMdn = Array(PjnzM(M), ShtMdTy(M), Mdn(M))
End Function

Function DoMdnP() As Drs
DoMdnP = DoMdnzP(CPj)
End Function

Function DoMdnzP(P As VBProject) As Drs
DoMdnzP = Drs(FoMdn, DyoMdnzP(P))
End Function

Function DyoMdnzP(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    Push DyoMdnzP, DroMdn(C.CodeModule)
Next
End Function