Attribute VB_Name = "QIde_Pj_Cur"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Pj_Cur."
Private Const Asm$ = "QIde"

Property Get CPj() As VBProject
Set CPj = CVbe.ActiveVBProject
End Property


Function HasMd(P As VBProject, Mdn) As Boolean
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Name = Mdn Then HasMd = True: Exit Function
Next
End Function

Sub ThwIf_NotMod(M As CodeModule, Fun$)
If Not IsMod(M) Then Thw Fun, "Should be a Mod", "Mdn MdTy", Mdn(M), ShtCmpTy(CmpTyzM(M))
End Sub

Function HasMod(P As VBProject, Modn) As Boolean
If Not HasMd(P, Modn) Then Exit Function
ThwIf_NotMod MdzPN(P, Modn), CSub
End Function
Function PjnyzX(X As Excel.Application) As String()
PjnyzX = PjnyzV(X.Vbe)
End Function
Property Get PjnyX() As String()
PjnyX = PjnyzX(Xls)
End Property
Property Get Pjn$()
Pjn = CPj.Name
End Property
Sub BrwPjp()
BrwPth PjpP
End Sub
