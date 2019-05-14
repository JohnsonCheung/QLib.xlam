Attribute VB_Name = "QIde_Md_Itr"
Option Explicit
Private Const CMod$ = "MIde_Md_Itr."
Private Const Asm$ = "QIde"
Function ModItr()
Dim C As VBComponent, O() As CodeModule
For Each C In CPj.VBComponents
    PushObj O, C.CodeModule
Next
Asg Itr(O), ModItr
End Function
