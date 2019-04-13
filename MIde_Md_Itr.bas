Attribute VB_Name = "MIde_Md_Itr"
Option Explicit
Function ModItr()
Dim C As VBComponent, O() As CodeModule
For Each C In CurPj.VBComponents
    PushObj O, C.CodeModule
Next
Asg Itr(O), ModItr
End Function
