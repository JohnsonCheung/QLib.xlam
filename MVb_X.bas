Attribute VB_Name = "MVb_X"
Option Explicit
Public xx$()
Sub X(S$)
Push xx, S
End Sub
Sub X0(S$)
If Si(xx) = 0 Then PushI xx, S: Exit Sub
X1 S
End Sub
Sub X1(S$)
Dim U&: U = UB(xx)
xx(U) = xx(U) & S
End Sub
