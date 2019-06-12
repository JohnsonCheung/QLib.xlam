Attribute VB_Name = "QVb_X"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_X."
Private Const Asm$ = "QVb"
Private A$()
Property Get XX() As String()
XX = A
Erase A
End Property
Sub XLy(Ly$())
PushIAy A, Ly
End Sub
Sub XLin(Optional Lin$)
PushI A, Lin
End Sub
Sub XBox(S$)
PushI A, Box(S)
End Sub
Sub XEnd()
PushI A, "End"
End Sub
Sub XDrs(Drs As Drs)
PushIAy A, FmtDrs(Drs)
End Sub
Sub X(V)
If IsArray(V) Then
    PushIAy A, V
Else
    PushI A, V
End If
End Sub

Function Y(A)
Stop
End Function
