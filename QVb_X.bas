Attribute VB_Name = "QVb_X"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_X."
Private Const Asm$ = "QVb"
Public XX$()

Sub X(V)
If IsArray(V) Then
    PushIAy XX, V
Else
    PushI XX, V
End If
End Sub

Function Y(A)
Stop
End Function
