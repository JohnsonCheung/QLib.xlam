Attribute VB_Name = "QVb_X"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_X."
Private Const Asm$ = "QVb"
Private A$()
Function XX() As String()
XX = A
Erase A
End Function

Sub XBox(S$)
X Box(S)
End Sub
Sub XEnd()
PushI A, "End"
End Sub
Sub XLin(Optional L$)
PushI A, L
End Sub
Sub XDrs(Drs As Drs)
PushIAy A, LinzDrs(Drs)
End Sub
Sub XTab(V)
If IsArray(V) Then
    X TabAy(V)
Else
    X vbTab & V
End If
End Sub
Sub X(V)
If IsArray(V) Then
    PushIAy A, V
Else
    PushI A, V
End If
End Sub
