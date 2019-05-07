Attribute VB_Name = "QVb_Str_Filter"
Option Explicit
Private Const CMod$ = "MVb_Str_Filter."
Private Const Asm$ = "QVb"
Function Filter(S, Src$()) As String()
Dim I
For Each I In Itr(Src)
    If Hit(S, I) Then PushI Filter, I
Next
End Function
Private Function Hit(S, Src) As Boolean
Dim P%, J%, P1%
For J = 1 To Len(S)
    P1 = InStr(P, Src, Mid(S, J, 1))
    If P1 = 0 Then Exit Function
    P = P1 + 1
Next
Hit = True
End Function
