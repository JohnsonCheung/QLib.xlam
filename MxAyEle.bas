Attribute VB_Name = "MxAyEle"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyEle."
Function LasSndEle(Ay)
Dim N&: N = Si(Ay)
If N <= 1 Then
    ThwMsg CSub, "Only 1 or no ele in Ay"
Else
    Asg Ay(N - 2), LasSndEle
End If
End Function

Function LasEle(Ay)
Dim N&: N = Si(Ay)
If N = 0 Then
    ThwMsg CSub, "No ele in Ay"
Else
    Asg Ay(N - 1), LasEle
End If
End Function
