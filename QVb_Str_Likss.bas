Attribute VB_Name = "QVb_Str_Likss"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Likss."
Private Const Asm$ = "QVb"
Function StrLikss(A, Likss) As Boolean
StrLikss = StrLikAy(A, SyzSS(Likss))
End Function

Function StrLikAy(A, LikAy$()) As Boolean
Dim I
For Each I In Itr(LikAy)
    If A Like I Then StrLikAy = True: Exit Function
Next
End Function

Function StrLikssAy(A, LikssAy) As Boolean
Dim Likss
For Each Likss In LikssAy
    If StrLikss(A, Likss) Then StrLikssAy = True: Exit Function
Next
End Function

Private Sub Z_T1zT1LikTLiny()
Dim A$(), Nm$
GoSub T1
GoSub T2
Exit Sub
T1:
    A = SplitVBar("a bb* *dd | c x y")
    Nm = "x"
    Ept = "c"
    GoTo Tst
T2:
    A = SplitVBar("a bb* *dd | c x y")
    Nm = "bb1"
    Ept = "a"
    GoTo Tst
Tst:
    Act = T1zT1LikTLiny(A, Nm)
    C
    Return
End Sub

Function T1zT1LikTLiny$(T1LikTLiny$(), Nm)
Dim L, T1$
If Si(T1LikTLiny) = 0 Then Exit Function
For Each L In T1LikTLiny
    'T1 = ShfT1(L)
    If StrLikss(Nm, L) Then
        T1zT1LikTLiny = T1
        Exit Function
    End If
Next
End Function


Private Sub ZZ()
Z_T1zT1LikTLiny
MVb_Str_Likss:
End Sub
