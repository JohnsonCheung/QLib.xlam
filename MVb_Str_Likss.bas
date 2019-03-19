Attribute VB_Name = "MVb_Str_Likss"
Option Explicit
Function StrLikss(A, Likss) As Boolean
StrLikss = StrLikAy(A, SySsl(Likss))
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

Private Sub Z_T1zT1LikTLinAy()
Dim A$(), Nm$
GoSub T1
GoSub T2
Exit Sub
T1:
    A = SplitVbar("a bb* *dd | c x y")
    Nm = "x"
    Ept = "c"
    GoTo Tst
T2:
    A = SplitVbar("a bb* *dd | c x y")
    Nm = "bb1"
    Ept = "a"
    GoTo Tst
Tst:
    Act = T1zT1LikTLinAy(A, Nm)
    C
    Return
End Sub

Function T1zT1LikTLinAy$(T1LikTLinAy$(), Nm)
Dim L, T1$
If Si(T1LikTLinAy) = 0 Then Exit Function
For Each L In T1LikTLinAy
    T1 = ShfT(L)
    If StrLikss(Nm, L) Then
        T1zT1LikTLinAy = T1
        Exit Function
    End If
Next
End Function


Private Sub Z()
Z_T1zT1LikTLinAy
MVb_Str_Likss:
End Sub
