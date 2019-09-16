Attribute VB_Name = "MxLikss"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxLikss."
Function IsLikzSS(A, Kss) As Boolean
IsLikzSS = IsLikzAy(A, SyzSS(Kss))
End Function

Function IsLikzAy(A, LikAy$()) As Boolean
Dim I
For Each I In Itr(LikAy)
    If A Like I Then IsLikzAy = True: Exit Function
Next
End Function

Function IsLikzSSAy(A, KssAy) As Boolean
Dim Kss
For Each Kss In KssAy
    If IsLikzSS(A, Kss) Then IsLikzSSAy = True: Exit Function
Next
End Function

Private Sub Z_T1zTkssLy()
Dim A$(), NM$
GoSub T1
GoSub T2
Exit Sub
T1:
    A = SplitVBar("a bb* *dd | c x y")
    NM = "x"
    Ept = "c"
    GoTo Tst
T2:
    A = SplitVBar("a bb* *dd | c x y")
    NM = "bb1"
    Ept = "a"
    GoTo Tst
Tst:
    Act = T1zTkssLy(A, NM)
    C
    Return
End Sub

Function T1zTkssLy$(TkssLy$(), NM)
':Tkss: :SS #T1-Likss# ! It is SS with T1 and Likss
':Kss:  :SS #Likss#    ! It is SS with each term is LikStr
Dim L: For Each L In Itr(TkssLy)
    Dim T1$: T1 = ShfT1(L)
    If IsLikzSS(NM, L) Then
        T1zTkssLy = T1
        Exit Function
    End If
Next
End Function

Private Sub Z()
Z_T1zTkssLy
MVb_Str_Kss:
End Sub