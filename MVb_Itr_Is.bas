Attribute VB_Name = "MVb_Itr_Is"
Option Explicit
Function IsAllTrue(Itr, P As IPred) As Boolean
Dim I
For Each I In Itr
    If Not P.Pred(I) Then Exit Function
Next
IsAllTrue = True
End Function
Function IsSomTrue(Itr, P As IPred) As Boolean
Dim I
For Each I In Itr
    If P.Pred(I) Then IsSomTrue = True: Exit Function
Next
End Function
Function IsAllFalse(Itr, P As IPred) As Boolean
Dim I
For Each I In Itr
    If P.Pred(I) Then Exit Function
Next
IsAllFalse = False
End Function
Function IsSomFalse(Itr, P As IPred) As Boolean
Dim I
For Each I In Itr
    If Not P.Pred(I) Then IsSomFalse = True: Exit Function
Next
End Function

Function IsLinesPred() As IPred: Static X As New PredzIsLines: Set IsLinesPred = X: End Function
Function IsNmPred() As IPred:    Static X As New PredzIsNm:   Set IsNmPred = X:   End Function
Function IsPrimPred() As IPred:  Static X As New PredzIsPrim: Set IsPrimPred = X: End Function
Function IsStrPred() As IPred:   Static X As New PredzIsStr:  Set IsStrPred = X:  End Function
Function IsSyPred() As IPred:    Static X As New PredzIsSy:   Set IsSyPred = X:  End Function

Function IsStrItr(Itr) As Boolean:   IsStrItr = IsAllTrue(Itr, IsStrPred):   End Function
Function IsPrimItr(Itr) As Boolean:  IsPrimItr = IsAllTrue(Itr, IsPrimPred): End Function
Function IsNmItr(Itr) As Boolean:    IsNmItr = IsAllTrue(Itr, IsNmPred):     End Function
Function IsSyItr(Itr) As Boolean:    IsSyItr = IsAllTrue(Itr, IsSyPred):     End Function
Function IsLinesItr(Itr) As Boolean: IsLinesItr = IsStrItr(Itr) And IsSomTrue(Itr, IsLinesPred): End Function

