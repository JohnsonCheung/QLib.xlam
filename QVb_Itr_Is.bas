Attribute VB_Name = "QVb_Itr_Is"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Itr_Is."
Private Const Asm$ = "QVb"
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

Function PredIsLines() As IPred: Static X As New PredIsLines: Set PredIsLines = X: End Function
Function PredIsNm() As IPred:    Static X As New PredIsNm:    Set PredIsNm = X:    End Function
Function PredIsPrim() As IPred:  Static X As New PredIsPrim:  Set PredIsPrim = X:  End Function
Function PredIsStr() As IPred:   Static X As New PredIsStr:   Set PredIsStr = X:   End Function
Function PredIsSy() As IPred:    Static X As New PredIsSy:    Set PredIsSy = X:    End Function

Function IsItrOfStr(Itr) As Boolean:   IsItrOfStr = IsAllTrue(Itr, PredIsStr):   End Function
Function IsItrOfPrim(Itr) As Boolean:  IsItrOfPrim = IsAllTrue(Itr, PredIsPrim): End Function
Function IsItrOfNm(Itr) As Boolean:    IsItrOfNm = IsAllTrue(Itr, PredIsNm):     End Function
Function IsItrOfSy(Itr) As Boolean:    IsItrOfSy = IsAllTrue(Itr, PredIsSy):     End Function
Function IsItrOfLines(Itr) As Boolean: IsItrOfLines = IsItrOfStr(Itr) And IsSomTrue(Itr, PredIsLines): End Function

