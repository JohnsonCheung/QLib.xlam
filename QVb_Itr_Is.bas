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

Function PredzIsLines() As IPred: Static X As New PredzIsLines: Set PredzIsLines = X: End Function
Function PredzIsNm() As IPred:    Static X As New PredzIsNm:    Set PredzIsNm = X:    End Function
Function PredzIsPrim() As IPred:  Static X As New PredzIsPrim:  Set PredzIsPrim = X:  End Function
Function PredzIsStr() As IPred:   Static X As New PredzIsStr:   Set PredzIsStr = X:   End Function
Function PredzIsSy() As IPred:    Static X As New PredzIsSy:    Set PredzIsSy = X:    End Function

Function IsItrOfStr(Itr) As Boolean:   IsItrOfStr = IsAllTrue(Itr, PredzIsStr):   End Function
Function IsItrOfPrim(Itr) As Boolean:  IsItrOfPrim = IsAllTrue(Itr, PredzIsPrim): End Function
Function IsItrOfNm(Itr) As Boolean:    IsItrOfNm = IsAllTrue(Itr, PredzIsNm):     End Function
Function IsItrOfSy(Itr) As Boolean:    IsItrOfSy = IsAllTrue(Itr, PredzIsSy):     End Function
Function IsItrOfLines(Itr) As Boolean: IsItrOfLines = IsItrOfStr(Itr) And IsSomTrue(Itr, PredzIsLines): End Function

