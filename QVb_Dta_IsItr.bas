Attribute VB_Name = "QVb_Dta_IsItr"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Itr_Is."
Private Const Asm$ = "QVb"
Function AllTrue(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If Not P.Pred(I) Then Exit Function
Next
AllTrue = True
End Function
Function SomFTrue(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If P.Pred(I) Then SomFTrue = True: Exit Function
Next
End Function
Function AllFalse(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If P.Pred(I) Then Exit Function
Next
AllFalse = False
End Function

Function SomFalse(Itr, P As IPred) As Boolean
Dim I: For Each I In Itr
    If Not P.Pred(I) Then SomFalse = True: Exit Function
Next
End Function

Function PredIsLines() As IPred: Static X As New PredIsLines: Set PredIsLines = X: End Function
Function PredIsNm() As IPred:    Static X As New PredIsNm:    Set PredIsNm = X:    End Function
Function PredIsPrim() As IPred:  Static X As New PredIsPrim:  Set PredIsPrim = X:  End Function
Function PredIsStr() As IPred:   Static X As New PredIsStr:   Set PredIsStr = X:   End Function
Function PredIsSy() As IPred:    Static X As New PredIsSy:    Set PredIsSy = X:    End Function

Function IsItrStr(Itr) As Boolean:   IsItrStr = AllTrue(Itr, PredIsStr):   End Function
Function IsItrPrim(Itr) As Boolean:  IsItrPrim = AllTrue(Itr, PredIsPrim): End Function
Function IsItrNm(Itr) As Boolean:    IsItrNm = AllTrue(Itr, PredIsNm):     End Function
Function IsItrSy(Itr) As Boolean:    IsItrSy = AllTrue(Itr, PredIsSy):     End Function
Function IsItrLines(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsLines(V) Then Exit Function
Next
IsItrLines = True
End Function

Function PredzLikss(Likss$) As IPred
Dim O As New PredLikAy
O.Init SyzSS(Likss)
Set PredzLikss = O
End Function
