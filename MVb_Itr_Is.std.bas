Attribute VB_Name = "MVb_Itr_Is"
Option Explicit
Function IsItrzLines(Itr) As Boolean
If Not IsItrzStr(Itr) Then Exit Function
Dim I
For Each I In Itr
    If IsLines(I) Then IsItrzLines = True: Exit Function
Next
End Function
Function IsItrzStr(Itr) As Boolean
Dim I
For Each I In Itr
    If Not IsStr(I) Then Exit Function
Next
IsItrzStr = True
End Function

Function IsItrzNm(Itr) As Boolean
Dim I
For Each I In Itr
    If Not IsNm(I) Then Exit Function
Next
IsItrzNm = True
End Function

Function IsItrzSy(Itr) As Boolean
Dim I
For Each I In Itr
    If Not IsSy(I) Then Exit Function
Next
IsItrzSy = True
End Function
