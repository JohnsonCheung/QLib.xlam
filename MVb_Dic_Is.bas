Attribute VB_Name = "MVb_Dic_Is"
Option Explicit
Function IsDiczEmp(A As Dictionary) As Boolean
IsDiczEmp = A.Count = 0
End Function

Function IsDiczSy(A) As Boolean
Dim D As Dictionary, I, V
If Not IsDic(A) Then Exit Function
IsDiczSy = IsSyItr(CvDic(A).Keys)
End Function
Function IsDiczLines2(A As Dictionary) As Boolean
If Not IsAllStrAy(A.Keys) Then Exit Function
IsDiczLines2 = IsLinesItr(A.Items)
End Function

Function IsDiczLines(A As Dictionary) As Boolean
If Not IsDiczPrim(A) Then Exit Function
IsDiczLines = True
If IsLinesItr(A.Items) Then Exit Function
If IsLinesItr(A.Keys) Then Exit Function
IsDiczLines = False
End Function
Function IsDiczPrim(A As Dictionary) As Boolean
If Not IsPrimItr(A.Keys) Then Exit Function
IsDiczPrim = IsPrimItr(A.Items)
End Function
Function IsDiczStr(A As Dictionary) As Boolean
If Not IsStrItr(A.Keys) Then Exit Function
IsDiczStr = IsStrItr(A.Items)
End Function

Private Function IsDiczLines1(A As Dictionary) As Boolean
IsDiczLines1 = True
Dim K
For Each K In A.Keys
    If IsLines(K) Then Exit Function
    If IsLines(A(K)) Then Exit Function
Next
IsDiczLines1 = False
End Function

Function DicTy$(A As Dictionary)
Dim O$
Select Case True
Case IsDiczEmp(A):   O = "EmpDic"
Case IsDiczStr(A):   O = "StrDic"
Case IsDiczLines(A): O = "LineszDic"
Case IsDiczSy(A):    O = "SyDic"
Case Else:           O = "Dic"
End Select
End Function
