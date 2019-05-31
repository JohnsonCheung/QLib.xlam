Attribute VB_Name = "QVb_Is"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Is."
Private Const Asm$ = "QVb"

Function IsBet(V, A, B) As Boolean
If A > V Then Exit Function
If V > B Then Exit Function
IsBet = True
End Function
Function IsErObj(A) As Boolean
IsErObj = TypeName(A) = "Error"
End Function
Function IsEmp(A) As Boolean
Select Case True
Case IsStr(A):    IsEmp = Trim(A) = ""
Case IsArray(A):  IsEmp = Si(A) = 0
Case IsEmpty(A), IsNothing(A), IsMissing(A), IsNull(A): IsEmp = True
End Select
End Function

Function IsNBet(V, A, B) As Boolean
IsNBet = Not IsBet(V, A, B)
End Function

Function IsSqBktQuoted(S) As Boolean
IsSqBktQuoted = IsQuoted(S, "[", "]")
End Function

Function Limit(V, A, B)
Select Case V
Case V > B: Limit = B
Case V < A: Limit = A
Case Else: Limit = V
End Select
End Function
