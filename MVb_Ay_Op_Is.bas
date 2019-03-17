Attribute VB_Name = "MVb_Ay_Op_Is"
Option Explicit
Function IsAllEleHasVyzDicKK(A) As Boolean
If Si(A) = 0 Then Exit Function
Dim I
For Each I In A
    If IsEmp(I) Then Exit Function
Next
IsAllEleHasVyzDicKK = True
End Function

Function IsAllEleEqAy(A) As Boolean
If Si(A) <= 1 Then IsAllEleEqAy = True: Exit Function
Dim A0, J&
A0 = A(0)
For J = 1 To UB(A)
    If A0 <> A(J) Then Exit Function
Next
IsAllEleEqAy = True
End Function

Function IsAllStrAy(A) As Boolean
If Not IsArray(A) Then Exit Function
If IsSy(A) Then IsAllStrAy = True: Exit Function
Dim I
For Each I In Itr(A)
    If Not IsStr(I) Then Exit Function
Next
IsAllStrAy = True
End Function

Function IsEqSz(A, B) As Boolean
IsEqSz = Si(A) = Si(B)
End Function

Function IsEqAy(A, B) As Boolean
If Not IsArray(A) Then Exit Function
If Not IsArray(B) Then Exit Function
If Not IsEqSz(A, B) Then Exit Function
Dim J&, X
For Each X In Itr(A)
    If Not IsEq(X, B(J)) Then Exit Function
    J = J + 1
Next
IsEqAy = True
End Function

Function IsSamAy(A, B) As Boolean
IsSamAy = IsEqDic(CntDic(A), CntDic(B))
End Function

Private Sub Z()
End Sub
