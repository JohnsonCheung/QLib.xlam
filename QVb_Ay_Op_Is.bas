Attribute VB_Name = "QVb_Ay_Op_Is"
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Is."
Private Const Asm$ = "QVb"
Function IsAllEleHasVyzDicKK(A) As Boolean
If Si(A) = 0 Then Exit Function
Dim I
For Each I In A
    If IsEmp(I) Then Exit Function
Next
IsAllEleHasVyzDicKK = True
End Function

Function IsEqzAllEle(Ay) As Boolean
If Si(Ay) <= 1 Then IsEqzAllEle = True: Exit Function
Dim A0, J&
A0 = Ay(0)
For J = 1 To UB(Ay)
    If A0 <> Ay(J) Then Exit Function
Next
IsEqzAllEle = True
End Function

Function IsAllStrAy(Ay) As Boolean
If Not IsArray(Ay) Then Exit Function
If IsSy(Ay) Then IsAllStrAy = True: Exit Function
Dim I
For Each I In Itr(Ay)
    If Not IsStr(I) Then Exit Function
Next
IsAllStrAy = True
End Function

Function IsEqSi(A, B) As Boolean
IsEqSi = Si(A) = Si(B)
End Function

Function IsEqAy(A, B) As Boolean
If Not IsArray(A) Then Exit Function
If Not IsArray(B) Then Exit Function
If Not IsEqSi(A, B) Then Exit Function
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
