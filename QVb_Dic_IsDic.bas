Attribute VB_Name = "QVb_Dic_IsDic"
Option Explicit
Private Const CMod$ = "MVb_Dic_Is."
Private Const Asm$ = "QVb"
Function IsEmpDic(A As Dictionary) As Boolean
IsEmpDic = A.Count = 0
End Function
Function TyNmAy(Ay) As String()
Dim V
For Each V In Itr(Ay)
    PushI TyNmAy, TypeName(V)
Next
End Function
Function VyzDKy(D As Dictionary, Ky) As Variant()
Dim K
For Each K In Itr(Ky)
    If Not D.Exists(K) Then Thw CSub, "Some N in Ny not found in Dic.Keys", "[N in Ny not fnd in Dic.Keys] DicKeys Ky", K, AvzItr(D.Keys), Ky
    Push VyzDKy, D(K)
Next
End Function
Function DicwKy(D As Dictionary, Ky) As Dictionary
Set DicwKy = New Dictionary
Dim Vy(): Vy = VyzDKy(D, Ky)
Dim K, J&
For Each K In Itr(Ky)
    DicwKy.Add K, Vy(J)
    J = J + 1
Next
End Function
Function Vy(A As Dictionary) As Variant()
Vy = IntozItr(EmpAv, A.Items)
End Function
Function TyNmAyzDic(A As Dictionary) As String()
TyNmAyzDic = TyNmAy(Vy(A))
End Function

Function IsDicOfSy(A As Dictionary) As Boolean
Dim D As Dictionary, I, V
If Not IsDic(A) Then Exit Function
IsDicOfSy = IsItrOfSy(CvDic(A).Items)
End Function

Function IsDicOfLines(A As Dictionary) As Boolean
IsDicOfLines = True
If IsItrOfLines(A.Items) Then Exit Function
If IsItrOfStr(A.Keys) Then Exit Function
IsDicOfLines = False
End Function
Function IsDicOfPrim(A As Dictionary) As Boolean
If Not IsItrOfPrim(A.Keys) Then Exit Function
IsDicOfPrim = IsItrOfPrim(A.Items)
End Function
Function IsDicOfStr(A As Dictionary) As Boolean
If Not IsItrOfStr(A.Keys) Then Exit Function
IsDicOfStr = IsItrOfStr(A.Items)
End Function

Function DicTy$(A As Dictionary)
Dim O$
Select Case True
Case IsEmpDic(A):   O = "EmpDic"
Case IsDicOfStr(A):   O = "StrDic"
Case IsDicOfLines(A): O = "LineszDic"
Case IsDicOfSy(A):    O = "SyDic"
Case Else:           O = "Dic"
End Select
End Function
