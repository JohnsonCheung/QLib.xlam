Attribute VB_Name = "MVb_Dic_Fmt"
Option Explicit
Private Sub Z_BrwDic()
Dim R As Dao.Recordset
Set R = Rs(SampDb_DutyDta, "Select Sku,BchNo from PermitD where BchNo<>''")
BrwDic JnStrDicTwoFldRs(R), True
End Sub

Sub BrwDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional UseVc As Boolean)
BrwAy FmtDic(A, InclDicValOptTy), UseVc:=True
End Sub

Sub DmpDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val")
D FmtDic(A, InclDicValOptTy, Tit)
End Sub
Function S1S2AyzSyDic(A As Dictionary) As S1S2()
Dim K
For Each K In A.Keys
    PushObj S1S2AyzSyDic, S1S2(K, JnCrLf(A(K)))
Next
End Function
Function FmtDic(A As Dictionary, Optional InclValTy As Boolean, Optional Nm1$ = "Key", Optional Nm2$ = "Val") As String()
If IsNothing(A) Then Exit Function
Select Case True
Case IsDiczSy(A):    FmtDic = FmtS1S2Ay(S1S2AyzSyDic(A), Nm1, Nm2)
Case IsDiczLines(A): FmtDic = FmtS1S2Ay(S1S2AyzDic(A), Nm1, Nm2)
Case Else:           FmtDic = FmtDic1(A)
End Select
End Function

Private Function FmtDic2(A As Dictionary) As String()
Dim K, O$(), J&
J = 1
For Each K In A.Keys
    PushI O, J & " " & K & " " & TypeName(A(K)) & " " & LinzVal(A(K))
    J = J + 1
Next
FmtDic2 = FmtAy2T(O)
End Function

Function FmtDic1(A As Dictionary, Optional Sep$ = " ") As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, W%, Ky
Ky = A.Keys
W = WdtzAy(Ky)
For Each K In Ky
   Push O, AlignL(K, W) & Sep & A(K)
Next
FmtDic1 = O
End Function


Function FmtDic3(A As Dictionary) As String()
Dim K
For Each K In A.Keys
    PushIAy FmtDic3, FmtDic4(K, A(K))
Next
End Function

Private Function FmtDic4(K, Lines) As String()
Dim L
For Each L In Itr(SplitCrLf(Lines))
    Push FmtDic4, K & " " & L
Next
End Function
