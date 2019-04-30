Attribute VB_Name = "MVb_Dic_Fmt"
Option Explicit
Private Sub Z_BrwDic()
Dim R As Dao.Recordset
Set R = Rs(SampDbzDutyDta, "Select Sku,BchNo from PermitD where BchNo<>''")
BrwDic JnStrDicTwoFldRs(R), True
End Sub

Sub BrwDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional UseVc As Boolean, Optional AddIx As Boolean)
Dim B As Dictionary
If AddIx Then
    Set B = DicAddIxToKey(A)
Else
    Set B = A
End If
BrwAy FmtDic(B, InclDicValOptTy), UseVc:=UseVc
End Sub

Sub DmpDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val")
D FmtDic(A, InclDicValOptTy, Tit)
End Sub

Function S1S2szSyDic(A As Dictionary) As S1S2s
Dim K
For Each K In A.Keys
    PushObj S1S2szSyDic, S1S2(K, JnCrLf(A(K)))
Next
End Function
Function FmtDicTit(A As Dictionary, Tit$) As String()
PushI FmtDicTit, Tit
PushI FmtDicTit, vbTab & "Count=" & A.Count
PushIAy FmtDicTit, SyAddPfx(FmtDic(A, InclValTy:=True), vbTab)
End Function

Function FmtDic(A As Dictionary, Optional InclValTy As Boolean, Optional Nm1$ = "Key", Optional Nm2$ = "Val", Optional AddIx As Boolean) As String()
If IsNothing(A) Then Exit Function
Select Case True
Case IsDiczSy(A):    FmtDic = FmtS1S2s(S1S2szSyDic(A), Nm1, Nm2)
Case IsDiczLines(A): FmtDic = FmtS1S2s(S1S2szDic(A), Nm1, Nm2)
Case Else:           FmtDic = FmtDic1(A)
End Select
End Function

Private Function FmtDic2(A As Dictionary) As String()
Dim K, O$(), J&
J = 1
For Each K In A.Keys
    PushI O, J & " " & K & " " & TypeName(A(K)) & " " & StrCellzV(A(K))
    J = J + 1
Next
FmtDic2 = FmtSyT2(O)
End Function
Function FmtDic1(A As Dictionary, Optional Sep$ = " ") As String()
If A.Count = 0 Then Exit Function
Dim Key: Key = A.Keys
Dim O$(): O = FmtAySamWdt(SyzItr(A.Keys))
Dim J&, I
For Each I In A.Items
   O(J) = O(J) & Sep & I
   J = J + 1
Next
FmtDic1 = O
End Function

Function FmtDic3(A As Dictionary) As String()
Dim K
For Each K In A.Keys
    PushIAy FmtDic3, FmtDic4(K, A(K))
Next
End Function

Private Function FmtDic4(K, Lines$) As String()
Dim L
For Each L In Itr(SplitCrLf(Lines))
    Push FmtDic4, K & " " & L
Next
End Function
