Attribute VB_Name = "QVb_Dic_Fmt"
Option Explicit
Private Const CMod$ = "MVb_Dic_Fmt."
Private Const Asm$ = "QVb"
Private Sub Z_BrwDic()
Dim R As DAO.Recordset
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
    PushS1S2 S1S2szSyDic, S1S2(K, JnCrLf(A(K)))
Next
End Function
Function FmtDicTit(A As Dictionary, Tit$) As String()
PushI FmtDicTit, Tit
PushI FmtDicTit, vbTab & "Count=" & A.Count
PushIAy FmtDicTit, AddPfxzSy(FmtDic(A, InclValTy:=True), vbTab)
End Function

Function FmtDic(A As Dictionary, Optional InclValTy As Boolean, Optional Nm1$ = "Key", Optional Nm2$ = "Val", Optional AddIx As Boolean) As String()
ThwIfNothing A, CSub
Select Case True
Case IsDiczSy(A):    FmtDic = FmtS1S2s(S1S2szSyDic(A), Nm1, Nm2)
Case IsDiczLines(A): FmtDic = FmtS1S2s(S1S2szDic(A), Nm1, Nm2)
Case Else:           FmtDic = FmtDiczLin(A, " ", InclValTy, Nm1, Nm2)
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

Function FmtDiczLin(A As Dictionary, Optional Sep$ = " ", Optional InclValTy As Boolean, Optional Nm1$, Optional Nm2$) As String()
If A.Count = 0 Then Exit Function
Dim Key: Key = A.Keys
Dim O$(): O = AlignLzSy(SyzItr(A.Keys))
Dim J&, I
For Each I In A.Items
   O(J) = O(J) & Sep & I
   J = J + 1
Next
FmtDiczLin = O
End Function
