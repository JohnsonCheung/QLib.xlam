Attribute VB_Name = "MxFmtDic"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxFmtDic."
Private Sub Z_BrwDic()
Dim R As dao.Recordset
Set R = Rs(SampDboDutyDta, "Select Sku,BchNo from PermitD where BchNo<>''")
BrwDic JnStrDicTwoFldRs(R), True
End Sub

Sub BrwDic(A As Dictionary, Optional InclValTy As Boolean, Optional ExlIx As Boolean, Optional Fnn$, Optional OupTy As EmOupTy = EmOupTy.EiOtBrw)
BrwAy FmtDic(A, InclValTy), Fnn, OupTy:=OupTy
End Sub

Sub DmpDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val")
D FmtDic(A, InclDicValOptTy, Tit)
End Sub

Function S12szDiT1qLy(A As Dictionary) As S12s
Dim K
For Each K In A.Keys
    PushS12 S12szDiT1qLy, S12(K, JnCrLf(A(K)))
Next
End Function

Function FmtDicTit(A As Dictionary, Tit$) As String()
PushI FmtDicTit, Tit
PushI FmtDicTit, vbTab & "Count=" & A.Count
PushIAy FmtDicTit, AddPfxzAy(FmtDic(A, InclValTy:=True), vbTab)
End Function

Function FmtDic(A As Dictionary, Optional InclValTy As Boolean, Optional Nm1$ = "Key", Optional Nm2$ = "Val", Optional IxCol As EmIxCol) As String()
ThwIf_Nothing A, "Dic", CSub
Select Case True
Case IsDicSy(A):    FmtDic = FmtS12s(S12szDiT1qLy(A), Nm1, Nm2, IxCol)
Case IsDicLines(A): FmtDic = FmtS12s(S12szDic(A), Nm1, Nm2, IxCol)
Case Else:          FmtDic = FmtDiczLin(A, " ", InclValTy, Nm1, Nm2)
End Select
End Function

Private Function FmtDiczLin(A As Dictionary, Optional Sep$ = " ", Optional InclValTy As Boolean, Optional Nm1$, Optional Nm2$) As String()
If A.Count = 0 Then Exit Function
Dim Key: Key = A.Keys
Dim O$(): O = AlignAy(SyzItr(A.Keys))
Dim J&, I
For Each I In A.Items
   O(J) = O(J) & Sep & I
   J = J + 1
Next
FmtDiczLin = O
End Function
