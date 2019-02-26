Attribute VB_Name = "Module3"
Option Explicit
Type ReduceCol
    DRs As DRs
    ReduceColDic As Dictionary
End Type

Function ReduceCol(A As DRs) As ReduceCol
Dim Dry(): Dry = A.Dry
If Sz(Dry) = 0 Then Set ReduceCol.DRs = A:        Exit Function
Dim F$(): F = FnyzReducibleCol(A)
Dim Vy: Vy = DrzDrs(A, F)
Set ReduceCol.ReduceColDic = DiczKyVy(F, Vy)
Set ReduceCol.DRs = DrsDrpCC(A, F)
End Function
Private Function FnyzReducibleCol(A As DRs) As String()
Dim NCol%: NCol = NColzDrs(A)
Dim J%, Dry(), Fny$()
Fny = A.Fny
Dry = A.Dry
For J = 0 To NCol - 1
    If IsAllEleEqAy(ColzDry(Dry, J)) Then
        PushI FnyzReducibleCol, Fny(J)
    End If
Next
End Function
Sub BrwReduceCol(A As ReduceCol)
BrwAy FmtReduceCol(A)
End Sub

Private Function FmtReduceCol(A As ReduceCol) As String()
PushIAy FmtReduceCol, FmtDic(A.ReduceColDic)
PushIAy FmtReduceCol, FmtDrs(A.DRs)
End Function

Sub BrwDrszRedCol(A As DRs)
BrwReduceCol ReduceCol(A)
End Sub
