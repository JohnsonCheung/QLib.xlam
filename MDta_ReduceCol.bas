Attribute VB_Name = "MDta_ReduceCol"
Option Explicit
Type ReduceCol
    Drs As Drs
    ReduceColDic As Dictionary
End Type

Function ReduceCol(A As Drs) As ReduceCol
Dim Dry(): Dry = A.Dry
If Si(Dry) = 0 Then Set ReduceCol.Drs = A:        Exit Function
Dim F$(): F = FnyzReducibleCol(A)
Dim Vy: Vy = DrzDrs(A, F)
Set ReduceCol.ReduceColDic = DiczKyVy(F, Vy)
Set ReduceCol.Drs = DrsDrpCC(A, F)
End Function
Private Function FnyzReducibleCol(A As Drs) As String()
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
PushIAy FmtReduceCol, FmtDrs(A.Drs)
End Function

Sub BrwDrszRedCol(A As Drs)
BrwReduceCol ReduceCol(A)
End Sub
