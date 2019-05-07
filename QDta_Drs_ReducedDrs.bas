Attribute VB_Name = "QDta_Drs_ReducedDrs"
Option Explicit
Private Const CMod$ = "MDta_ReducedDrs."
Private Const Asm$ = "QDta"
Type ReducedDrs
    Drs As Drs
    ReducedColDic As Dictionary
End Type

Function ReducedDrs(A As Drs) As ReducedDrs
Dim Dry(): Dry = A.Dry
If Si(Dry) = 0 Then Set ReducedDrs.Drs = A:        Exit Function
Dim C$(): C = ReducibleCny(A)
Dim Vy: Vy = DrvzDrs(A, C, 0)
Set ReducedDrs.ReducedColDic = DiczKyVy(C, Vy)
ReducedDrs.Drs = DrpCny(A, C)
End Function
Private Function ReducibleCny(A As Drs) As String()
Dim NCol%: NCol = NColzDrs(A)
Dim J%, Dry(), Fny$()
Fny = A.Fny
Dry = A.Dry
For J = 0 To NCol - 1
    If IsEqzAllEle(ColzDry(Dry, J)) Then
        PushI ReducibleCny, Fny(J)
    End If
Next
End Function
Sub BrwReducedDrs(A As ReducedDrs)
BrwAy FmtReducedDrs(A)
End Sub

Private Function FmtReducedDrs(A As ReducedDrs) As String()
PushIAy FmtReducedDrs, FmtDic(A.ReducedColDic)
PushIAy FmtReducedDrs, FmtDrs(A.Drs)
End Function
