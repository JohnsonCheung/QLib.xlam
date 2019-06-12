Attribute VB_Name = "QDta_Drs_ReducedDrs"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_ReducedDrs."
Private Const Asm$ = "QDta"
Type ReducedDrs
    Drs As Drs
    ReducedColDic As Dictionary
End Type

Function ReducedDrs(A As Drs) As ReducedDrs
If NoReczDrs(A) Then GoTo X
Dim C$(): C = ReducibleCny(A)
If Si(C) = 0 Then GoTo X
Dim Ixy&(): Ixy = IxyzSubAy(A.Fny, C)
Dim Dr: Dr = A.Dry(0)
Dim Vy: Vy = AywIxy(Dr, Ixy)
Set ReducedDrs.ReducedColDic = DiczKyVy(C, Vy)
ReducedDrs.Drs = DrpColzFny(A, C)
Exit Function
X:
    ReducedDrs.Drs = A
    Set ReducedDrs.ReducedColDic = New Dictionary
End Function
Private Function ReducibleCny(A As Drs) As String() '
'Ret : ColNy ! if any col in Drs-A has all sam val, this col is reduciable.  Return them
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
Sub BrwDrszReduce(A As ReducedDrs)
BrwAy FmtDrszReduce(A)
End Sub

Private Function FmtDrszReduce(A As ReducedDrs) As String()
PushIAy FmtDrszReduce, FmtDic(A.ReducedColDic)
PushIAy FmtDrszReduce, FmtDrs(A.Drs)
End Function
