Attribute VB_Name = "QDta_Drs_RduDrs"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_RduDrs."
Private Const Asm$ = "QDta"
Private Type RduDrs  ' #Reduced-Drs ! if a drs col all val are sam, mov those cols to @RduColDic (Dic-of-coln-to-val).
    Drs As Drs       '              ! the drs aft rmv the sam val col
    RduColDic As Dictionary '        ! one entry is one col.  Key is coln and val is coln val.
End Type

Private Function RduDrs(A As Drs) As RduDrs
'Ret : @A as :t:RduDrs
If NoReczDrs(A) Then GoTo X
Dim C$(): C = ReducibleCny(A)
If Si(C) = 0 Then GoTo X
Dim Ixy&(): Ixy = IxyzSubAy(A.Fny, C)
Dim Dr: Dr = A.Dy(0)
Dim Vy: Vy = AwIxy(Dr, Ixy)
Set RduDrs.RduColDic = DiczKyVy(C, Vy)
RduDrs.Drs = DrpColFny(A, C)
Exit Function
X:
    RduDrs.Drs = A
    Set RduDrs.RduColDic = New Dictionary
End Function
Private Function ReducibleCny(A As Drs) As String() '
'Ret : ColNy ! if any col in Drs-A has all sam val, this col is reduciable.  Return them
Dim NCol%: NCol = NColzDrs(A)
Dim J%, Dy(), Fny$()
Fny = A.Fny
Dy = A.Dy
For J = 0 To NCol - 1
    If IsEqzAllEle(ColzDy(Dy, J)) Then
        PushI ReducibleCny, Fny(J)
    End If
Next
End Function

Sub BrwDrsRdu(A As Drs)
'Ret : Brw @A in reduced fmt @@
BrwAy FmtRduDrs(RduDrs(A))
End Sub
Function FmtDrszRdu(A As Drs) As String()
FmtDrszRdu = FmtRduDrs(RduDrs(A))
End Function
Private Function FmtRduDrs(A As RduDrs) As String()
PushIAy FmtRduDrs, FmtDic(A.RduColDic)
PushIAy FmtRduDrs, FmtDrs(A.Drs)
End Function
