Attribute VB_Name = "MDao_Li"
Option Explicit
Private Sub Z_LnkImp()
Dim Db As Database
Set Db = LnkImp(SampLiPm)
BrwFb Db.Name
Stop
End Sub
Function LnkImp(A As LiPm) As Database
ThwEr ChkColzLiPm(A), CSub
WOpn A.Apn
LnkTblz W, LtPm(A)
RunSqy W, ImpSqyzLi(A)
Set LnkImp = W
End Function
Private Sub Z_ChkColzLiPm()
Brw ChkColzLiPm(ShpCstLiPm)
End Sub
Private Function ChkColzLiPm(A As LiPm) As String()
ChkColzLiPm = MsgzLiMis(LiMis(A, LiAct(A)))
End Function

Private Function ImpSqyzLi(A As LiPm) As String()
PushIAy ImpSqyzLi, ImpSqyFb(A.Fb)
PushIAy ImpSqyzLi, ImpSqyFx(A.Fx)
End Function
Private Function ImpSqyFb(A() As LiFb) As String()
Dim J%
For J = 0 To UB(A)
    PushI ImpSqyFb, ImpSqlFb(A(J))
Next
End Function

Private Function ImpSqyFx(A() As LiFx) As String()
Dim J%
For J = 0 To UB(A)
    PushI ImpSqyFx, ImpSqlFx(A(J))
Next
End Function

Private Function ImpSqlFx$(A As LiFx)
Dim Fm$: Fm = ">" & A.T
Dim Into$: Into = "#I" & A.T
Dim Bexpr$: Bexpr = A.Bexpr
ImpSqlFx = SqlSel_FF_ExtNy_Into_Fm(A.Fny, A.ExtNy, Into, Fm, Bexpr)
End Function

Private Function ImpSqlFb$(A As LiFb)
With A
Dim FF$(): FF = .Fset.Sy
Dim Fm$: Fm = ">" & .T
Dim Into$: Into = "#I" & .T
Dim Bexpr$: Bexpr = .Bexpr
End With
ImpSqlFb = SqlSel_FF_Into_Fm(FF, Into, FF, Bexpr)
End Function

