Attribute VB_Name = "MDao_Lid"
Option Explicit

Private Sub CpyFilzToWPth(A As LidPm)
CpyFilzIfDif SyzOyPrp(A.Fil, "Ffn"), WPth(A.Apn)
End Sub

Sub LnkImpzLidPm(A As LidPm)
ThwEr ErzLidPmzV1(A), CSub
WIniOpn A.Apn
CpyFilzIfDif SyzOyPrp(A.Fil, "Ffn"), WPth(A.Apn)
ThwEr ErzLnkTblzLtPm(W, LtPmzLid(A)), CSub
RunSqy W, ImpSqyzLidPm(A)
WCls
End Sub

Private Function ErzLidPm(A As LidPm) As String()
ErzLidPm = MsgzLidMis(LidMis(A))
End Function

Private Function ImpSqyzLidPm(A As LidPm) As String()
PushIAy ImpSqyzLidPm, ImpSqyzFbAy(A.Fb)
PushIAy ImpSqyzLidPm, ImpSqyzFxAy(A.Fx)
End Function

Private Function ImpSqyzFbAy(A() As LidFb) As String()
Dim J%
For J = 0 To UB(A)
    PushI ImpSqyzFbAy, ImpSqyzFb(A(J))
Next
End Function

Private Function ImpSqyzFxAy(A() As LidFx) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy ImpSqyzFxAy, ImpSqyzFx(A(J))
Next
End Function

Private Function ImpSqyzFx(A As LidFx) As String()
Dim Fm$: Fm = ">" & A.T
Dim Into$: Into = "#I" & A.T
Dim Bexpr$: Bexpr = A.Bexpr
Dim Fny$(): Fny = FnyzLidFxcAy(A.Fxc)
Dim ExtNy$(): ExtNy = ExTnyzLidFxcAy(A.Fxc)
Dim O$()
PushI O, SqlSel_FF_ExtNy_Into_Fm(Fny, ExtNy, Into, Fm, Bexpr)
ImpSqyzFx = O
End Function

Private Function ExTnyzLidFxcAy(A() As LidFxc) As String()
Dim J%
For J = 0 To UB(A)
    PushI ExTnyzLidFxcAy, A(J).ExtNm
Next
End Function

Private Function FnyzLidFxcAy(A() As LidFxc) As String()
Dim J%
For J = 0 To UB(A)
    PushI FnyzLidFxcAy, A(J).ColNm
Next
End Function

Private Function ImpSqyzFb(A As LiFb) As String()
With A
Dim FF$(): FF = .Fset.Sy
Dim Fm$: Fm = ">" & .T
Dim Into$: Into = "#I" & .T
Dim Bexpr$: Bexpr = .Bexpr
End With
Dim O$()
PushI O, "Drop table [" & Into & "]"
PushI O, SqlSel_FF_Into_Fm(FF, Into, FF, Bexpr)
ImpSqyzFb = O
End Function

Private Sub Z_ImpSqy()
FmtSql = True
Brw ImpSqyzLidPm(RptLidPm)
End Sub

Private Sub Z_ErzLidPm()
Brw ErzLidPm(RptLidPm)
End Sub
Private Sub Z()

End Sub


