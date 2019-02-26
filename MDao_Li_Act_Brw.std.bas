Attribute VB_Name = "MDao_Li_Act_Brw"
Option Explicit

Sub BrwLiAct(A As LiAct)
BrwDrs LiActDrs(A)
End Sub

Function LiActDrs(A As LiAct) As DRs
Set LiActDrs = DRs("FilNm Wsn Fld ShtTy T FF Ffn", LiActDry(A))
End Function

Private Function LiActDry(A As LiAct) As Variant()
PushIAy LiActDry, LiActDryFxAy(A.Fx)
PushIAy LiActDry, LiActDryFbAy(A.Fb)
End Function
Private Function LiActDryFxAy(A() As LiActFx) As Variant()
Dim J%
For J = 0 To UB(A)
    PushIAy LiActDryFxAy, LiActDryFx(A(J))
Next
End Function
Private Function LiActDryFbAy(A() As LiActFb) As Variant()
Dim J%
For J = 0 To UB(A)
    PushI LiActDryFbAy, LiActDrFb(A(J))
Next
End Function

Private Function LiActDrFb(A As LiActFb) As Variant()
With A
LiActDrFb = Array(.Fbn, Empty, Empty, Empty, .T, .Fset.TermLin, .Fb)
End With
End Function
Private Function LiActDryFx(A As LiActFx) As Variant()
With A
    Dim Fld, D As Dictionary
    Set D = A.ShtTyDic
    For Each Fld In D.Keys
        PushI LiActDryFx, Array(.Fxn, .Wsn, Fld, D(Fld), Empty, Empty, .Fx)
    Next
End With
End Function

Private Sub Z_LiAct()
BrwDrs LiActDrs(LiAct(ShpCstLiPm))
End Sub

