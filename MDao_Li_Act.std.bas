Attribute VB_Name = "MDao_Li_Act"
Option Explicit

Function LiAct(A As LiPm) As LiAct
Set LiAct = New LiAct
Dim D As Dictionary: Set D = A.ExistFilNmToFfnDic
LiAct.Init LiActFxAy(A.Fx, D), LiActFbAy(A.Fb, D)
End Function

Private Function LiActFxOpt(A As LiFx, Fx$) As LiActFx
If Not HasFxw(Fx, A.Wsn) Then Exit Function
Set LiActFxOpt = New LiActFx
With A
LiActFxOpt.Init Fx, .Fxn, .Wsn, ShtTyDic(Fx, .Wsn)
End With
End Function
Private Function LiActFxAy(A() As LiFx, ExistFilNmToFfnDic As Dictionary) As LiActFx()
Dim J%, Fx$
For J = 0 To UB(A)
    If ExistFilNmToFfnDic.Exists(A(J).Fxn) Then
        Fx = ExistFilNmToFfnDic(A(J).Fxn)
        PushNonNothing LiActFxAy, LiActFxOpt(A(J), Fx)
    End If
Next
End Function

Private Function LiActFbAy(A() As LiFb, ExistFilNmToFfnDic As Dictionary) As LiActFb()
Dim J%, Fb$
For J = 0 To UB(A)
    If ExistFilNmToFfnDic.Exists(A(J).Fbn) Then
        Fb = ExistFilNmToFfnDic(A(J).Fbn)
        PushNonNothing LiActFbAy, LiActFbOpt(A(J), Fb)
    End If
Next
End Function

Private Function LiActFbOpt(A As LiFb, Fb$) As LiActFb
If Not HasFbt(Fb, A.T) Then Exit Function
Set LiActFbOpt = New LiActFb
With A
LiActFbOpt.Init Fb, .Fbn, .T, AsetzAy(FnyzFbt(Fb, .T))
End With
End Function

