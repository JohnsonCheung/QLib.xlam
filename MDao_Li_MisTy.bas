Attribute VB_Name = "MDao_Li_MisTy"
Option Explicit

Function MisTyFxAy(A() As LiFx, B() As LiActFx) As LiMisTy()
Dim J%
For J = 0 To UB(A)
    With A(J)
        PushNonNothing MisTyFxAy, MisTyFxOpt(A(J), B)
    End With
Next
End Function

Private Function MisTyFxOpt(A As LiFx, B() As LiActFx) As LiMisTy
Dim Act As LiActFx:    Set Act = LiActFxOpt(A.Fxn, A.Wsn, B): If IsNothing(Act) Then Exit Function
Dim TycAy() As LiMisTyc: TycAy = MisColAy(A.FxcAy, Act.ShtTyDic): If Si(TycAy) = 0 Then Exit Function
Set MisTyFxOpt = New LiMisTy
MisTyFxOpt.Init Act.Fx, A.Fxn, A.Wsn, TycAy
End Function

Private Function MisColAy(Ept() As LiFxc, ActShtTyDic As Dictionary) As LiMisTyc()
Dim J%
For J = 0 To UB(Ept)
    PushNonNothing MisColAy, MisColOpt(Ept(J), ActShtTyDic)
Next
End Function

Private Function MisColOpt(Ept As LiFxc, ActShtTyDic As Dictionary) As LiMisTyc
If Not ActShtTyDic.Exists(Ept.ExtNm) Then Exit Function
Dim ActShtTy$: ActShtTy = ActShtTyDic(Ept.ExtNm)
If Not IsMisTy(Ept.ShtTyLis, ActShtTy) Then Exit Function
Set MisColOpt = New LiMisTyc
MisColOpt.Init Ept.ExtNm, ActShtTy, Ept.ShtTyLis
End Function

Private Function IsMisTy(EptShtTyLis$, ActShtTy$) As Boolean
IsMisTy = Not HasEle(CmlAy(EptShtTyLis), ActShtTy)
End Function

Private Function LiActFxOpt(Fxn$, Wsn$, A() As LiActFx) As LiActFx
Dim J%
For J = 0 To UB(A)
    With A(J)
        If Fxn = .Fxn Then
            If Wsn = .Wsn Then
                Set LiActFxOpt = A(J)
                Exit Function
            End If
        End If
    End With
Next
End Function
