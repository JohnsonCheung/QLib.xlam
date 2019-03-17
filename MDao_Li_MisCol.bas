Attribute VB_Name = "MDao_Li_MisCol"
Option Explicit
Function MisCol(A As LiPm, B As LiAct) As LiMisCol()
PushObjzAy MisCol, MisColFx(A.Fx, B.Fx)
PushObjzAy MisCol, MisColFb(A.Fb, B.Fb)
End Function

Private Function MisColFx(A() As LiFx, B() As LiActFx) As LiMisCol()
Dim J%, ActFset As Aset, Fx$
For J = 0 To UB(A)
    AsgActFsetFx ActFset, Fx, A(J).T, B
    If Not IsNothing(ActFset) Then
        PushObj MisColFx, MisColFxOpt(A(J), Fx, ActFset)
    End If
Next
End Function

Private Function MisColFb(A() As LiFb, B() As LiActFb) As LiMisCol()
Dim J%, ActFset As Aset, Fb$
For J = 0 To UB(A)
    AsgActFsetFb ActFset, Fb, A(J).T, B
    If Not IsNothing(ActFset) Then
        PushObj MisColFb, MisColFbOpt(A(J), Fb, ActFset)
    End If
Next
End Function

Private Function MisColFxOpt(A As LiFx, Fx$, ActFset As Aset) As LiMisCol
Dim EptFset As New Aset: Set EptFset = A.EptFset
If EptFset.Minus(ActFset).Cnt = 0 Then Exit Function
Set MisColFxOpt = New LiMisCol
MisColFxOpt.Init Fx, A.T, EptFset, ActFset, A.Wsn
End Function

Private Function MisColFbOpt(A As LiFb, Fb$, ActFset As Aset) As LiMisCol
Dim EptFset As New Aset: Set EptFset = A.Fset
If EptFset.Minus(ActFset).Cnt = 0 Then Exit Function
Set MisColFbOpt = New LiMisCol
MisColFbOpt.Init Fb, A.T, EptFset, ActFset
End Function


Private Sub AsgActFsetFb(OActFsetOpt As Aset, OFb$, T$, B() As LiActFb)
Set OActFsetOpt = Nothing
Dim J%
For J = 0 To UB(B)
    With B(J)
    If .T = T Then
        Set OActFsetOpt = .Fset
        OFb = .Fb
        Exit Sub
    End If
    End With
Next
End Sub

Private Sub AsgActFsetFx(OActFsetOpt As Aset, OFx$, Wsn$, B() As LiActFx)
Set OActFsetOpt = Nothing
Dim J%
For J = 0 To UB(B)
    With B(J)
    If .Wsn = Wsn Then
        Set OActFsetOpt = .ShtTyDic
        OFx = .Fx
        Exit Sub
    End If
    End With
Next
End Sub


