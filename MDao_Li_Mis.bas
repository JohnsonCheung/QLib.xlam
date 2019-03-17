Attribute VB_Name = "MDao_Li_Mis"
Option Explicit
Function LiMis(A As LiPm, B As LiAct) As LiMis
Set LiMis = New LiMis
Dim A1() As LiMisTbl: A1 = MisTbl(A, B)
Dim A2() As LiMisCol: A2 = MisCol(A, B)
Dim A3() As LiMisTy:  A3 = MisTyFxAy(A.Fx, B.Fx)
LiMis.Init A.MisFfn, A1, A2, A3
End Function
Private Function MisTbl(A As LiPm, B As LiAct) As LiMisTbl()
PushObjzAy MisTbl, MisTblFb(A.Fb, B.Fb)
PushObjzAy MisTbl, MisTblFx(A.Fx, B.Fx)
End Function
Private Function MisTblFx(A() As LiFx, B() As LiActFx) As LiMisTbl()
Dim J%
For J = 0 To UB(A)
    PushNonNothing MisTblFx, MisTblFxOpt(A(J), B)
Next
End Function
Private Function MisTblFb(A() As LiFb, B() As LiActFb) As LiMisTbl()
Dim J%
For J = 0 To UB(A)
    PushNonNothing MisTblFb, MisTblFbOpt(A(J), B)
Next
End Function
Private Function MisTblFbOpt(A As LiFb, B() As LiActFb) As LiMisTbl
Dim Fb$: Fb = A.ExistFb(B): If Fb <> "" Then Exit Function
Set MisTblFbOpt = New LiMisTbl
With A
MisTblFbOpt.Init Fb, .Fbn, .T
End With
End Function

Private Function MisTblFxOpt(A As LiFx, B() As LiActFx) As LiMisTbl
Dim Fx$: Fx = A.ExistFx(B): If Fx <> "" Then Exit Function
Set MisTblFxOpt = New LiMisTbl
With A
MisTblFxOpt.Init Fx, .Fxn, .T, .Wsn
End With
End Function
