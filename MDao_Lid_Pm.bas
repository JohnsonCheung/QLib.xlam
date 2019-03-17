Attribute VB_Name = "MDao_Lid_Pm"
Option Explicit
Private A$()

Function LidPm(Src$()) As LidPm
If Si(Src) = 0 Then Thw CSub, "No lines in Src"
If T1(Src(0)) <> "LidPm" Then Thw CSub, "First line must be LidPm", "Src", Src
A = Src
Set LidPm = New LidPm
LidPm.Init Apn, Fil, Fx, Fb
End Function

Private Property Get Apn$()
Apn = RmvT1(FstEleT1(A, "Apn"))
End Property
Private Function LyT1(T1$) As String()
LyT1 = AywRmvT1(A, T1)
End Function
Private Property Get Fx() As LidFx()
Dim Ay$(): Ay = LyT1("Ws")
Dim WsLin
For Each WsLin In Itr(Ay)
    PushObj Fx, Fxi(WsLin)
Next
End Property

Private Function Fxi(WsLin) As LidFx
Dim T$, Fxn$, Wsn$, Bexpr$
Asg3TRst WsLin, T, Fxn, Wsn, Bexpr
Set Fxi = New LidFx
Fxi.Init Fxn, Wsn, T, FxcAy(T), Bexpr
End Function
Private Function FxcAy(T$) As LidFxc()
Dim L
For Each L In Itr(LyzWsCol(T))
    PushObj FxcAy, Fxc(L)
Next
End Function

Private Function Fxc(WsColLin) As LidFxc
Dim ColNm$, ShtTyLis$, ExtNm$
Asg2TRst WsColLin, ColNm, ShtTyLis, ExtNm
Set Fxc = New LidFxc
Fxc.Init ColNm, ShtTyLis, ExtNm
End Function

Private Function LyzWsCol(TblNm$)
LyzWsCol = AywRmvTT(A, "WsCol", TblNm)
End Function

Private Function FfnDic() As Dictionary
Set FfnDic = Dic(AywRmvT1(A, "Fil"))
End Function

Private Property Get Fb() As LidFb()
Dim TblLin, D As Dictionary
Set D = FfnDic
For Each TblLin In ItrzAywRmvT1(A, "Tbl")
    PushObj Fb, Fbi(TblLin, D)
Next
End Property

Private Function Fbi(TblLin, FfnDic As Dictionary) As LidFb
Dim T$, Fbn$, FF$, Bexpr$, Fb$
Asg3TRst TblLin, T, Fbn, FF, Bexpr
Fb = FfnDic(Fbn)
Set Fbi = New LidFb
Fbi.Init Fbn, T, AsetzFF(FF), Bexpr, Fb
End Function

Private Property Get Fil() As LidFil()
Dim L
For Each L In ItrzAywRmvT1(A, "Fil")
    PushObj Fil, Fili(L)
Next
End Property

Private Function Fili(L) As LidFil
Dim T$, Rst$
AsgTRst L, T, Rst
Set Fili = New LidFil
Fili.Init T, Rst
End Function

