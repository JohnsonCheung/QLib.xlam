Attribute VB_Name = "MDao_Lid_LtPm"
Option Explicit
Function LtPmzLid(A As LidPm) As LtPm()
Dim O() As LtPm, D As Dictionary, P$
Set D = FilNmToFfnDiczLidPm(A.Fil)
P = WPth(A.Apn)
PushObjAy O, LtPmAyFb(A, D, P)
PushObjAy O, LtPmAyFx(A, D, P)
LtPmzLid = O
End Function

Private Function LtPmAyFx(Pm As LidPm, FfnDic As Dictionary, WPth$) As LtPm()
Dim J%, Fx$, M As LtPm, A() As LidFx
A = Pm.Fx
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        Fx = WPth & Fn(FfnDic(.Fxn))
        PushObj LtPmAyFx, M.Init(">" & .T, .Wsn & "$", CnStrzFxDAO(Fx))
    End With
Next
End Function
Private Function LtPmAyFb(Pm As LidPm, FfnDic As Dictionary, WPth$) As LtPm()
Dim J%, Fb$, M As LtPm, A() As LidFb
A = Pm.Fb
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        Fb = WPth & Fn(FfnDic(.Fbn))
        PushObj LtPmAyFb, M.Init(">" & .T, .T, CnStrzFxAdo(Fb))
    End With
Next
End Function

Private Function FilNmToFfnDiczLidPm(A() As LidFil) As Dictionary
Dim J%
Set FilNmToFfnDiczLidPm = New Dictionary
For J = 0 To UB(A)
    With A(J)
        FilNmToFfnDiczLidPm.Add .FilNm, .Ffn
    End With
Next
End Function
