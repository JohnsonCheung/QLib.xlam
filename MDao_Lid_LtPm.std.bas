Attribute VB_Name = "MDao_Lid_LtPm"
Option Explicit
Function LtPmzLid(A As LidPm) As LtPm()
Dim O() As LtPm, D As Dictionary
Set D = FilNmToFfnDicvLid(A.Fil)
PushObjAy O, LtPmAyFb(A.Fb, D)
PushObjAy O, LtPmAyFx(A.Fx, D)
LtPmzLid = O
End Function

Private Function LtPmAyFx(A() As LidFx, FfnDic As Dictionary) As LtPm()
Dim J%, Fx$, M As LtPm
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        Fx = FfnDic(.Fxn)
        PushObj LtPmAyFx, M.Init(">" & .T, .Wsn & "$", CnStrzFxDAO(Fx))
    End With
Next
End Function
Private Function LtPmAyFb(A() As LidFb, FfnDic As Dictionary) As LtPm()
Dim J%, Fb$, M As LtPm
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        Fb = FfnDic(.Fbn)
        PushObj LtPmAyFb, M.Init(">" & .T, .T, CnStrzFxAdo(Fb))
    End With
Next
End Function

Private Function FilNmToFfnDicvLid(A() As LidFil) As Dictionary
Dim J%
Set FilNmToFfnDicvLid = New Dictionary
For J = 0 To UB(A)
    With A(J)
        FilNmToFfnDicvLid.Add .FilNm, .Ffn
    End With
Next
End Function
