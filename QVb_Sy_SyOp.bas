Attribute VB_Name = "QVb_Sy_SyOp"
Option Explicit
Private Const CMod$ = "BAyOp."

Function RmvFstChrzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI RmvFstChrzAy, RmvFstChr(CStr(I))
Next
End Function

Function RmvFstNonLetterzAy(Ay) As String() 'Gen:AyXXX
Dim I
For Each I In Itr(Ay)
    PushI RmvFstNonLetterzAy, RmvFstNonLetter(CStr(I))
Next
End Function
Function RmvLasChrzAy(Ay) As String()
'Gen:AyFor RmvLasChr
Dim I
For Each I In Itr(Ay)
    PushI RmvLasChrzAy, RmvLasChr(CStr(I))
Next
End Function

Function RmvPfxzAy(Ay, Pfx$) As String()
Dim I
For Each I In Itr(Ay)
    PushI RmvPfxzAy, RmvPfx(CStr(I), Pfx)
Next
End Function

Function AyeSngQRmk(Ay) As String()
Dim I, S$
For Each I In Itr(Ay)
    S = I
    If Not IsSngQRmk(S) Then PushI AyeSngQRmk, S
Next
End Function

Function RmvSngQuotezAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI RmvSngQuotezAy, RmvSngQuote(CStr(I))
Next
End Function

Function RmvT1zAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI RmvT1zAy, RmvT1(CStr(I))
Next
End Function

Function RmvTTzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI RmvTTzAy, RmvTT(CStr(I))
Next
End Function

Function RplAy(Ay, Fm$, By$, Optional Cnt& = 1) As String()
Dim I
For Each I In Itr(Ay)
    PushI RplAy, Replace(I, Fm, By, Count:=Cnt)
Next
End Function
Function Rmv2DashzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI Rmv2DashzAy, Rmv2Dash(CStr(I))
Next
End Function

Function RplStarzAy(Ay, By) As String()
Dim I
For Each I In Itr(Ay)
    PushI RplStarzAy, Replace(I, By, "*")
Next
End Function

Function RplT1zAy(Ay, NewT1) As String()
RplT1zAy = AddPfxzAy(RmvT1zAy(Ay), NewT1 & " ")
End Function

Function AddIxPfx(Ay, Optional BegFm&) As String()
Dim I, J&, N%
J = BegFm
N = Len(CStr(Si(Ay)))
For Each I In Itr(Ay)
    PushI AddIxPfx, AlignR(CStr(J), N) & ": " & I
    J = J + 1
Next
End Function

Function T1Ay(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI T1Ay, T1(CStr(I))
Next
End Function

Function T2Ay(Ay) As String()
Dim L
For Each L In Itr(Ay)
    PushI T2Ay, T2(CStr(L))
Next
End Function

Function T3Ay(Ay) As String()
Dim L
For Each L In Itr(Ay)
    PushI T3Ay, T3(CStr(L))
Next
End Function

