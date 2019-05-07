Attribute VB_Name = "QVb_Sy_SyOp"
Option Explicit
Private Const CMod$ = "BSyOp."

Function RmvFstChrzSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI RmvFstChrzSy, RmvFstChr(CStr(I))
Next
End Function

Function RmvFstNonLetterzSy(Sy$()) As String() 'Gen:SyXXX
Dim I
For Each I In Itr(Sy)
    PushI RmvFstNonLetterzSy, RmvFstNonLetter(CStr(I))
Next
End Function
Function RmvLasChrzSy(Sy$()) As String()
'Gen:SyFor RmvLasChr
Dim I
For Each I In Itr(Sy)
    PushI RmvLasChrzSy, RmvLasChr(CStr(I))
Next
End Function

Function RmvPfxzSy(Sy$(), Pfx$) As String()
Dim I
For Each I In Itr(Sy)
    PushI RmvPfxzSy, RmvPfx(CStr(I), Pfx)
Next
End Function

Function SyeSngQRmk(Sy$()) As String()
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If Not IsSngQRmk(S) Then PushI SyeSngQRmk, S
Next
End Function

Function RmvSngQuotezSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI RmvSngQuotezSy, RmvSngQuote(CStr(I))
Next
End Function

Function RmvT1zSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI RmvT1zSy, RmvT1(CStr(I))
Next
End Function

Function RmvTTzSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI RmvTTzSy, RmvTT(CStr(I))
Next
End Function

Function RplSy(Sy$(), Fm$, By$, Optional Cnt& = 1) As String()
Dim I
For Each I In Itr(Sy)
    PushS RplSy, Replace(I, Fm, By, Count:=Cnt)
Next
End Function
Function Rmv2DashzSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI Rmv2DashzSy, Rmv2Dash(CStr(I))
Next
End Function

Function RplStarzSy(Sy$(), By$) As String()
Dim I
For Each I In Itr(Sy)
    PushI RplStarzSy, Replace(I, By, "*")
Next
End Function

Function RplT1zSy(Sy$(), NewT1$) As String()
RplT1zSy = AddPfxzSy(RmvT1zSy(Sy), NewT1 & " ")
End Function

Function AddIxPfx(Sy$(), Optional BegFm&) As String()
Dim I, J&, N%
J = BegFm
N = Len(CStr(Si(Sy)))
For Each I In Itr(Sy)
    PushI AddIxPfx, AlignR(CStr(J), N) & ": " & I
    J = J + 1
Next
End Function

Function T1Sy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI T1Sy, T1(CStr(I))
Next
End Function

Function T2Sy(Sy$()) As String()
Dim L
For Each L In Itr(Sy)
    PushI T2Sy, T2(CStr(L))
Next
End Function

Function T3Sy(Sy$()) As String()
Dim L
For Each L In Itr(Sy)
    PushI T3Sy, T3(CStr(L))
Next
End Function

