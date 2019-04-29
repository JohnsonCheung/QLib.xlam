Attribute VB_Name = "BSyOp"
Option Explicit
Private Sub Y(S$, X$)
PushI XX, RplQ(S, X)
End Sub

Function SyRmvFstChr(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyRmvFstChr, RmvFstChr(CStr(I))
Next
End Function

Function SyRmvFstNonLetter(Sy$()) As String() 'Gen:SyXXX
Dim I
For Each I In Itr(Sy)
    PushI SyRmvFstNonLetter, RmvFstNonLetter(CStr(I))
Next
End Function
Function SyRmvLasChr(Sy$()) As String()
'Gen:SyFor RmvLasChr
Dim I
For Each I In Itr(Sy)
    PushI SyRmvLasChr, RmvLasChr(CStr(I))
Next
End Function

Function SyRmvPfx(Sy$(), Pfx$) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyRmvPfx, RmvPfx(CStr(I), Pfx)
Next
End Function

Function SyeSngQRmk(Sy$()) As String()
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If Not IsSngQRmk(S) Then PushI SyeSngQRmk, S
Next
End Function

Function SyRmvSngQuote(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyRmvSngQuote, RmvSngQuote(CStr(I))
Next
End Function

Function SyRmvT1(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyRmvT1, RmvT1(CStr(I))
Next
End Function

Function SyRmvTT(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyRmvTT, RmvTT(CStr(I))
Next
End Function

Function SyRpl(Sy$(), Fm$, By$, Optional Cnt& = 1) As String()
Dim I
For Each I In Itr(Sy)
    PushS SyRpl, Replace(I, Fm, By, Count:=Cnt)
Next
End Function
Function SyRmv2Dash(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyRmv2Dash, Rmv2Dash(CStr(I))
Next
End Function

Function SyRplStar(Sy$(), By$) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyRplStar, Replace(I, By, "*")
Next
End Function

Function SyRplT1(Sy$(), NewT1$) As String()
SyRplT1 = SyAddPfx(SyRmvT1(Sy), NewT1 & " ")
End Function

Function SyAddIxPfx(Sy$(), Optional BegFm&) As String()
Dim I, J&, N%
J = BegFm
N = Len(CStr(Si(Sy)))
For Each I In Itr(Sy)
    PushI SyAddIxPfx, AlignR(J, N) & ": " & I
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

