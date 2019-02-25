Attribute VB_Name = "MVb_Ay_Map_Transform"
Option Explicit

Function AyAddIxPfx(A, Optional BegFm&) As String()
Dim I, J&, N%
J = BegFm
N = Len(CStr(Sz(A)))
For Each I In Itr(A)
    PushI AyAddIxPfx, AlignR(J, N) & ": " & I
    J = J + 1
Next
End Function
Function AyIncEle1(A)
AyIncEle1 = AyIncEleN(A, 1)
End Function

Function AyIncEleN(A, N)
Dim O: O = A
Dim J&
For J = 0 To UB(O)
    O(J) = O(J) + N
Next
AyIncEleN = O
End Function

Function T1Ay(Ay) As String()
Dim O$(), L, J&, N&
N = Sz(Ay)
ReszAyN O, N
For Each L In Itr(Ay)
    AsgBrk L, " ", O(J)
    J = J + 1
Next
End Function

