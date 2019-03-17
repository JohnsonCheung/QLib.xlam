Attribute VB_Name = "MVb_Ay_Map_Transform"
Option Explicit

Function AyAddIxPfx(A, Optional BegFm&) As String()
Dim I, J&, N%
J = BegFm
N = Len(CStr(Si(A)))
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
Dim L
For Each L In Itr(Ay)
    PushI T1Ay, T1(L)
Next
End Function

Function T2Ay(Ay) As String()
Dim L
For Each L In Itr(Ay)
    PushI T2Ay, T2(L)
Next
End Function


Function TermAsetzTLinAy(TLinAy$()) As Aset
Dim I, O$()
For Each I In Itr(TLinAy)
    PushIAy O, SySsl(I)
Next
Set TermAsetzTLinAy = AsetzAy(O)
End Function

