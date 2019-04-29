Attribute VB_Name = "MVb_Ay_Map_Transform"
Option Explicit
Public Const DocOfTLin$ = "Is a line with Terms separated by spc."
Function AyIncEle1(Ay)
AyIncEle1 = AyIncEleN(Ay, 1)
End Function

Function AyIncEleN(Ay, N)
Dim O: O = Ay
Dim J&
For J = 0 To UB(O)
    O(J) = O(J) + N
Next
AyIncEleN = O
End Function

Function TermAsetzTLinAy(TLinAy$()) As Aset
Dim I, O$(), TLin$
For Each I In Itr(TLinAy)
    TLin = I
    PushIAy O, TermAy(TLin)
Next
Set TermAsetzTLinAy = AsetzAy(O)
End Function

