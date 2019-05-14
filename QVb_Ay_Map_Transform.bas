Attribute VB_Name = "QVb_Ay_Map_Transform"
Option Explicit
Private Const CMod$ = "MVb_Ay_Map_Transform."
Private Const Asm$ = "QVb"
Public Const DoczTLin = "Is a line with Terms separated by spc."
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

Function TermAsetzTLiny(TLiny$()) As Aset
Dim I, O$(), TLin
For Each I In Itr(TLiny)
    TLin = I
    PushIAy O, TermAy(TLin)
Next
Set TermAsetzTLiny = AsetzAy(O)
End Function

