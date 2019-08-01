Attribute VB_Name = "QVb_Ay_Ayab"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_AB."
Private Const Asm$ = "QVb"

Function LyzAyab(AyA, AyB, Optional Sep$) As String()
ThwIf_DifSi AyA, AyB, CSub
Dim A, J&: For Each A In Itr(AyA)
    PushI LyzAyab, A & Sep & AyB(J)
    J = J + 1
Next
End Function

Function LyzAyabSpc(AyA, AyB) As String()
LyzAyabSpc = LyzAyab(AyA, AyB, " ")
End Function

Function FmtAyab(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As String()
FmtAyab = FmtS12s(S12szAyab(A, B), N1, N2)
End Function

Function LyzAyabNEmpB(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
LyzAyabNEmpB = O
End Function

Sub AsgAyaReSzMax(A, B, OA, OB)
OA = A
OB = B
ResiMax OA, OB
End Sub

Sub ThwImpossible(Fun$)
Thw Fun, "Impossible to reach here"
End Sub
