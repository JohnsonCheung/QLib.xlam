Attribute VB_Name = "QVb_Ay_AB"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_AB."
Private Const Asm$ = "QVb"

Function JnAyab(A, B, Optional Sep$) As String()
ThwIf_DifSi A, B, CSub
Dim J&
For J = 0 To UB(A)
    PushI JnAyab, A(J) & Sep & B(J)
Next
End Function

Function JnAyabSpc(A, B) As String()
JnAyabSpc = JnAyab(A, B, " ")
End Function

Function FmtAyab(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As String()
FmtAyab = FmtS1S2s(S1S2szAyab(A, B), N1, N2)
End Function

Function LyzAyabJnSepForEmpB(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
LyzAyabJnSepForEmpB = O
End Function

Sub AsgAyaReSzMax(A, B, OA, OB)
OA = A
OB = B
ResiMax OA, OB
End Sub
Sub ThwImpossible(Fun$)
Thw Fun, "Impossible to reach here"
End Sub
Sub ThwIf_AyabNE(A, B, Optional N1$ = "Exp", Optional N2$ = "Act")
ThwIf_Er ChkEqAy(A, B, N1, N2), CSub
End Sub
