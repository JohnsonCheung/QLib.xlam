Attribute VB_Name = "MVb_Ay_AB"
Option Explicit
Function JnAyab(A, B, Optional Sep$) As String()
Dim AA, BB: AA = A: BB = B
ReszAyabMax AA, BB
Dim J&, U&
U = UB(AA)
If U = -1 Then Exit Function
ReDim O$(U)
For J = 0 To U
    O(J) = A(J) & Sep & B(J)
Next
JnAyab = O
End Function

Function JnAyabSpc(A, B) As String()
JnAyabSpc = JnAyab(A, B, " ")
End Function

Function DicAyab(A, B) As Dictionary
ThwDifSz A, B, CSub
Dim N1&, N2&
N1 = Si(A)
N2 = Si(B)
If N1 <> N2 Then Stop
Set DicAyab = New Dictionary
Dim J&, X
For Each X In Itr(A)
    DicAyab.Add X, B(J)
    J = J + 1
Next
End Function

Function FmtAyab(A, B) As String()
FmtAyab = FmtS1S2Ay(S1S2AyAyab(A, B))
End Function

Function LyAyabJnsepForNonEmpB(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
LyAyabJnsepForNonEmpB = O
End Function

Sub AsgAyaReSzMax(A, B, OA, OB)
OA = A
OB = B
ReszAyabMax OA, OB
End Sub
Sub ReszAyabMax(OA, OB)
Dim U1&, U2&
U1 = UB(OA)
U2 = UB(OB)
Select Case True
Case U1 > U2: ReDim Preserve OB(U1)
Case U1 < U2: ReDim Preserve OA(U2)
End Select
End Sub

Sub ThwAyabNE(A, B, Optional N1$ = "Exp", Optional N2$ = "Act")
'ChkAss ChkEqAy(A, B, N1, N2)
End Sub
