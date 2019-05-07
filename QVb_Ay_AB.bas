Attribute VB_Name = "QVb_Ay_AB"
Option Explicit
Private Const CMod$ = "MVb_Ay_AB."
Private Const Asm$ = "QVb"

Function JnAyab(A, B, Optional Sep$) As String()
Dim AA, BB: AA = A: BB = B
Dim J&, U&
U = UB(ResiMax(AA, BB))
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

Function AddDic(A As Dictionary, B As Dictionary) As Dictionary
Set AddDic = New Dictionary
PushDic AddDic, A
PushDic AddDic, B
End Function

Sub PushDic(O As Dictionary, A As Dictionary)
Dim K
For Each K In A.Keys
    If O.Exists(A) Then Thw CSub, "O already has K.  Cannot push Dic-A to Dic-O", "K Dic-O Dic-A", K, O, A
    O.Add K, A(K)
Next
End Sub

Function DiczAyab(A, B) As Dictionary
ThwIfDifSi A, B, CSub
Dim N1&, N2&
N1 = Si(A)
N2 = Si(B)
If N1 <> N2 Then Stop
Set DiczAyab = New Dictionary
Dim J&, X
For Each X In Itr(A)
    DiczAyab.Add X, B(J)
    J = J + 1
Next
End Function

Function FmtAyab(A, B) As String()
FmtAyab = FmtS1S2s(S1S2szAyab(A, B))
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
Sub ThwIfAyabNE(A, B, Optional N1$ = "Exp", Optional N2$ = "Act")
ThwIfEr ChkEqAy(A, B, N1, N2), CSub
End Sub
