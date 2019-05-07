Attribute VB_Name = "QVb_Lin_Term_TermN"
Option Explicit
Private Const CMod$ = "MVb_Lin_Term_TermN."
Private Const Asm$ = "QVb"
Function T1zS$(S$)
T1zS = T1(S)
End Function

Function T1$(S$)
Dim O$: O = LTrim(S)
If FstChr(O) = "[" Then
    Dim P%
    P = InStr(S, "]")
    If P = 0 Then
        Thw CSub, "S has fstchr [, but no ]", "S", S
    End If
    T1 = Mid(S, 2, P - 2)
    Exit Function
End If
T1 = BefOrAll(S, " ")
End Function
Function T2zS$(S$)
T2zS = T2(S)
End Function

Function T2$(S$)
T2 = TermN(S, 2)
End Function

Function T3$(S$)
T3 = TermN(S, 3)
End Function

Function TermN$(S$, N%)
Dim L$, J%
L = LTrim(S)
For J = 1 To N - 1
    L = RmvT1(L)
Next
TermN = T1(L)
End Function

Private Sub Z_TermN()
Dim N%, A$
N = 1: A = "a b c": Ept = "a": GoSub Tst
N = 2: A = "a b c": Ept = "b": GoSub Tst
N = 3: A = "a b c": Ept = "c": GoSub Tst
Exit Sub
Tst:

    Act = TermN(A, N)
    C
    Return
End Sub


Private Sub Z()
Z_TermN
MVb_Lin_Term_TermN:
End Sub
