Attribute VB_Name = "MVb_Lin_Term_TermN"
Option Explicit
Function T1zLin$(A)
T1zLin = T1(A)
End Function
Function T1$(Lin)
T1 = TermN(Lin, 1)
End Function
Function T2zLin$(A)
T2zLin = TermN(A, 2)
End Function

Function T2$(A)
T2 = TermN(A, 2)
End Function

Function T3$(A)
T3 = TermN(A, 3)
End Function

Function TermN$(Lin, N%)
Dim L$, J%
L = LTrim(Lin)
For J = 1 To N - 1
    L = RmvT1(L)
Next
TermN = TakT1(L)
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
