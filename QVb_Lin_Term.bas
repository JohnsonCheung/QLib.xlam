Attribute VB_Name = "QVb_Lin_Term"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Lin_Term_Asg."
Private Const Asm$ = "QVb"

Sub AsgN2tRst(Lin, OT1, OT2, ORst$)
AsgAp Syz2TRst(Lin), OT1, OT2, ORst
End Sub

Sub AsgN3tRst(Lin, OT1, OT2, OT3, ORst$)
AsgAp Syz3TRst(Lin), OT1, OT2, OT3, ORst
End Sub

Sub AsgN4t(Lin, O1$, O2$, O3$, O4$)
AsgAp Fst4Term(Lin), O1, O2, O3, O4
End Sub

Sub AsgN4tRst(Lin, O1$, O2$, O3$, O4$, ORst$)
AsgAp Syz4TRst(Lin), O1, O2, O3, O4, ORst
End Sub

Sub AsgTRst(Lin, OT1, ORst)
AsgAp SyzTRst(Lin), OT1, ORst
End Sub

Sub AsgN2t(Lin, O1, O2)
AsgAp Syz2TRst(Lin), O1, O2
End Sub

Sub AsgT1FldLikAy(OT1, OFldLikAy$(), Lin)
Dim Rst$
AsgTRst Lin, OT1, Rst
OFldLikAy = SyzSS(Rst)
End Sub


Function Fst2Term(Lin) As String()
Fst2Term = FstNTerm(Lin, 2)
End Function

Function Fst3Term(Lin) As String()
Fst3Term = FstNTerm(Lin, 3)
End Function
Function Fst4Term(Lin) As String()
Fst4Term = FstNTerm(Lin, 4)
End Function

Function FstNTerm(Lin, N%) As String()
Dim J%, L$
L = Lin
For J = 1 To N
    PushI FstNTerm, ShfT1(L)
Next
End Function


Function SyzTRst(Lin) As String()
SyzTRst = SyzNTermRst(Lin, 1)
End Function

Function Syz2TRst(Lin) As String()
Syz2TRst = SyzNTermRst(Lin, 2)
End Function

Function Syz3TRst(Lin) As String()
Syz3TRst = SyzNTermRst(Lin, 3)
End Function

Function Syz4TRst(Lin) As String()
Syz4TRst = SyzNTermRst(Lin, 4)
End Function

Function SyzNTermRst(Lin, N%) As String()
Dim L$, J%
L = Lin
For J = 1 To N
    PushI SyzNTermRst, ShfT1(L)
Next
PushI SyzNTermRst, L
End Function

Private Sub Z_SyzNTermRst()
Dim Lin
Lin = "  [ksldfj ]":  Ept = "ksldfj ": GoSub Tst
Lin = "  [ ksldfj ]": Ept = " ksldf ": GoSub Tst
Lin = "  [ksldfj]":  Ept = "ksldf": GoSub Tst
Exit Sub
Tst:
    Act = T1(Lin)
    C
    Return
End Sub

Private Sub ZZ()
End Sub
Function SrcT1AsetP() As Aset
Set SrcT1AsetP = T1Aset(SrczP(CPj))
End Function
Function T1Aset(Ly$()) As Aset
Dim O As New Aset, L
For Each L In Itr(Ly)
    O.PushItm T1(L)
Next
Set T1Aset = O
End Function
Function T1zS$(S)
T1zS = T1(S)
End Function

Function T1$(S)
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
Function T2zS$(S)
T2zS = T2(S)
End Function

Function T2$(S)
T2 = TermN(S, 2)
End Function

Function T3$(S)
T3 = TermN(S, 3)
End Function

Function TermN$(S, N%)
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

