Attribute VB_Name = "QVb_S1S2"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_S1S2."
Private Const Asm$ = "QVb"
Type S1S2: S1 As String: S2 As String: End Type
Type S1S2s: N As Long: Ay() As S1S2: End Type
Type S3: A As String: B As String: C As String: End Type
Function SwapS1S2s(A As S1S2s) As S1S2s
Dim J&, Ay() As S1S2: Ay = A.Ay
Dim O As S1S2s: O = A
For J = 1 To A.N - 1
    O.Ay(J) = SwapS1S2(Ay(J))
Next
SwapS1S2s = O
End Function
Function SwapS1S2(A As S1S2) As S1S2
With SwapS1S2
    .S1 = A.S2
    .S2 = A.S1
End With
End Function
Sub PushS1S2(O As S1S2s, M As S1S2)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function AddS1S2(A As S1S2, B As S1S2) As S1S2s
Dim O As S1S2s
O = S1S2szU(1)
O.Ay(0) = A
O.Ay(1) = B
AddS1S2 = O
End Function
Private Function Y_S1S2s() As S1S2s
Dim O As S1S2s
PushS1S2 O, S1S2("sldjflsdkjf", "lksdjf")
PushS1S2 O, S1S2("sldjflsdkjf", "lksdjf")
PushS1S2 O, S1S2("sldjf", "lksdjf")
PushS1S2 O, S1S2("sldjdkjf", "lksdjf")
Y_S1S2s = O
End Function
Function S1S2szU(U&) As S1S2s
S1S2szU.N = U + 1
ReDim S1S2szU.Ay(U)
End Function
Function S1S2szAyab(A, B, Optional NoTrim As Boolean) As S1S2s
ThwIf_DifSi A, B, CSub
Dim U&, O As S1S2s
U = UB(A)
O = S1S2szU(U)
Dim J&
For J = 0 To U
    O.Ay(J) = S1S2(A(J), B(J), NoTrim)
Next
S1S2szAyab = O
End Function

Function S1S2(Optional S1, Optional S2, Optional NoTrim As Boolean) As S1S2
If NoTrim Then
    S1S2.S1 = S1
    S1S2.S2 = S2
Else
    S1S2.S1 = Trim(S1)
    S1S2.S2 = Trim(S2)
End If
End Function

Sub AsgS1S2(A As S1S2, O1, O2)
O1 = A.S1
O2 = A.S2
End Sub

Function LinzS1S2$(A As S1S2, Optional Sep$ = " ", Optional W%)
LinzS1S2 = AlignL(A.S1, W) & Sep & A.S2
End Function
Function W1zS1S2s%(A As S1S2s)
Dim O%, J&
For J = 0 To A.N - 1
    O = Max(O, Len(A.Ay(J).S1))
Next
End Function
Function W2zS1S2s%(A As S1S2s)
Dim O%, J&
For J = 0 To A.N - 1
    O = Max(O, Len(A.Ay(J).S2))
Next
End Function
Function LyzS1S2s(A As S1S2s, Optional Sep$ = "") As String()
Dim O$(), J&, W%, Ay() As S1S2
Ay = A.Ay
W = W1zS1S2s(A)
For J = 0 To A.N - 1
   PushI LyzS1S2s, LinzS1S2(Ay(J), Sep, W)
Next
End Function

Sub BrwS1S2s(A As S1S2s)
BrwAy FmtS1S2s(A)
End Sub

Function S1S2szDic(A As Dictionary) As S1S2s
Dim K
For Each K In A.Keys
    PushS1S2 S1S2szDic, S1S2(K, LineszV(A(K)))
Next
End Function

Function DiczS1S2s(A As S1S2s, Optional Sep$ = " ") As Dictionary
Dim J&, O As New Dictionary
For J = 0 To A.N - 1
    With A.Ay(J)
        If O.Exists(.S1) Then
            O(.S1) = O(.S1) & " " & O(.S2)
        Else
            O.Add .S1, .S2
        End If
    End With
Next
Set DiczS1S2s = O
End Function

Function Sy1zS1S2s(A As S1S2s) As String()
Dim J&
For J = 0 To A.N - 1
   PushI Sy1zS1S2s, A.Ay(J).S1
Next
End Function

Function Sy2zS1S2s(A As S1S2s) As String()
Dim O$(), J&
For J = 0 To A.N - 1
   Push Sy2zS1S2s, A.Ay(J).S2
Next
Sy2zS1S2s = O
End Function

Function SqzS1S2s(A As S1S2s, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Variant()
If A.N = 0 Then Exit Function
Dim O(), I, R&, J&
ReDim O(1 To A.N + 1, 1 To 2)
R = 2
O(1, 1) = Nm1
O(1, 2) = Nm2
For J = 0 To A.N - 1
    With A.Ay(J)
        O(R, 1) = .S1
        O(R, 2) = .S2
        R = R + 1
    End With
Next
SqzS1S2s = O
End Function
Function S1S2szColonVbl(ColonVbl$) As S1S2s
Dim I
For Each I In SplitVBar(ColonVbl)
    PushS1S2 S1S2szColonVbl, BrkBoth(I, ":")
Next
End Function

Function S1S2szSySep(Sy$(), Sep$, Optional NoTrim As Boolean) As S1S2s
Dim O As S1S2s, J%
Dim U&: U = UB(Sy)
O = S1S2szU(U)
For J = 0 To U
    O.Ay(J) = Brk1(Sy(J), Sep, NoTrim)
Next
S1S2szSySep = O
End Function
Private Sub Z_S1S2szDic()
Dim A As New Dictionary
A.Add "A", "BB"
A.Add "B", "CCC"
Dim Act As S1S2s
Act = S1S2szDic(A)
Stop
End Sub


Private Sub ZZ()
Z_S1S2szDic
MVb__S1S2:
End Sub
