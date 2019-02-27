Attribute VB_Name = "MVb_S1S2"
Option Explicit

Private Property Get ZZS1S2Ay1() As S1S2()
Dim O() As S1S2
PushObj O, S1S2("sldjflsdkjf", "lksdjf")
PushObj O, S1S2("sldjflsdkjf", "lksdjf")
PushObj O, S1S2("sldjf", "lksdjf")
PushObj O, S1S2("sldjdkjf", "lksdjf")
ZZS1S2Ay1 = O
End Property

Function S1S2AyAyab(A, B, Optional NoTrim As Boolean) As S1S2()
ThwDifSz A, B, CSub
Dim U&, O() As S1S2
U = UB(A)
ReszAyU O, U
Dim J&
For J = 0 To UB(A)
    Set O(J) = S1S2(A(J), B(J), NoTrim)
Next
S1S2AyAyab = O
End Function

Function CvS1S2(A) As S1S2
Set CvS1S2 = A
End Function
Function S1S2Ay(ParamArray S1S2Ap()) As S1S2()
Dim Av(): Av = S1S2Ap
Dim I
For Each I In Av
    PushObj S1S2Ay, I
Next
End Function
Function S1S2(S1, S2, Optional NoTrim As Boolean) As S1S2
Set S1S2 = New S1S2
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

Function S1S2Clone(A As S1S2) As S1S2
Set S1S2Clone = S1S2(A.S1, A.S2)
End Function

Function S1S2Lin$(A As S1S2, Optional Sep$ = " ", Optional W1%)
S1S2Lin = AlignL(A.S1, W1) & Sep & A.S2
End Function

Function S1S2AyAddAsLy(A() As S1S2, Optional Sep$ = "") As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S1 & Sep & A(J).S2
Next
S1S2AyAddAsLy = O
End Function

Sub BrwS1S2Ay(A() As S1S2)
BrwAy FmtS1S2Ay(A)
End Sub

Function S1S2AyzDic(A As Dictionary) As S1S2()
Dim K
For Each K In A.Keys
    PushObj S1S2AyzDic, S1S2(K, LineszVal(A(K)))
Next
End Function

Function DiczS1S2Ay(A() As S1S2, Optional Sep$ = " ") As Dictionary
Dim J&, O As New Dictionary
For J = 0 To UB(A)
    With A(J)
        If O.Exists(.S1) Then
            O(.S1) = O(.S1) & " " & O(.S2)
        Else
            O.Add .S1, .S2
        End If
    End With
Next
Set DiczS1S2Ay = O
End Function

Function Sy1zS1S2Ay(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S1
Next
Sy1zS1S2Ay = O
End Function

Function Sy2zS1S2Ay(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S2
Next
Sy2zS1S2Ay = O
End Function

Function SqzS1S2Ay(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), I, R&
ReDim O(1 To Sz(A), 1 To 2)
R = 2
O(1, 1) = Nm1
O(1, 2) = Nm2
For Each I In Itr(A)
    With CvS1S2(I)
        O(R, 1) = .S1
        O(R, 2) = .S2
        R = R + 1
    End With
Next
SqzS1S2Ay = O
End Function
Function S1S2AyzColonVbl(ColonVbl) As S1S2()
Dim I
For Each I In ItrVbl(ColonVbl)
    PushObj S1S2AyzColonVbl, BrkBoth(I, ":")
Next
End Function

Function S1S2AyzAySep(Ay, Sep$, Optional NoTrim As Boolean) As S1S2()
Dim O() As S1S2, J%
Dim U&: U = UB(Ay)
ReszAyU O, U
For J = 0 To U
    Set O(J) = Brk1(Ay(J), Sep, NoTrim)
Next
S1S2AyzAySep = O
End Function
Private Sub Z_S1S2AyzDic()
Dim A As New Dictionary
A.Add "A", "BB"
A.Add "B", "CCC"
Dim Act() As S1S2
Act = S1S2AyzDic(A)
Stop
End Sub


Private Sub Z()
Z_S1S2AyzDic
MVb__S1S2:
End Sub
