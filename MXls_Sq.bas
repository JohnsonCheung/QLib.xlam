Attribute VB_Name = "MXls_Sq"
Option Explicit

Function NewSq(R&, C&) As Variant()
Dim O()
ReDim O(1 To R, 1 To C)
NewSq = O
End Function

Function SqAddSngQuote(A)
Dim NC%, C%, R&, O
O = A
NC = UBound(A, 2)
For R = 1 To UBound(A, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
SqAddSngQuote = O
End Function

Sub BrwSq(Sq())
BrwDry DryzSq(Sq)
End Sub

Function Sq_Col(A, C%) As Variant()
Dim O()
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C)
Next
Sq_Col = O
End Function

Function IntozSqC(Into, Sq(), C&)
Dim O
O = AyCln(Into)
Erase O
Dim NR&, R&
NR = UBound(Sq, 1)
ReDim O(NR - 1)
For R = 1 To NR
    O(R - 1) = Sq(R, C)
Next
IntozSqC = O
End Function

Function SyzSq(Sq(), Optional C& = 1) As String()
SyzSq = IntozSqC(EmpSy, Sq(), C)
End Function

Function DrzSqR(Sq(), R&) As Variant()
Dim C%, O()
ReDim O(UBound(Sq(), 2) - 1)
For C = 1 To UBound(Sq(), 2)
    O(C - 1) = Sq(R, C)
Next
DrzSqR = O
End Function

Function SqInsDr(Sq(), Dr(), Optional Row& = 1)
Dim O(), C&, R&, NC&, NR&
NC = NColzSq(Sq)
NR = NRowzSq(Sq)
ReDim O(1 To NR + 1, 1 To NC)
For R = 1 To Row - 1
    For C = 1 To NC
        O(R, C) = Sq(R, C)
    Next
Next
For C = 1 To NC
    O(Row, C) = Dr(C - 1)
Next
For R = NR To Row Step -1
    For C = 1 To NC
        O(R + 1, C) = Sq(R, C)
    Next
Next
SqInsDr = O
End Function

Function IsEmpSq(Sq) As Boolean
IsEmpSq = True
On Error GoTo X
Dim A
If UBound(Sq(), 1) < 0 Then Exit Function
If UBound(Sq(), 2) < 0 Then Exit Function
IsEmpSq = False
Exit Function
X:
End Function

Function IsEqSq(A, B) As Boolean
Dim NR&, NC&
NR = UBound(A, 1)
NC = UBound(A, 2)
If NR <> UBound(B, 1) Then Exit Function
If NC <> UBound(B, 2) Then Exit Function
Dim R&, C&
For R = 1 To NR
    For C = 1 To NC
        If A(R, C) <> B(R, C) Then
            Exit Function
        End If
    Next
Next
IsEqSq = True
End Function

Function LyzSq(Sq()) As String()
Dim R&
For R = 1 To UBound(Sq(), 1)
    Push LyzSq, TermAyzDr(DrzSqR(Sq, R))
Next
End Function

Function NColzSq&(Sq())
On Error Resume Next
NColzSq = UBound(Sq(), 2)
End Function
Function NewLoSqAt(Sq(), At As Range) As ListObject
Set NewLoSqAt = LozRg(RgzSq(Sq(), At))
End Function
Function NewLoSq(Sq(), Optional Wsn$ = "Data") As ListObject
Set NewLoSq = NewLoSqAt(Sq(), NewA1(Wsn))
End Function

Function WszSq(Sq(), Optional Wsn$) As Worksheet
Set WszSq = LozRg(RgzSq(Sq(), NewA1(Wsn)))
End Function

Function NRowzSq&(Sq())
NRowzSq = UBound(Sq(), 1)
End Function


Function Transpose(Sq()) As Variant()
Dim NR&, NC&
NR = NRowzSq(Sq): If NR = 0 Then Exit Function
NC = NColzSq(Sq): If NC = 0 Then Exit Function
Dim O(), J&, I&
ReDim O(1 To NC, 1 To NR)
For J = 1 To NR
    For I = 1 To NC
        O(I, J) = Sq(J, I)
    Next
Next
Transpose = O
End Function

Private Sub ZZ()
Dim A&
Dim B As Variant
Dim C$
Dim D%
Dim E&()
Dim F As ListObject
NewSq A, A
Sq_Col B, D
IsEmpSq B
IsEqSq B, B
End Sub

Property Get SampSq() As Variant()
Const NR% = 10
Const NC% = 10
Dim O(), R%, C%
ReDim O(1 To NR, 1 To NC)
SampSq = O
For R = 1 To NR
    For C = 1 To NC
        O(R, C) = R * 1000 + C
    Next
Next
SampSq = O
End Property

