Attribute VB_Name = "QVb_Dta_Sq"
Option Explicit
Option Compare Text
Sub SetSqr(OSq(), Dr, Optional R = 1, Optional NoTxtSngQ As Boolean)
Dim J&
If NoTxtSngQ Then
    For J = 0 To UB(Dr)
        If IsStr(Dr(J)) Then
            OSq(R, J + 1) = QuoteSng(CStr(Dr(J)))
        Else
            OSq(R, J + 1) = Dr(J)
        End If
    Next
Else
    For J = 0 To UB(Dr)
        OSq(R, J + 1) = Dr(J)
    Next
End If
End Sub

Sub PushSq(OSq(), Sq())
Dim NR&: NR = UBound(OSq, 1) + UBound(Sq, 1)
Dim Nc&: Nc = UBound(OSq, 2)
Dim NC2&: NC2 = UBound(Sq, 2)
If Nc <> NC2 Then Thw CSub, "NC of { OSq, Sq } are dif", "OSq-NC Sq-NC", Nc, NC2
ReDim Preserve OSq(1 To NR, 1 To Nc)
Dim R&, C&
For R = 1 To NC2
    For C = 1 To Nc
        OSq(R + NR, C) = Sq(R, C)
    Next
Next
End Sub

Function Sq(R&, C&) As Variant()
Dim O()
ReDim O(1 To R, 1 To C)
Sq = O
End Function

Function AddSngQuotezSq(Sq())
Dim Nc%, C%, R&, O
O = Sq
Nc = UBound(Sq, 2)
For R = 1 To UBound(Sq, 1)
    For C = 1 To Nc
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
AddSngQuotezSq = O
End Function
Function JnSq(Sq$(), SepChr$) As String()
Dim Nc&: Nc = UBound(Sq, 2)
Dim R&
For R = 1 To UBound(Sq, 1)
    PushI JnSq, JnSqr(Sq, Nc, R, SepChr)
Next
End Function

Private Function JnSqr$(Sq$(), Nc&, R&, SepChr$)
JnSqr = Join(SyzSqr(Sq, Nc, R), SepChr)
End Function

Function SyzSqr(Sq$(), Nc&, R&) As String()
Dim J&
For J = 1 To Nc
    PushI SyzSqr, Sq(R, J)
Next
End Function

Function FmtSq(Sq(), Optional SepChr$ = " ") As String()
FmtSq = JnSq(SqzAlign(Sq), SepChr)
End Function

Sub BrwSq(Sq())
Brw FmtSq(Sq)
End Sub

Function SqzAlign(Sq()) As String()
If Si(Sq) = 0 Then Exit Function
Dim C&, O$(), NR&, Nc&
NR = UBound(Sq, 1)
Nc = UBound(Sq, 2)
ReDim O(1 To NR, 1 To Nc)
For C = 1 To UBound(O, 2)
    AlignColzSq O, Sq, C, WdtzSqc(Sq, C)
Next
SqzAlign = O
End Function
Private Function WdtzSqc%(Sq(), C&)
Dim R&, O%
For R = 1 To UBound(Sq, 1)
    O = Max(O, Len(Sq(R, C)))
Next
WdtzSqc = O
End Function
Private Sub AlignColzSq(OSq$(), Sq(), C&, W%)
Dim R&
For R = 1 To UBound(Sq, 1)
    OSq(R, C) = Align(Sq(R, C), W)
Next
End Sub


Function ColzSq(Sq(), Optional C = 1) As Variant()
ColzSq = IntozSqc(EmpAv, Sq, C)
End Function

Function DrzSq(Sq(), Optional C = 1) As Variant()
DrzSq = IntozSqc(EmpAv, Sq, C)
End Function

Function IntozSqc(Into, Sq(), C)
Dim NR&: NR = UBound(Sq, 1)
Dim O:    O = ResiN(Into, NR)
Dim R&: For R = 1 To NR
    O(R - 1) = Sq(R, C)
Next
IntozSqc = O
End Function

Function IntozSqr(Into, Sq(), R)
Dim Nc&:    Nc = UBound(Sq, 2)
Dim O:    O = ResiN(Into, Nc)
Dim C&: For C = 1 To Nc
    O(C - 1) = Sq(R, C)
Next
IntozSqr = O
End Function

Function SyzSq(Sq(), Optional C& = 1) As String()
SyzSq = IntozSqc(EmpSy, Sq(), C)
End Function

Function DrzSqr(Sq(), Optional R = 1) As Variant()
DrzSqr = IntozSqr(EmpAv, Sq, R)
End Function

Function InsSqr(Sq(), Dr(), Optional Row& = 1)
Dim O(), C&, R&, Nc&, NR&
Nc = NColzSq(Sq)
NR = NRowzSq(Sq)
ReDim O(1 To NR + 1, 1 To Nc)
For R = 1 To Row - 1
    For C = 1 To Nc
        O(R, C) = Sq(R, C)
    Next
Next
For C = 1 To Nc
    O(Row, C) = Dr(C - 1)
Next
For R = NR To Row Step -1
    For C = 1 To Nc
        O(R + 1, C) = Sq(R, C)
    Next
Next
InsSqr = O
End Function

Function IsEqSq(A, B) As Boolean
Dim NR&, Nc&
NR = UBound(A, 1)
Nc = UBound(A, 2)
If NR <> UBound(B, 1) Then Exit Function
If Nc <> UBound(B, 2) Then Exit Function
Dim R&, C&
For R = 1 To NR
    For C = 1 To Nc
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
    Push LyzSq, TermAyzDr(DrzSqr(Sq, R))
Next
End Function

Function NColzSq&(Sq())
On Error Resume Next
NColzSq = UBound(Sq, 2)
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
On Error Resume Next
NRowzSq = UBound(Sq, 1)
End Function


Function Transpose(Sq()) As Variant()
Dim NR&, Nc&
NR = NRowzSq(Sq): If NR = 0 Then Exit Function
Nc = NColzSq(Sq): If Nc = 0 Then Exit Function
Dim O(), J&, I&
ReDim O(1 To Nc, 1 To NR)
For J = 1 To NR
    For I = 1 To Nc
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
Sq A, A
IsEqSq B, B
End Sub

Function SampSq() As Variant()
Const NR% = 10
Const Nc% = 10
Dim O(), R%, C%
ReDim O(1 To NR, 1 To Nc)
SampSq = O
For R = 1 To NR
    For C = 1 To Nc
        O(R, C) = R * 1000 + C
    Next
Next
SampSq = O
End Function


