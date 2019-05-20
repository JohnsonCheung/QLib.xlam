Attribute VB_Name = "QVb_Dta_Sq"
Option Explicit
Option Compare Text
Sub SetSqr(OSq(), Drv, Optional R = 1, Optional NoTxtSngQ As Boolean)
Dim J&
If NoTxtSngQ Then
    For J = 0 To UB(Drv)
        If IsStr(Drv(J)) Then
            OSq(R, J + 1) = QuoteSng(CStr(Drv(J)))
        Else
            OSq(R, J + 1) = Drv(J)
        End If
    Next
Else
    For J = 0 To UB(Drv)
        OSq(R, J + 1) = Drv(J)
    Next
End If
End Sub

Sub PushSq(OSq(), Sq())
Dim NR&: NR = UBound(OSq, 1) + UBound(Sq, 1)
Dim NC&: NC = UBound(OSq, 2)
Dim NC2&: NC2 = UBound(Sq, 2)
If NC <> NC2 Then Thw CSub, "NC of { OSq, Sq } are dif", "OSq-NC Sq-NC", NC, NC2
ReDim Preserve OSq(1 To NR, 1 To NC)
Dim R&, C&
For R = 1 To NC2
    For C = 1 To NC
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
Dim NC%, C%, R&, O
O = Sq
NC = UBound(Sq, 2)
For R = 1 To UBound(Sq, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
AddSngQuotezSq = O
End Function
Function JnSq(Sq$(), SepChr$) As String()
Dim NC&: NC = UBound(Sq, 2)
Dim R&
For R = 1 To UBound(Sq, 1)
    PushI JnSq, JnSqr(Sq, NC, R, SepChr)
Next
End Function

Private Function JnSqr$(Sq$(), NC&, R&, SepChr$)
JnSqr = Join(SyzSqr(Sq, NC, R), SepChr)
End Function

Function SyzSqr(Sq$(), NC&, R&) As String()
Dim J&
For J = 1 To NC
    PushI SyzSqr, Sq(R, J)
Next
End Function

Function FmtSq(Sq(), Optional SepChr$ = " ") As String()
FmtSq = JnSq(AlignSq(Sq), SepChr)
End Function

Sub BrwSq(Sq())
Brw FmtSq(Sq)
End Sub
Function AlignSq(Sq()) As String()
If Si(Sq) = 0 Then Exit Function
Dim C&, O$(), NR&, NC&
NR = UBound(Sq, 1)
NC = UBound(Sq, 2)
ReDim O(1 To NR, 1 To NC)
For C = 1 To UBound(O, 2)
    AlignColzSCW O, Sq, C, ColWdtzSC(Sq, C)
Next
AlignSq = O
End Function
Private Function ColWdtzSC%(Sq(), C&)
Dim R&, O%
For R = 1 To UBound(Sq, 1)
    O = Max(O, Len(Sq(R, C)))
Next
ColWdtzSC = O
End Function
Private Sub AlignColzSCW(OSq$(), Sq(), C&, W%)
Dim R&
For R = 1 To UBound(Sq, 1)
    OSq(R, C) = Align(Sq(R, C), W)
Next
End Sub
Function ColzSq(Sq(), C&) As Variant()
Dim O()
Dim NR&, J&
NR = UBound(Sq, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = Sq(J, C)
Next
ColzSq = O
End Function

Function IntozSqc(Into, Sq(), C&)
Dim NR&, O
    NR = UBound(Sq, 1)
    O = ResiN(Into, NR)
Dim R&
For R = 1 To NR
    O(R - 1) = Sq(R, C)
Next
IntozSqc = O
End Function
Function VyrzSqr(Sq(), Optional R& = 1) As Variant()
VyrzSqr = IntozSqr(EmpAv, Sq, R)
End Function
Function VyczSqc(Sq(), Optional C& = 1) As Variant()
VyczSqc = IntozSqc(EmpAv, Sq, C)
End Function

Function IntozSqr(Into, Sq(), R)
Dim NC&, O
    NC = UBound(Sq, 2)
    O = ResiN(Into, NC)
Dim C&
For C = 1 To NC
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
InsSqr = O
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
    Push LyzSq, TermAyzDr(DrzSqr(Sq, R))
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
Sq A, A
IsEqSq B, B
End Sub

Function SampSq() As Variant()
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
End Function


