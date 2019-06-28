Attribute VB_Name = "QVb_Dta_Sq"
Option Explicit
Option Compare Text
Sub SetSqr(OSq(), Dr, Optional R = 1, Optional NoTxtSngQ As Boolean)
Dim J&
If NoTxtSngQ Then
    For J = 0 To UB(Dr)
        If IsStr(Dr(J)) Then
            OSq(R, J + 1) = QteSng(CStr(Dr(J)))
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

Function VzLStr(LStr)
'Ret : ! a val (Str|Dbl|Bool|Dte|Empty) fm @LStr.  :LStr: is #Letter-Str.  A str wi fst letter can-determine the str can converted to what value.
'      ! If fst letter is
'      !   ['] is a str wi \r\n\t
'      !   [D] is a str of date, if cannot convert to date, ret empty and debug.print msg.
'      !   [T] is true
'      !   [F] is false
'      !   rest is dbl, if cannot convert to dbl, ret empty and debug.print msg @@
Dim F$: F = FstChr(LStr)
Dim O$
Select Case F
Case "'": O = UnSlashCrLfTab(RmvFstChr(LStr))
Case "T": O = True
Case "F": O = False
Case "D": O = CvDte(RmvFstChr(LStr))
Case ""
Case Else: O = CvDbl(RmvFstChr(LStr))
End Select
VzLStr = O
End Function
Function Sq(R&, C&) As Variant()
'Ret : a Sq(1 to @R, 1 to @C)
Dim O()
ReDim O(1 To R, 1 To C)
Sq = O
End Function

Function AddSngQtezSq(Sq())
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
AddSngQtezSq = O
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
FmtSq = JnSq(SqzAlign(Sq), SepChr)
End Function

Sub BrwSq(Sq())
Brw FmtSq(Sq)
End Sub

Function SqzAlign(Sq()) As String()
If Si(Sq) = 0 Then Exit Function
Dim C&, O$(), NR&, NC&
NR = UBound(Sq, 1)
NC = UBound(Sq, 2)
ReDim O(1 To NR, 1 To NC)
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
Dim NC&:    NC = UBound(Sq, 2)
Dim O:    O = ResiN(Into, NC)
Dim C&: For C = 1 To NC
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

Private Sub Z()
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


Function CvDte(S, Optional Fun$)
'Ret : a date fm @S if can be converted, otherwise empty and debug.print @S
On Error GoTo X
Dim O As Date: O = S
If SubStrCnt(S, "/") <> 2 Then GoTo X ' ! one [/]-str is cv to yyyy/mm, which is not consider as a dte.
'                                       ! so use 2-[/] to treat as a dte str.
If Year(O) < 2000 Then GoTo X         ' ! year < 2000, treat it as str or not
CvDte = O
Exit Function
X: If Fun <> "" Then Inf CSub, "str[" & S & "] cannot cv to dte, emp is ret"
End Function
Private Sub Z_SqStr()
Brw SqStrzDrs(DoMthP)
End Sub
Function SqStrzDy$(Dy())
SqStrzDy = SqStr(SqzDy(Dy))
End Function
Function SqStrzDrs$(D As Drs)
SqStrzDrs = SqStrzDy(D.Dy)
End Function
Private Sub Z_SqStrzWs()
Dim S As Worksheet: Set S = WsMthP
Dim Bef$: Bef = SqStrzWs(S)
EnsSprp S, "A", Bef
Dim Aft$: Aft = Sprp(S, "A")
If Bef <> Aft Then Stop
VcStr Bef
End Sub
Function SqStrzWs$(S As Worksheet)
SqStrzWs = SqStrzLo(FstLo(S))
End Function
Function SqStrzLo$(L As ListObject)
SqStrzLo = SqStrzRg(L.DataBodyRange)
End Function

Function SqStr$(Sq())
'Ret : :SqStr:   ! it is Lines of xstr lin.
'      :XStrLin: ! #XStr-Lin is :XStr separaterd by vbTab with ending vbTab.
'                ! the fst XStr-Lin will not trim the ending vbTab, because it is used to determine how many col.
'                ! if any non-Lin1-XStr-Lin has more fld than lin1-XStr-lin-fld, the extra fld are ignored and inf (this is done in %SqS%
'                ! the reverse fun is %SqzS @@
Dim L$(), O$()
Dim UC%: UC = UBound(Sq, 1)
Dim R&: For R = 1 To UBound(Sq, 1)
    ReDim L(UC)
    Dim C&: For C = 1 To UBound(Sq, 2)
        L(C - 1) = XStr(Sq(R, C))
    Next
    PushI O, JnTab(L)
Next
SqStr = JnCrLf(O)
End Function
Function LStr$(V, Optional Fun$)
'Ret : :LStr from a val-@V.
Dim T$: T = TypeName(V)
Dim O$
Select Case T
Case "String": O = "'" & SlashCrLfTab(V)
Case "Boolean": O = IIf(V, "T", "F")
Case "Integer", "Single", "Double", "Currency", "Long": O = V
Case "Date": O = "D" & V
Case Else: If Fun <> "" Then Inf CSub, "Val-of-TypeName[" & T & "] cannot cv to :LStr"
End Select
LStr = O
End Function

Function XStr$(V, Optional Fun$)
'Ret : :XStr fm a val-@V.
Dim T$: T = TypeName(V)
Dim O$
Select Case T
Case "Boolean", "Long", "Integer", "Date", "Currency", "Single", "Double": XStr = V
Case "String": If IsDblStr(V) Then XStr = "'" & V Else XStr = SlashCrLfTab(V)
Case Else: If Fun <> "" Then Inf Fun, "Val-of-TypeName[" & T & "] cannot cv to :XStr"
End Select
End Function

Function SqzS(SqStr$) As Variant()
'Ret : a :Sq from :SqStr
Dim Ry$(): Ry = SplitCrLf(SqStr): If Si(Ry) = 0 Then Exit Function
Dim NR&: NR = Si(Ry)
Dim R1$: R1 = Ry(0)
Dim NC%: NC = Si(SplitTab(R1))
Dim O(): ReDim O(1 To NR, 1 To NC)
Dim IR&, IC%
Dim R: For Each R In Ry
    IR = IR + 1
    Dim C: For Each C In SplitTab(R)
        IC = IC + 1
        If IC > NC Then Exit For ' ign the extra fld, if it has more fld then lin1-fld-cnt
        O(IR, IC) = VzLStr(C)
    Next
Next
End Function
Function SqStrzRg$(R As Range)
SqStrzRg = SqStr(SqzRg(R))
End Function

