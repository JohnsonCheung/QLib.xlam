Attribute VB_Name = "MxRs"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxRs."

Function AsetzRs(Rs As dao.Recordset, F$) As Aset
Set AsetzRs = New Aset
With Rs
    While Not .EOF
        AsetzRs.PushItm .Fields(F).Value
        .MoveNext
    Wend
End With
End Function

Sub AsgRs(A As dao.Recordset, ParamArray OAp())
Dim F As dao.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In A.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub

Function AvRsCol(A As dao.Recordset, Optional Fld = 0) As Variant()
AvRsCol = IntozRs(EmpAv, A, Fld)
End Function

Sub BrwRs(A As dao.Recordset)
BrwDrs DrszRs(A)
End Sub

Sub BrwSngRec(A As dao.Recordset)
BrwAy FmtRec(A)
End Sub

Function ColSetzRs(A As dao.Recordset, Optional Fld = 0) As Aset
Set ColSetzRs = New Aset
With A
    While Not .EOF
        ColSetzRs.PushItm .Fields(Fld).Value
        .MoveNext
    Wend
End With
End Function

Function CsvLinzRs$(A As dao.Recordset)
CsvLinzRs = CsvzFds(A.Fields)
End Function

Function CsvLyzRs(A As dao.Recordset, Optional FF$) As String()
Dim Fny$(), Flds As Fields, F
Dim O$(), J&, I%, UFld%, Dr()
Fny = TermAy(FF)
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "CsvLinzRsLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    Set Flds = A.Fields
    For Each F In Fny
        Dr(I) = CvCsv(Flds(F).Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
CsvLyzRs = O
End Function

Function CsvLyzRs1(A As dao.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As dao.Field
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "CsvLinzRsLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    PushI CsvLyzRs1, CsvLinzRs(A)
    A.MoveNext
Wend
End Function

Function CvRs(A) As dao.Recordset
Set CvRs = A
End Function

Sub DltRs(A As dao.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub

Sub DmpRs(A As Recordset, FF$)
DmpAy CsvLyzRs(A, FF)
A.MoveFirst
End Sub

Function DrszRs(A As dao.Recordset) As Drs
DrszRs = Drs(FnyzRs(A), DyoRs(A))
End Function

Function DrzRs(A As dao.Recordset, Optional FF$) As Variant()
DrzRs = DrzFds(A.Fields, FF)
End Function

Function DrzRsFny(A As dao.Recordset, Fny$()) As Variant()
Dim F
For Each F In Fny
    PushI DrzRsFny, EmptyIfNull(A.Fields(F).Value)
Next
End Function

Function DyoRs(A As dao.Recordset, Optional IsIncFldn As Boolean) As Variant()
If IsIncFldn Then
    PushI DyoRs, FnyzRs(A)
End If
If Not HasRec(A) Then Exit Function
With A
    .MoveFirst
    While Not .EOF
        PushI DyoRs, DrzFds(.Fields)
        .MoveNext
    Wend
    .MoveFirst
End With
End Function

Function EmptyIfNull(V)
EmptyIfNull = IIf(IsNull(V), Empty, V)
End Function

Function FmtRec(A As dao.Recordset)
FmtRec = LyzNyAv(FnyzRs(A), DrzRs(A))
End Function

Function FnyzRs(A As dao.Recordset) As String()
FnyzRs = Itn(A.Fields)
End Function

Function HasRec(A As dao.Recordset) As Boolean
HasRec = Not NoRec(A)
End Function

Function HasReczFEv(Rs As dao.Recordset, F, Ev) As Boolean
With Rs
    If .BOF Then
        If .EOF Then Exit Function
    End If
    .MoveFirst
    While Not .EOF
        If .Fields(F) = Ev Then HasReczFEv = True: Exit Function
        .MoveNext
    Wend
End With
End Function

Sub InsRs(Rs As dao.Recordset, Dr)
Rs.AddNew
SetRs Rs, Dr
Rs.Update
End Sub

Sub InsRszAp(Rs As dao.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
InsRs Rs, Dr
End Sub

Sub InsRszDy(A As dao.Recordset, Dy())
Dim Dr
With A
    For Each Dr In Itr(Dy)
        InsRs A, Dr
    Next
End With
End Sub

Function IntAyzRs(A As dao.Recordset, Optional Fld = 0) As Integer()
IntAyzRs = IntozRs(IntAyzRs, A, Fld)
End Function

Function IntozRs(Into, Rs As Recordset, Optional Fld = 0)
IntozRs = ResiU(Into)
While Not Rs.EOF
    PushI IntozRs, Nz(Rs(Fld).Value, Empty)
    Rs.MoveNext
Wend
End Function

Function LngAyzRs(A As dao.Recordset, Optional Fld = 0) As Long()
LngAyzRs = IntozRs(LngAyzRs, A, Fld)
End Function

Function NoRec(A As dao.Recordset) As Boolean
If Not A.EOF Then Exit Function
If Not A.BOF Then Exit Function
NoRec = True
End Function

Function NReczRs&(A As dao.Recordset)
Dim O&
With A
    .MoveFirst
    While Not .EOF
        O = O + 1
        .MoveNext
    Wend
    .MoveFirst
End With
NReczRs = O
End Function

Sub RsDlt(A As dao.Recordset)
With A
    If .EOF Then Exit Sub
    If .BOF Then Exit Sub
    .Delete
End With
End Sub

Function RsLin(A As dao.Recordset, Optional Sep$ = " ")
RsLin = Join(DrzRs(A), Sep)
End Function

Function RsLy(A As dao.Recordset, Optional Sep$ = " ") As String()
Dim O$()
With A
    Push O, Join(FnyzRs(A), Sep)
    While Not .EOF
        Push O, RsLin(A, Sep)
        .MoveNext
    Wend
End With
RsLy = O
End Function

Sub SetRs(Rs As dao.Recordset, Dr)
If Si(Dr) <> Rs.Fields.Count Then
    Thw CSub, "Si of Rs & Dr are diff", _
        "Si-Rs and Si-Dr Rs-Fny Dr", Rs.Fields.Count, Si(Dr), Itn(Rs.Fields), Dr
End If
Dim V, J%
For Each V In Dr
    If IsEmpty(V) Then
        Rs(J).Value = Rs(J).DefaultValue
    Else
        Rs(J).Value = V
    End If
    J = J + 1
Next
End Sub

Function SqzRs(A As dao.Recordset, Optional ExlFldNm As Boolean) As Variant()
SqzRs = SqzDy(DyoRs(A, ExlFldNm))
End Function


Function SyzRs(A As dao.Recordset, Optional F = 0) As String()
Dim O$()
With A
    While Not .EOF
        Push O, .Fields(F).Value
        .MoveNext
    Wend
End With
SyzRs = O
End Function

Sub UpdRs(Rs As dao.Recordset, Dr)
Rs.Edit
SetRs Rs, Dr
Rs.Update
End Sub

Sub UpdRszAp(Rs As dao.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
UpdRs Rs, Dr
End Sub

Private Sub Y_AsgRs()
Dim Y As Byte, M As Byte
'AsgRs TblRs("YM"), Y, M
Stop
End Sub