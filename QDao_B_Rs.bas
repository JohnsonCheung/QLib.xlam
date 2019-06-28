Attribute VB_Name = "QDao_B_Rs"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Rs."
Sub UpdRs(Rs As Dao.Recordset, Dr)
Rs.Edit
SetRs Rs, Dr
Rs.Update
End Sub

Private Sub Y_AsgRs()
Dim Y As Byte, M As Byte
'AsgRs TblRs("YM"), Y, M
Stop
End Sub
Function CvRs(A) As Dao.Recordset
Set CvRs = A
End Function

Function NoRec(A As Dao.Recordset) As Boolean
If Not A.EOF Then Exit Function
If Not A.BOF Then Exit Function
NoRec = True
End Function

Function HasRec(A As Dao.Recordset) As Boolean
HasRec = Not NoRec(A)
End Function

Sub AsgRs(A As Dao.Recordset, ParamArray OAp())
Dim F As Dao.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In A.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub

Sub BrwRs(A As Dao.Recordset)
BrwDrs DrszRs(A)
End Sub

Sub BrwSngRec(A As Dao.Recordset)
BrwAy FmtRec(A)
End Sub

Sub RsDlt(A As Dao.Recordset)
With A
    If .EOF Then Exit Sub
    If .BOF Then Exit Sub
    .Delete
End With
End Sub

Function CsvLinzRs$(A As Dao.Recordset)
CsvLinzRs = CsvzFds(A.Fields)
End Function

Function CsvLyzRs1(A As Dao.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As Dao.Field
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

Function CsvLyzRs(A As Dao.Recordset, Optional FF$) As String()
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
Function AsetzRs(Rs As Dao.Recordset, F$) As Aset
Set AsetzRs = New Aset
With Rs
    While Not .EOF
        AsetzRs.PushItm .Fields(F).Value
        .MoveNext
    Wend
End With
End Function

Sub DmpRs(A As Recordset, FF$)
DmpAy CsvLyzRs(A, FF)
A.MoveFirst
End Sub

Function DrzRs(A As Dao.Recordset, Optional FF$) As Variant()
DrzRs = DrzFds(A.Fields, FF)
End Function

Function DrszRs(A As Dao.Recordset) As Drs
DrszRs = Drs(FnyzRs(A), DyoRs(A))
End Function

Function DyoRs(A As Dao.Recordset, Optional IsIncFldn As Boolean) As Variant()
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

Function FnyzRs(A As Dao.Recordset) As String()
FnyzRs = Itn(A.Fields)
End Function

Function HasReczFEv(Rs As Dao.Recordset, F, Ev) As Boolean
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

Function IntAyzRs(A As Dao.Recordset, Optional Fld = 0) As Integer()
IntAyzRs = IntozRs(IntAyzRs, A, Fld)
End Function

Function RsLin(A As Dao.Recordset, Optional Sep$ = " ")
RsLin = Join(DrzRs(A), Sep)
End Function

Function LngAyzRs(A As Dao.Recordset, Optional Fld = 0) As Long()
LngAyzRs = IntozRs(LngAyzRs, A, Fld)
End Function

Function RsLy(A As Dao.Recordset, Optional Sep$ = " ") As String()
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

Function FmtRec(A As Dao.Recordset)
FmtRec = LyzNyAv(FnyzRs(A), DrzRs(A))
End Function

Function NReczRs&(A As Dao.Recordset)
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


Function SyzRs(A As Dao.Recordset, Optional F = 0) As String()
Dim O$()
With A
    While Not .EOF
        Push O, .Fields(F).Value
        .MoveNext
    Wend
End With
SyzRs = O
End Function

Function StruzRs$(A As Dao.Recordset)
Dim O$(), F As Dao.Field2
For Each F In A.Fields
    PushI O, FdStr(F)
Next
StruzRs = JnCrLf(O)
End Function

Function EmptyIfNull(V)
EmptyIfNull = IIf(IsNull(V), Empty, V)
End Function

Function DrzRsFny(A As Dao.Recordset, Fny$()) As Variant()
Dim F
For Each F In Fny
    PushI DrzRsFny, EmptyIfNull(A.Fields(F).Value)
Next
End Function

Function IntozRs(Into, Rs As Recordset, Optional Fld = 0)
IntozRs = ResiU(Into)
While Not Rs.EOF
    PushI IntozRs, Nz(Rs(Fld).Value, Empty)
    Rs.MoveNext
Wend
End Function

Function AvRsCol(A As Dao.Recordset, Optional Fld = 0) As Variant()
AvRsCol = IntozRs(EmpAv, A, Fld)
End Function

Function ColSetzRs(A As Dao.Recordset, Optional Fld = 0) As Aset
Set ColSetzRs = New Aset
With A
    While Not .EOF
        ColSetzRs.PushItm .Fields(Fld).Value
        .MoveNext
    Wend
End With
End Function

Function SqzRs(A As Dao.Recordset, Optional ExlFldNm As Boolean) As Variant()
SqzRs = SqzDy(DyoRs(A, ExlFldNm))
End Function


Sub InsRszDy(A As Dao.Recordset, Dy())
Dim Dr
With A
    For Each Dr In Itr(Dy)
        InsRs A, Dr
    Next
End With
End Sub


Sub SetRs(Rs As Dao.Recordset, Dr)
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


Sub InsRszAp(Rs As Dao.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
InsRs Rs, Dr
End Sub

Sub InsRs(Rs As Dao.Recordset, Dr)
Rs.AddNew
SetRs Rs, Dr
Rs.Update
End Sub

Sub UpdRszAp(Rs As Dao.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
UpdRs Rs, Dr
End Sub


Sub DltRs(A As Dao.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub



