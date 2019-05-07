Attribute VB_Name = "QDao_Rs"
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Rs."
Sub UpdRs(Rs As DAO.Recordset, Dr)
Rs.Edit
SetRs Rs, Dr
Rs.Update
End Sub

Private Sub ZZ_AsgRs()
Dim Y As Byte, M As Byte
'AsgRs TblRs("YM"), Y, M
Stop
End Sub
Function CvRs(A) As DAO.Recordset
Set CvRs = A
End Function

Function NoRec(A As DAO.Recordset) As Boolean
If Not A.EOF Then Exit Function
If Not A.BOF Then Exit Function
NoRec = True
End Function

Function HasRec(A As DAO.Recordset) As Boolean
HasRec = Not NoRec(A)
End Function

Sub AsgRs(A As DAO.Recordset, ParamArray OAp())
Dim F As DAO.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In A.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub

Sub BrwRs(A As DAO.Recordset)
BrwDrs DrszRs(A)
End Sub

Sub BrwSngRec(A As DAO.Recordset)
BrwAy FmtRec(A)
End Sub

Sub RsDlt(A As DAO.Recordset)
With A
    If .EOF Then Exit Sub
    If .BOF Then Exit Sub
    .Delete
End With
End Sub

Function CsvLinzRs$(A As DAO.Recordset)
CsvLinzRs = CsvzFds(A.Fields)
End Function

Function CsvLyzRs1(A As DAO.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As DAO.Field
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

Function CsvLyzRs(A As DAO.Recordset, Optional FF$) As String()
Dim Fny$(), Flds As Fields, F
Dim O$(), J&, I%, UFld%, Dr()
Fny = TermSy(FF)
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
Function AsetzRs(Rs As DAO.Recordset, F$) As Aset
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

Function DrzRs(A As DAO.Recordset, Optional FF$) As Variant()
DrzRs = DrzFds(A.Fields, FF)
End Function
Function DrszRs(A As DAO.Recordset) As Drs
DrszRs = Drs(FnyzRs(A), DryzRs(A))
End Function

Function DryzRs(A As DAO.Recordset, Optional IsExlFldNm As Boolean) As Variant()
If Not HasRec(A) Then Exit Function
If Not IsExlFldNm Then
    PushI DryzRs, FnyzRs(A)
End If
With A
    .MoveFirst
    While Not .EOF
        PushI DryzRs, DrzFds(.Fields)
        .MoveNext
    Wend
    .MoveFirst
End With
End Function

Function FnyzRs(A As DAO.Recordset) As String()
FnyzRs = Itn(A.Fields)
End Function

Function HasReczFEv(Rs As DAO.Recordset, F, Ev) As Boolean
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

Function IntAyzRs(A As DAO.Recordset, Optional Fld = 0) As Integer()
IntAyzRs = IntozRs(IntAyzRs, A, Fld)
End Function

Function RsLin$(A As DAO.Recordset, Optional Sep$ = " ")
RsLin = Join(DrzRs(A), Sep)
End Function

Function LngAyzRs(A As DAO.Recordset, Optional Fld = 0) As Long()
LngAyzRs = IntozRs(LngAyzRs, A, Fld)
End Function

Function RsLy(A As DAO.Recordset, Optional Sep$ = " ") As String()
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

Function FmtRec(A As DAO.Recordset)
FmtRec = LyzNyAv(FnyzRs(A), DrzRs(A))
End Function

Function NReczRs&(A As DAO.Recordset)
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


Function SyzRs(A As DAO.Recordset, Optional F = 0) As String()
Dim O$()
With A
    While Not .EOF
        Push O, .Fields(F).Value
        .MoveNext
    Wend
End With
SyzRs = O
End Function

Function StruzRs$(A As DAO.Recordset)
Dim O$(), F As DAO.Field2
For Each F In A.Fields
    PushI O, FdStr(F)
Next
StruzRs = JnCrLf(O)
End Function

Function EmptyIfNull(V)
EmptyIfNull = IIf(IsNull(V), Empty, V)
End Function

Function DrzRsFny(A As DAO.Recordset, Fny$()) As Variant()
Dim F
For Each F In Fny
    PushI DrzRsFny, EmptyIfNull(A.Fields(F).Value)
Next
End Function

Function IntozRs(Into, Rs As Recordset, Optional Fld = 0)
IntozRs = Resi(Into)
While Not Rs.EOF
    PushI IntozRs, Nz(Rs(Fld).Value, Empty)
    Rs.MoveNext
Wend
End Function

Function AvRsCol(A As DAO.Recordset, Optional Fld = 0) As Variant()
AvRsCol = IntozRs(EmpAv, A, Fld)
End Function

Function ColSetzRs(A As DAO.Recordset, Optional Fld = 0) As Aset
Set ColSetzRs = New Aset
With A
    While Not .EOF
        ColSetzRs.PushItm .Fields(Fld).Value
        .MoveNext
    Wend
End With
End Function

Function SqzRs(A As DAO.Recordset, Optional ExlFldNm As Boolean) As Variant()
SqzRs = SqzDry(DryzRs(A, ExlFldNm))
End Function

