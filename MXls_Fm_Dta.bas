Attribute VB_Name = "MXls_Fm_Dta"
Option Explicit

Function RgzDrs(A As Drs, At As Range) As Range
Set RgzDrs = RgzSq(SqzDrs(A), At)
End Function

Function LozDrs(Drs As Drs, At As Range, Optional LoNm$) As ListObject
Set LozDrs = LozRg(RgzDrs(Drs, At), LoNm)
End Function

Function WszAy(Ay, Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet, R As Range
Set O = NewWs(Wsn)
O.Range("A1").Value = "Array"
Set R = RgzSq(SqzAyV(Ay), O.Range("A2"))
LozRg RgzMoreTop(R)
Set WszAy = O
End Function

Function WszDrs(Drs As Drs, Optional Wsn$ = "Sheet1", Optional Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs(Wsn)
LozDrs Drs, O.Range("A1")
Set WszDrs = SetWsVis(O, Vis)
End Function

Function RgzAyV(Ay, At As Range) As Range
Set RgzAyV = RgzSq(SqzAyV(Ay), At)
End Function

Function RgzAyH(Ay, At As Range) As Range
Set RgzAyH = RgzSq(SqzAyH(Ay), At)
End Function

Function RgzDry(Dry(), At As Range) As Range
Set RgzDry = RgzSq(SqzDry(Dry), At)
End Function

Function WszDry(Dry(), Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(Wsn)
RgzDry Dry, A1zWs(O)
Set WszDry = O
End Function

Function WbzDs(A As Ds) As Workbook
Dim O As Workbook
Set O = NewWb
With FstWs(O)
   .Name = "Ds"
   .Range("A1").Value = A.DsNm
End With
Dim I
For Each I In Itr(A.DtAy)
    WszWbDt O, CvDt(I)
Next
Set WbzDs = O
End Function

Function WszDs(A As Ds) As Worksheet
Dim O As Worksheet: Set O = NewWs
A1zWs(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = WsRC(O, 2, 1)
Dim I, BelowN&, Dt As Dt
For Each I In Itr(A.DtAy)
    Set Dt = I
    LozDt Dt, At
    BelowN = 2 + Sz(Dt.Dry)
    Set At = CellBelow(At, BelowN)
Next
Set WszDs = O
End Function

Function RgzDt(A As Dt, At As Range, Optional DtIx%)
Dim Pfx$: If DtIx > 0 Then Pfx = QuoteBkt(DtIx)
At.Value = Pfx & A.DtNm
RgzSq SqzDrs(DrszDt(A)), CellBelow(At)
End Function

Function LozDt(A As Dt, At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set LozDt = LozDrs(DrszDt(A), R)
RgRC(R, 0, 1).Value = A.DtNm
End Function

Function WszWbDt(Wb As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = AddWs(Wb, Dt.DtNm)
LozDrs DrszDt(Dt), A1zWs(O)
Set WszWbDt = O
End Function

Function RgzSq(Sq, At As Range) As Range
Dim O As Range
Set O = RgzResz(At, Sq)
O.MergeCells = False
O.Value = Sq
Set RgzSq = O
End Function

Private Sub ZZ_WszDs()
WsVis WszDs(SampDs)
End Sub

Private Sub ZZ()
Dim A As Drs
Dim B As Range
Dim C$()
Dim D$
Dim E As Variant
Dim F As Ds
Dim G As Dt
Dim H%
Dim I As Workbook
RgzDrs A, B
LozDrs A, B
RgzSq E, B
End Sub

Private Sub Z()
End Sub