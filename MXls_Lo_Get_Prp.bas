Attribute VB_Name = "MXls_Lo_Get_Prp"
Option Explicit
Const CMod$ = "MXls_Lo."

Sub AddFml(Lo As ListObject, ColNm$, Fml$)
Dim O As ListColumn
Set O = Lo.ListColumns.Add
O.Name = ColNm
O.DataBodyRange.Formula = Fml
End Sub
Function LoNm$(T)
LoNm = "T_" & RmvFstNonLetter(T)
End Function

Function CvLo(A) As ListObject
Set CvLo = A
End Function

Function LoAllCol(A As ListObject) As Range
Set LoAllCol = RgzLoCC(A, 1, LoNCol(A))
End Function

Function LoAllEntCol(A As ListObject) As Range
Set LoAllEntCol = LoAllCol(A).EntireColumn
End Function

Sub AutoFitLo(A As ListObject)
Dim C As Range: Set C = LoAllEntCol(A)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = EntRgC(C, J)
   If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
End Sub

Function BdrLoAround(A As ListObject)
Dim R As Range
Set R = RgzMoreTop(A.DataBodyRange)
If A.ShowTotals Then Set R = RgzMoreBelow(R)
BdrRgAround R
End Function

Sub BrwLo(A As ListObject)
BrwDrs DrszLo(A)
End Sub

Function RgzLoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = R1Lo(A, InclHdr)
R2 = R2Lo(A, InclTot)
mC1 = LoWsCno(A, C1)
mC2 = LoWsCno(A, C2)
Set RgzLoCC = WsRCRC(LoWs(A), R1, mC1, R2, mC2)
End Function

Function RgzLc(A As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R As Range
Set R = A.ListColumns(C).DataBodyRange
If Not InclTot And Not InclHdr Then
    Set RgzLc = R
    Exit Function
End If
If InclTot Then Set RgzLc = RgzMoreBelow(R, 1)
If InclHdr Then Set RgzLc = RgzMoreTop(R, 1)
End Function

Function LozWsDta(A As Worksheet, Optional LoNm$) As ListObject
Set LozWsDta = LozRg(RgzWs(A), LoNm)
End Function

Sub DltLo(A As ListObject)
Dim R As Range, R1, C1, R2, C2, Ws As Worksheet
Set Ws = LoWs(A)
Set R = RgzMoreBelow(RgzMoreTop(A.DataBodyRange))
AsgRRRCCRg R, R1, C1, R2, C2
A.QueryTable.Delete
WsRCRC(Ws, R1, C1, R2, C2).ClearContents
End Sub

Function DrszLo(A As ListObject) As Drs
Set DrszLo = Drs(FnyzLo(A), DryLo(A))
End Function
Function DryLo(A As ListObject) As Variant()
DryLo = DryzSq(SqzLo(A))
End Function
Function DryRgColAy(Rg As Range, ColIxAy) As Variant()
DryRgColAy = DryzSqCol(Rg.Value, ColIxAy)
End Function
Function DryRgzLoCC(Lo As ListObject, CC) As Variant() _
' Return as many column as columns in [CC] from Lo
DryRgzLoCC = DryRgColAy(Lo.DataBodyRange, IxAy(FnyzLo(Lo), CC))
End Function

Function DtaAdrzLo$(A As ListObject)
DtaAdrzLo = WsRgAdr(A.DataBodyRange)
End Function

Function EntColzLo(A As ListObject, C) As Range
Set EntColzLo = RgzLc(A, C).EntireColumn
End Function

Function FbtStrLo$(A As ListObject)
FbtStrLo = FbtStrQt(A.QueryTable)
End Function

Function FnyzLo(A As ListObject) As String()
FnyzLo = Itn(A.ListColumns)
End Function

Function HasLoC(Lo As ListObject, C) As Boolean
HasLoC = HasItn(Lo.ListColumns, C)
End Function

Function HasLoFny(A As ListObject, Fny$()) As Boolean
HasLoFny = HasEleAy(FnyzLo(A), Fny)
End Function

Function LoHasNoDta(A As ListObject) As Boolean
LoHasNoDta = IsNothing(A.DataBodyRange)
End Function

Function LoHdrCell(A As ListObject, FldNm) As Range
Dim Rg As Range: Set Rg = A.ListColumns(FldNm).Range
Set LoHdrCell = RgRC(Rg, 1, 1)
End Function

Sub LoKeepFstCol(A As ListObject)
Dim J%
For J = A.ListColumns.Count To 2 Step -1
    A.ListColumns(J).Delete
Next
End Sub

Sub LoKeepFstRow(A As ListObject)
Dim J%
For J = A.ListRows.Count To 2 Step -1
    A.ListRows(J).Delete
Next
End Sub

Function LoNCol%(A As ListObject)
LoNCol = A.ListColumns.Count
End Function


Function LoPc(A As ListObject) As PivotCache
Dim O As PivotCache
Set O = WbLo(A).PivotCaches.Create(xlDatabase, A.Name, 6)
O.MissingItemsLimit = xlMissingItemsNone
Set LoPc = O
End Function

Function LoQt(A As ListObject) As QueryTable
On Error Resume Next
Set LoQt = A.QueryTable
End Function

Function R1Lo&(A As ListObject, Optional InclHdr As Boolean)
If LoHasNoDta(A) Then
   R1Lo = A.ListColumns(1).Range.Row + 1
   Exit Function
End If
R1Lo = A.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function R2Lo&(A As ListObject, Optional InclTot As Boolean)
If LoHasNoDta(A) Then
   R2Lo = R1Lo(A)
   Exit Function
End If
R2Lo = A.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Function SqzLo(A As ListObject)
SqzLo = A.DataBodyRange.Value
End Function

Function WsLo(A As ListObject) As Worksheet
Set WsLo = A.Parent
End Function
Function WbLo(A As ListObject) As Workbook
Set WbLo = WbzWs(LoWs(A))
End Function

Function LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Function

Function LoWsCno%(A As ListObject, Col)
LoWsCno = A.ListColumns(Col).Range.Column
End Function

Function LoNmzTblNm$(TblNm)
LoNmzTblNm = "T_" & RmvFstNonLetter(TblNm)
End Function

Private Sub ZZ_LoKeepFstCol()
LoKeepFstCol LoVis(SampLo)
End Sub

Private Sub Z_AutoFitLo()
Dim Ws As Worksheet: Set Ws = NewWs
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "A"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "A")
Ws.Range("A1:B2").Value = Sq
AutoFitLo LozWsDta(Ws)
ClsWsNoSav Ws
End Sub

Private Sub Z_BrwLo()
BrwLo SampLo
Stop
End Sub

Private Sub Z_NewPtLoAtRDCP()
Dim At As Range, Lo As ListObject
Set Lo = SampLo
'Set At = RgVis(A1zWs(AddWs(WbLo(Lo))))
ShwPt NewPtLoAtRDCP(Lo, At, "A B", "C D", "F", "E")
Stop
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B As ListObject
Dim C As Boolean
Dim D As Worksheet
Dim E$
Dim F&
Dim G$()
Dim H As Range
CvLo A
LoAllCol B
LoAllEntCol B
AutoFitLo B
BdrLoAround B
RgzLoCC B, A, A, C, C
RgzLc B, A, C, C
RgzLc B, A
FbtStrLo B
FnyzLo B
HasLoC B, A
HasLoFny B, G
LoHasNoDta B
LoHdrCell B, A
LoKeepFstCol B
LoKeepFstRow B
LoNCol B
NewPtLoAtRDCPNm B
LoQt B
R1Lo B, C
R2Lo B, C
SqzLo B
LoVis B
WbLo B
LoWs B
LoWsCno B, A
LoNmzTblNm A
End Sub

Private Sub Z()
Z_AutoFitLo
End Sub
