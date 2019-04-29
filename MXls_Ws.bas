Attribute VB_Name = "MXls_Ws"
Option Explicit
Sub ShwWs(A As Worksheet)
A.Application.Visible = True
End Sub
Function AddWs(A As Workbook, Optional Wsn$, Optional AtBeg As Boolean, Optional AtEnd As Boolean, Optional BefWsn$, Optional AftWsn$) As Worksheet
Dim O As Worksheet
DltWs A, Wsn
Select Case True
Case AtBeg:         Set O = A.Sheets.Add(FstWs(A))
Case AtEnd:         Set O = A.Sheets.Add(LasWs(A))
Case BefWsn <> "": Set O = A.Sheets.Add(A.Sheets(BefWsn))
Case AftWsn <> "": Set O = A.Sheets.Add(, A.Sheets(AftWsn))
Case Else:          Set O = A.Sheets.Add
End Select
End Function

Function WsC(A As Worksheet, C) As Range
Dim R As Range
Set R = A.Columns(C)
Set WsC = R.EntireColumn
End Function

Function WsCC(A As Worksheet, C1, C2) As Range
Set WsCC = WsRCC(A, 1, C1, C2).EntireColumn
End Function

Sub DltLo(A As Worksheet)
Dim Ay() As ListObject, J%
Ay = IntozItr(Ay, A.ListObjects)
For J = 0 To UB(Ay)
    Ay(J).Delete
Next
End Sub

Sub ClsWsNoSav(A As Worksheet)
ClsWbNoSav WbzWs(A)
End Sub

Property Get CurWs() As Worksheet
Set CurWs = Xls.ActiveSheet
End Property

Function WsCRR(A As Worksheet, C, R1, R2) As Range
Set WsCRR = WsRCRC(A, R1, C, R2, C)
End Function

Function DltWs(A As Workbook, WsIx) As Boolean
If HasWs(A, WsIx) Then WszWb(A, WsIx).Delete: Exit Function
DltWs = True
End Function

Function RgzWs(A As Worksheet) As Range
Dim R, C
With LasCell(A)
   R = .Row
   C = .Column
End With
Set RgzWs = WsRCRC(A, 1, 1, R, C)
End Function

Function HasLo(A As Worksheet, LoNm$) As Boolean
HasLo = HasItn(A.ListObjects, LoNm)
End Function

Function LasCell(A As Worksheet) As Range
Set LasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function LasCno%(A As Worksheet)
LasCno = LasCell(A).Column
End Function

Function LasRno&(A As Worksheet)
LasRno = LasCell(A).Row
End Function

Function PtNyzWs(A As Worksheet) As String()
PtNyzWs = Itn(A.PivotTables)
End Function

Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function

Function WsRCC(A As Worksheet, R, C1, C2) As Range
Set WsRCC = WsRCRC(A, R, C1, R, C2)
End Function

Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function

Function WsRR(A As Worksheet, R1, R2) As Range
Set WsRR = A.Range(WsRC(A, R1, 1), WsRC(A, R2, 1)).EntireRow
End Function

Function SetWsNm(A As Worksheet, Nm$) As Worksheet
If Nm <> "" Then
    If HasWs(WbzWs(A), Nm) Then A.Name = Nm
End If
Set SetWsNm = A
End Function

Function SqzWs(A As Worksheet) As Variant()
SqzWs = RgzWs(A).Value
End Function
Function SetVisOfWs(A As Worksheet, Vis As Boolean) As Worksheet
A.Application.Visible = Vis
Set SetVisOfWs = A
End Function
Function A1zWb(A As Workbook, Optional NewWsn$) As Range ' Return A1 of a new Ws (with NewWsn) in Wb
Set A1zWb = A1zWs(AddWs(A, NewWsn))
End Function

Function A2zWs(A As Worksheet) As Range
Set A2zWs = A.Range("A2")
End Function

Function A1zWs(A As Worksheet) As Range
Set A1zWs = A.Range("A1")
End Function

Function CvWs(A) As Worksheet
Set CvWs = A
End Function

Function WbzWs(A As Worksheet) As Workbook
Set WbzWs = A.Parent
End Function

Function WbNmzWs$(A As Worksheet)
WbNmzWs = WbzWs(A).FullName
End Function

Sub DltColFm(Ws As Worksheet, FmCol)
WsCC(Ws, FmCol, LasCno(Ws)).Delete
End Sub
Sub DltRowFm(Ws As Worksheet, FmRow)
WsRR(Ws, FmRow, LasRno(Ws)).Delete
End Sub
Sub HidColFm(Ws As Worksheet, FmCol)
WsCC(Ws, FmCol, MaxWsCol).Hidden = True
End Sub

Sub HidRowFm(Ws As Worksheet, FmRow&)
WsRR(Ws, FmRow, MaxWsRow).EntireRow.Hidden = True
End Sub

Function CnozBefFstHid%(Ws As Worksheet)
Dim J%, O%
For J% = 1 To MaxWsCol
    If WsC(Ws, J).Hidden Then CnozBefFstHid = J - 1: Exit Function
Next
Stop
End Function


