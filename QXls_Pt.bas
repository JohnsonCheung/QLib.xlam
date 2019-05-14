Attribute VB_Name = "QXls_Pt"
Option Explicit
Private Const CMod$ = "MXls_Pt."
Private Const Asm$ = "QXls"
Function PtCpyToLo(A As PivotTable, At As Range) As ListObject
Dim R1, R2, C1, C2, NC, NR
    R1 = A.RowRange.Row
    C1 = A.RowRange.Column
    R2 = LasRowRg(A.DataBodyRange)
    C2 = LasColRg(A.DataBodyRange)
    NC = C2 - C1 + 1
    NR = R2 - C1 + 1
WsRCRC(WsPt(A), R1, C1, R2, C2).Copy
At.PasteSpecial xlPasteValues

Set PtCpyToLo = CrtLozRg(RgRCRC(At, 1, 1, NR, NC))
End Function

Sub SetPtffOri(A As PivotTable, FF$, Ori As XlPivotFieldOrientation)
Dim F, J%, T
T = Array(False, False, False, False, False, False, False, False, False, False, False, False)
J = 1
For Each F In Itr(SyzSS(FF))
    With PivFld(A, F)
        .Orientation = Ori
        .Position = J
        If Ori = xlColumnField Or Ori = xlRowField Then
            .Subtotals = T
        End If
    End With
    J = J + 1
Next
End Sub

Private Sub FmtPt(Pt As PivotTable)

End Sub

Function WbNm$(A As Workbook)
On Error GoTo X
WbNm = A.FullName
Exit Function
X: WbNm = "WbNmErr:[" & Err.Description & "]"
End Function
Function LasWb() As Workbook
Set LasWb = LasWbz(Xls)
End Function
Function LasWbz(A As Excel.Application) As Workbook
Set LasWbz = A.Workbooks(A.Workbooks.Count)
End Function

Sub ThwHasWbzWs(Wb As Workbook, Wsn$, Fun$)
If HasWs(Wb, Wsn) Then
    Thw Fun, "Wb should have not have Ws", "Wb Ws", Wb.FullName, Wsn
End If
End Sub

Function PtzRg(A As Range, Optional Wsn$, Optional PtNm$) As PivotTable
Dim Wb As Workbook: Set Wb = WbzRg(A)
Dim Ws As Worksheet: Set Ws = AddWs(Wb)
Dim A1 As Range: Set A1 = A1zWs(Ws)
Dim Pc As PivotCache: Set Pc = WbzRg(A).PivotCaches.Create(xlDatabase, A.Address, Version:=6)
Dim Pt As PivotTable: Set Pt = Pc.CreatePivotTable(A1, DefaultVersion:=6)
End Function
Function PivCol(Pt As PivotTable, PivColNm) As PivotField

End Function
Function PivRow(Pt As PivotTable, PivRowNm) As PivotField
Set PivRow = Pt.ColumnFields(PivRowNm)
End Function
Function PivFld(A As PivotTable, F) As PivotField
Set PivFld = A.PageFields(F)
End Function
Function ColEntPt(A As PivotTable, PivColNm) As Range
Set ColEntPt = RgR(PivCol(A, PivColNm).DataRange, 1).EntireColumn
End Function
Function PivColEnt(Pt As PivotTable, ColNm) As Range
Set PivColEnt = PivCol(Pt, ColNm).EntireColumn
End Function

Sub SetPtWdt(A As PivotTable, Colss$, ColWdt As Byte)
If ColWdt <= 1 Then Stop
Dim C
For Each C In Itr(SyzSS(Colss))
    ColEntPt(A, C).ColumnWidth = ColWdt
Next
End Sub

Sub SetPtOutLin(A As PivotTable, Colss$, Optional Lvl As Byte = 2)
If Lvl <= 1 Then Stop
Dim F, C As VBComponent
For Each C In Itr(SyzSS(Colss))
    ColEntPt(A, F).OutlineLevel = Lvl
Next
End Sub

Sub SetPtRepeatLbl(A As PivotTable, Rowss$)
Dim F
For Each F In Itr(SyzSS(Rowss))
    PivFld(A, F).RepeatLabels = True
Next
End Sub

Sub ShwPt(A As PivotTable)
ShwXls A.Application
End Sub

Function WbPt(A As PivotTable) As Workbook
Set WbPt = WbzWs(WsPt(A))
End Function

Function WsPt(A As PivotTable) As Worksheet
Set WsPt = A.Parent
End Function
Function SampPt() As PivotTable
Set SampPt = PtzRg(SampRg)
End Function
Function SampRg() As Range
Set SampRg = ShwRg(PutSq(SampSq, NewA1))
End Function

Function PutSq(Sq(), At As Range) As Range
Dim O As Range
Set O = ResiRg(At, Sq)
CrtLozRg O
O.Value = Sq
Set PutSq = O
End Function

Function PtzLo(A As ListObject, At As Range, Rowss$, Dtass$, Optional Colss$, Optional Pagss$) As PivotTable
If WbzLo(A).FullName <> WbzRg(At).FullName Then Stop: Exit Function
Dim O As PivotTable
Set O = LoPc(A).CreatePivotTable(TableDestination:=At, TableName:=PtNmzLo(A))
With O
    .ShowDrillIndicators = False
    .InGridDropZones = False
    .RowAxisLayout xlTabularRow
End With
O.NullString = ""
SetPtffOri O, Rowss, xlRowField
SetPtffOri O, Colss, xlColumnField
SetPtffOri O, Pagss, xlPageField
SetPtffOri O, Dtass, xlDataField
Set PtzLo = O
End Function

Function PtNmzLo$(A As ListObject)

End Function
