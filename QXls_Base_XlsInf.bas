Attribute VB_Name = "QXls_Base_XlsInf"
Option Compare Text
Option Explicit
Public Const XlsPgmFfn$ = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
Public Const DoczLowZ$ = "z when used in Nm, it has special meaning.  It can occur in Cml las-one, las-snd, las-thrid chr, else it is er."
Public Const DoczNmBrk$ = "NmBrk is z or zx or zxx where z is letter-z and x is lowcase or digit.  NmBrk must be sfx of a cml."
Public Const DoczNmBrk_za$ = "It means `and`."
Enum EmAddgWay ' Adding data to ws way
    EiWcWay
    EiSqWay
End Enum
Function WszWc(Wc As WorkbookConnection) As Worksheet
Dim Wb As Workbook, Ws As Worksheet
Set Wb = Wc.Parent
Set Ws = AddWs(Wb, Wc.Name)
PutWc Wc, A1zWs(Ws)
Set WszWc = Ws
End Function

Function WsnzLo$(A As ListObject)
WsnzLo = WszLo(A).Name
End Function

Function CWs() As Worksheet
Set CWs = Xls.ActiveSheet
End Function

Sub ClrDtaRg(A As Worksheet)
DtaRgzWs(A).Clear
End Sub
Function DtaRgzWs(Ws As Worksheet) As Range
Set DtaRgzWs = Ws.Range(A1zWs(Ws), LasCell(Ws))
End Function

Function A1zLo(Lo As ListObject) As Range
Set A1zLo = RgRC(Lo.ListColumns(1).Range, 2, 1)
End Function

Function A1zWb(A As Workbook, Optional NewWsn$) As Range ' Return A1 of a new Ws (with NewWsn) in Wb
Set A1zWb = A1zWs(AddWs(A, NewWsn))
End Function


Function A1zWs(A As Worksheet) As Range
Set A1zWs = A.Range("A1")
End Function

Function HasLo(A As Worksheet, Lon$) As Boolean
HasLo = HasItn(A.ListObjects, Lon)
End Function


Function LasCell(A As Worksheet) As Range
Set LasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function LasRno&(A As Worksheet)
LasRno = LasCell(A).Row
End Function


Function LasCno%(A As Worksheet)
LasCno = LasCell(A).Column
End Function



Function RgzWs(A As Worksheet) As Range
Dim R, C
With LasCell(A)
   R = .Row
   C = .Column
End With
Set RgzWs = WsRCRC(A, 1, 1, R, C)
End Function


Function PtNyzWs(A As Worksheet) As String()
PtNyzWs = Itn(A.PivotTables)
End Function




Function WbzWs(A As Worksheet) As Workbook
Set WbzWs = A.Parent
End Function


Function WsC(A As Worksheet, C) As Range
Dim R As Range
Set R = A.Columns(C)
Set WsC = R.EntireColumn
End Function


Function WsCC(A As Worksheet, C1, C2) As Range
Set WsCC = WsRCC(A, 1, C1, C2).EntireColumn
End Function


Function WsCRR(A As Worksheet, C, R1, R2) As Range
Set WsCRR = WsRCRC(A, R1, C, R2, C)
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


Function SqzWs(A As Worksheet) As Variant()
SqzWs = RgzWs(A).Value
End Function

Property Get CWb() As Workbook
Set CWb = Xls.ActiveWorkbook
End Property

Function CvWbs(A) As Workbooks
Set CvWbs = A
End Function

Function LasWs(A As Workbook) As Worksheet
Set LasWs = A.Sheets(A.Sheets.Count)
End Function



Function CvWb(A) As Workbook
Set CvWb = A
End Function

Function FstWs(A As Workbook) As Worksheet
Set FstWs = A.Sheets(1)
End Function

Function FxzWb$(A As Workbook)
Dim F$
F = A.FullName
If F = A.Name Then Exit Function
FxzWb = F
End Function



Function MainLo(A As Workbook) As ListObject
Dim O As Worksheet, Lo As ListObject
Set O = MainWs(A):              If IsNothing(O) Then Exit Function
Set MainLo = LozWs(O, "T_Main")
End Function

Function MainQt(A As Workbook) As QueryTable
Dim Lo As ListObject
Set Lo = MainLo(A): If IsNothing(A) Then Exit Function
Set MainQt = Lo.QueryTable
End Function


Function MainWs(A As Workbook) As Worksheet
Set MainWs = WszCdNm(A, "WsOMain")
End Function



Function PtNy(A As Workbook) As String()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushIAy PtNy, PtNyzWs(Ws)
Next
End Function

Function TxtWc(A As Workbook) As TextConnection
Dim C As WorkbookConnection
For Each C In A.Connections
    If Not IsNothing(TxtCnzWc(C)) Then
        Set TxtWc = C.TextConnection
        Exit Function
    End If
Next
Stop
'XHalt_Impossible CSub
End Function


Function TxtWcCnt%(A As Workbook)
Dim C As WorkbookConnection, Cnt%
For Each C In A.Connections
    If Not IsNothing(TxtCnzWc(C)) Then Cnt = Cnt + 1
Next
TxtWcCnt = Cnt
End Function

Function TxtWcStr$(A As Workbook)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = TxtWc(A)
If IsNothing(T) Then Exit Function
TxtWcStr = T.Connection
End Function

Function WcyzOle(A As Workbook) As OLEDBConnection()
Dim O() As OLEDBConnection, Wc As WorkbookConnection
For Each Wc In A.Connections
    PushObjzExlNothing O, Wc.OLEDBConnection
Next
WcyzOle = OyeNothing(IntozItrPrp(WcyzOle, A.Connections, "OLEDBConnection"))
End Function

Function WcnyzWb(A As Workbook) As String()
WcnyzWb = Itn(A.Connections)
End Function


Function WcsyzWbOLE(A As Workbook) As String()
WcsyzWbOLE = SyzOyPrp(WcyzOle(A), PrpPth("Connection"))
End Function

Function WszWb(A As Workbook, WsIx) As Worksheet
Set WszWb = A.Sheets(WsIx)
End Function

Function WsnyzRg(A As Range) As String()
WsnyzRg = Wsny(WbzRg(A))
End Function

Function Wsny(A As Workbook) As String()
Wsny = Itn(A.Sheets)
End Function


Function WszCdNm(A As Workbook, WsCdNm$) As Worksheet
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws.CodeName = WsCdNm Then Set WszCdNm = Ws: Exit Function
Next
End Function

Function WsCdNy(A As Workbook) As String()
WsCdNy = SyzItrPrp(A.Sheets, "CodeName")
End Function


Function WbFullNm$(A As Workbook)
On Error Resume Next
WbFullNm = A.FullName
End Function

Private Sub Z_XlsOfGetObj()
Debug.Print XlsOfGetObj.Name
End Sub

Function XlsOfGetObj() As Excel.Application
Set XlsOfGetObj = GetObject(XlsPgmFfn)
End Function

Function Xls() As Excel.Application
Set Xls = Excel.Application
End Function

Function HasAddinFn(A As Excel.Application, AddinFn$) As Boolean
HasAddinFn = HasItn(A.AddIns, AddinFn)
End Function
Function DftXls(A As Excel.Application) As Excel.Application
If IsNothing(A) Then
    Set DftXls = NewXls
Else
    Set DftXls = A
End If
End Function



Function NRowzLo&(A As ListObject)
NRowzLo = A.DataBodyRange.Rows.Count
End Function

Function Lon$(T)
Lon = "T_" & RmvFstNonLetter(T)
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


Function DrszLo(A As ListObject) As Drs
DrszLo = Drs(FnyzLo(A), DryzLo(A))
End Function
Function DryzLo(A As ListObject) As Variant()
DryzLo = DryzSq(SqzLo(A))
End Function
Function DryRgColAy(Rg As Range, ColIxy) As Variant()
DryRgColAy = DryzSqCol(SqzRg(Rg), ColIxy)
End Function
Function DryRgzLoCC(Lo As ListObject, CC) As Variant() _
' Return as many column as columns in [CC] from Lo
DryRgzLoCC = DryRgColAy(Lo.DataBodyRange, Ixy(FnyzLo(Lo), CC))
End Function

Function DtaAdrzLo$(A As ListObject)
DtaAdrzLo = WsRgAdr(A.DataBodyRange)
End Function

Function EntColzLo(A As ListObject, C) As Range
Set EntColzLo = RgzLc(A, C).EntireColumn
End Function

Function RgzLoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = R1Lo(A, InclHdr)
R2 = R2Lo(A, InclTot)
mC1 = WsCnozLc(A, C1)
mC2 = WsCnozLc(A, C2)
Set RgzLoCC = WsRCRC(WszLo(A), R1, mC1, R2, mC2)
End Function

Function LozWsDta(A As Worksheet, Optional Lon$) As ListObject
Set LozWsDta = LozRg(RgzWs(A), Lon)
End Function

Function FbtStrzLo$(A As ListObject)
FbtStrzLo = FbtStrzQt(A.QueryTable)
End Function

Function FnyzLo(A As ListObject) As String()
FnyzLo = Itn(A.ListColumns)
End Function

Function HasLoC(Lo As ListObject, ColNm$) As Boolean
HasLoC = HasItn(Lo.ListColumns, ColNm)
End Function

Function IsLozNoDta(A As ListObject) As Boolean
IsLozNoDta = IsNothing(A.DataBodyRange)
End Function

Function HdrCellzLo(A As ListObject, FldNm) As Range
Dim Rg As Range: Set Rg = A.ListColumns(FldNm).Range
Set HdrCellzLo = RgRC(Rg, 1, 1)
End Function

Function LoNCol%(A As ListObject)
LoNCol = A.ListColumns.Count
End Function
Function LozWs(A As Worksheet, Lon$) As ListObject 'Return LoOpt
Set LozWs = FstItmzNm(A.ListObjects, Lon)
End Function

Function FstLo(A As Worksheet) As ListObject 'Return LoOpt
Set FstLo = FstItm(A.ListObjects)
End Function






Function Wbn$(A As Workbook)
On Error GoTo X
Wbn = A.FullName
Exit Function
X: Wbn = "WbnErr:[" & Err.Description & "]"
End Function
Function LasWb() As Workbook
Set LasWb = LasWbz(Xls)
End Function
Function LasWbz(A As Excel.Application) As Workbook
Set LasWbz = A.Workbooks(A.Workbooks.Count)
End Function

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
Function EntColzPt(A As PivotTable, PivColNm) As Range
Set EntColzPt = RgR(PivCol(A, PivColNm).DataRange, 1).EntireColumn
End Function
Function PivColEnt(Pt As PivotTable, ColNm) As Range
Set PivColEnt = PivCol(Pt, ColNm).EntireColumn
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

Function WbzPt(A As PivotTable) As Workbook
Set WbzPt = WbzWs(WszPt(A))
End Function


Function WszPt(A As PivotTable) As Worksheet
Set WszPt = A.Parent
End Function
Function SampPt() As PivotTable
Set SampPt = PtzRg(SampRg)
End Function
Function SampRg() As Range
Set SampRg = ShwRg(PutSq(SampSq, NewA1))
End Function

Function FbtStrzQt$(A As QueryTable)
If IsNothing(A) Then Exit Function
Dim Ty As XlCmdType, Tbl$, CnStr$
With A
    Ty = .CommandType
    If Ty <> xlCmdTable Then Exit Function
    Tbl = .CommandText
    CnStr = .Connection
End With
FbtStrzQt = FmtQQ("[?].[?]", DtaSrczScvl(CnStr), Tbl)
End Function

Function A2zWs(A As Worksheet) As Range
Set A2zWs = A.Range("A2")
End Function

Function CvWs(A) As Worksheet
Set CvWs = A
End Function

Function CnozBefFstHid%(Ws As Worksheet)
Dim J%, O%
For J% = 1 To MaxWsCol
    If WsC(Ws, J).Hidden Then CnozBefFstHid = J - 1: Exit Function
Next
Stop
End Function



Function TxtCnzWc(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set TxtCnzWc = A.TextConnection
End Function

Property Get SampLoVis() As ListObject
Set SampLoVis = ShwLo(SampLo)
End Property

Property Get SampLo() As ListObject
Set SampLo = LozRg(RgzSq(SampSqWithHdr, NewA1), "Sample")
End Property


Property Get SampLofTp() As String()
Dim O$()
PushI O, "Lo  Nm     *Nm"
PushI O, "Lo  Fld    *Fld.."
PushI O, "Align Left   *Fld.."
PushI O, "Align Right  *Fld.."
PushI O, "Align Center *Fld.."
PushI O, "Bdr Left   *Fld.."
PushI O, "Bdr Right  *Fld.."
PushI O, "Bdr Col    *Fld.."
PushI O, "Tot Sum    *Fld.."
PushI O, "Tot Avg    *Fld.."
PushI O, "Tot Cnt    *Fld.."
PushI O, "Fmt *Fmt   *Fld.."
PushI O, "Wdt *Wdt   *Fld.."
PushI O, "Lvl *Lvl   *Fld.."
PushI O, "Cor *Cor   *Fld.."
PushI O, "Fml *Fld   *Formula"
PushI O, "Bet *Fld   *Fld1 *Fld2"
PushI O, "Tit *Fld   *Tit"
PushI O, "Lbl *Fld   *Lbl"
SampLofTp = O
End Property


Property Get SampDr_AToJ() As Variant()
Const Nc% = 10
Dim J%
For J = 0 To Nc - 1
    PushI SampDr_AToJ, Chr(Asc("A") + J)
Next
End Property

Property Get SampSq1() As Variant()
Dim O(), R&, C&
Const NR& = 1000
Const Nc& = 100
ReDim O(1 To NR, 1 To Nc)
For R = 1 To NR
For C = 1 To Nc
    O(R, C) = R + C
Next
Next
SampSq1 = O
End Property
Function VyrzAt(At As Range) As Variant()
VyrzAt = VyrzSqr(SqzRg(HRgWiVal(A1(At))))
End Function
Function VyczAt(At As Range) As Variant()
VyczAt = VyczSqc(SqzRg(VRgWiVal(A1(At))))
End Function


Function IsCellInRg(A As Range, Rg As Range) As Boolean
Dim R&, C%, R1&, R2&, C1%, C2%
R = A.Row
R1 = Rg.Row
If R < R1 Then Exit Function
R2 = R1 + Rg.Rows.Count
If R > R2 Then Exit Function
C = A.Column
C1 = Rg.Column
If C < C1 Then Exit Function
C2 = C1 + Rg.Columns.Count
If C > C2 Then Exit Function
IsCellInRg = True
End Function

Function IsCellInRgAp(Cell As Range, ParamArray RgAp()) As Boolean
Dim Av(): Av = RgAp
'IsCellInRgAp = IsCellInRgAv(A, Av)
End Function

Function IsCellInRgAv(A As Range, RgAv()) As Boolean
Dim V
For Each V In RgAv
    If IsCellInRg(A, CvRg(V)) Then IsCellInRgAv = True: Exit Function
Next
End Function


Function VbarRgAt(At As Range, Optional AtLeastOneCell As Boolean) As Range
If IsEmpty(At.Value) Then
    If AtLeastOneCell Then
        Set VbarRgAt = A1zRg(At)
    End If
    Exit Function
End If
Dim R1&: R1 = At.Row
Dim R2&
    If IsEmpty(RgRC(At, 2, 1)) Then
        R2 = At.Row
    Else
        R2 = At.End(xlDown).Row
    End If
Dim C%: C = At.Column
Set VbarRgAt = WsCRR(WszRg(At), C, R1, R2)
End Function

Property Get SampSqWithHdr() As Variant()
SampSqWithHdr = InsSqr(SampSq, SampDr_AToJ)
End Property

Property Get SampWs() As Worksheet
Dim O As Worksheet
Set O = NewWs
LozDrs SampDrs, WsRC(O, 2, 2)
Set SampWs = O
ShwWs O
End Property

Function FstWsn$(Fx)
FstWsn = FstItm(Wny(Fx))
End Function

Function OleCnStrzFx$(Fx)
OleCnStrzFx = "OLEDb;" & CnStrzFxAdo(Fx)
End Function

Function HasFx(Fx) As Boolean
Dim Wb As Workbook
For Each Wb In Xls.Workbooks
    If Wb.FullName = Fx Then HasFx = True: Exit Function
Next
End Function

Private Sub Z_RgzMoreBelow()
Dim R As Range, Act As Range, Ws As Worksheet
Set Ws = NewWs
Set R = Ws.Range("A3:B5")
Set Act = RgzMoreTop(R, 1)
Debug.Print Act.Address
Stop
Debug.Print RgRR(R, 1, 2).Address
Stop
End Sub

Function CvRg(A) As Range
Set CvRg = A
End Function
Function HRgWiVal(At As Range) As Range
Dim A1 As Range: Set A1 = A1zRg(At)
If IsEmpty(A1.Value) Then Set HRgWiVal = A1: Exit Function
Dim C2&: C2 = A1.End(xlRight).Column - A1.Column + 1
Set HRgWiVal = RgCRR(A1, 1, 1, C2)
End Function
Function VRgWiVal(At As Range) As Range
Dim A1 As Range: Set A1 = A1zRg(At)
If IsEmpty(A1.Value) Then Set VRgWiVal = A1: Exit Function
Dim R2&: R2 = A1.End(xlDown).Row - A1.Row + 1
Set VRgWiVal = RgCRR(A1, 1, 1, R2)
End Function
Function A1(Ws As Worksheet) As Range
Set A1 = WsRC(Ws, 1, 1)
End Function
Function A1zRg(A As Range) As Range
Set A1zRg = RgRC(A, 1, 1)
End Function
Function IsA1(A As Range) As Boolean
If A.Row <> 1 Then Exit Function
If A.Column <> 1 Then Exit Function
IsA1 = True
End Function
Function WsRgAdr$(A As Range)
WsRgAdr = "'" & WszRg(A).Name & "'!" & A.Address
End Function

Function RRRCCzRg(A As Range) As RRCC
With RRRCCzRg
.R1 = A.Row
.R2 = .R1 + A.Rows.Count - 1
.C1 = A.Column
.C2 = .C1 + A.Columns.Count - 1
End With
End Function

Function RgCEnt(A As Range, C) As Range
Set RgCEnt = RgC(A, C).EntireColumn
End Function
Function NxtCellBelow(A As Range) As Range
Dim O As Range: Set O = RgRC(A, 2, 1)
If IsEmpty(O.Value) Then Exit Function
Set NxtCellBelow = O
End Function
Function RgC(A As Range, C) As Range
Set RgC = RgCC(A, C, C)
End Function

Function RgCC(A As Range, C1, C2) As Range
Set RgCC = RgRCRC(A, 1, C1, NRowzRg(A), C2)
End Function

Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function

Function EntRgC(A As Range, C) As Range
Set EntRgC = RgC(A, C).EntireColumn
End Function

Function EntRgRR(A As Range, R1, R2) As Range
Set EntRgRR = RgRR(A, R1, R2).EntireRow
End Function

Function FstColRg(A As Range) As Range
Set FstColRg = RgC(A, 1)
End Function

Function FstRowRg(A As Range) As Range
Set FstRowRg = RgR(A, 1)
End Function

Function IsHBarRg(A As Range) As Boolean
IsHBarRg = A.Rows.Count = 1
End Function

Function IsVbarRg(A As Range) As Boolean
IsVbarRg = A.Columns.Count = 1
End Function

Function LasVCell(At As Range) As Range
Set LasVCell = RgR(At, NRowzRg(At))
End Function


Function LasHCell(Cell As Range) As Range
Set LasHCell = RgC(Cell, NColRg(Cell))
End Function

Function NColRg%(A As Range)
NColRg = A.Columns.Count
End Function

Function RgzMoreBelow(A As Range, Optional N% = 1)
Set RgzMoreBelow = RgRR(A, 1, A.Rows.Count + N)
End Function

Function DicValToWbAdrzRg(Rg As Range) As Dictionary
Dim R As Range, K
Set DicValToWbAdrzRg = New Dictionary
For Each R In Rg
    K = R.Value
    If IsStr(K) Then
        DicValToWbAdrzRg.Add K, R.Address(External:=True)
    End If
Next
End Function

Function LasRow(Lo As ListObject) As Range
If Lo.ListRows.Count = 0 Then Thw CSub, "There is no LasRow in Lo", "Lon Wsn", Lo.Name, WszLo(Lo).Name
Set LasRow = Lo.ListRows(Lo.ListRows.Count).Range
End Function

Function LasRowCell(Lo As ListObject, C) As Range
Dim Ix%: Ix = Lo.ListColumns(C).Index
Set LasRowCell = RgRC(LasRow(Lo), 1, Lo.ListColumns(C).Index)
End Function

Function LoCC(Lo As ListObject, C1$, C2$) As Range
Dim A%, B%
With Lo
    A = .ListColumns(C1).Index
    B = .ListColumns(C2).Index
    Set LoCC = RgCC(.DataBodyRange, A, B)
End With
End Function

Function DicNmAdrzWsNmPfx(Ws As Worksheet, NmPfx$) As Dictionary
Dim N As Name
Set DicNmAdrzWsNmPfx = New Dictionary
For Each N In Ws.Names
    If HasPfx(N.Name, NmPfx) Then
        DicNmAdrzWsNmPfx.Add N.Name, N.RefersTo
    End If
Next
End Function
Function RgzMoreTop(A As Range, Optional N = 1)
Dim O As Range
Set O = RgRR(A, 1 - N, A.Rows.Count)
Set RgzMoreTop = O
End Function

Function NRowzRg&(A As Range)
NRowzRg = A.Rows.Count
End Function

Function RgR(A As Range, R)
Set RgR = RgRR(A, R, R)
End Function

Function CellBelow(Cell As Range, Optional N = 1) As Range
Set CellBelow = RgRC(Cell, 1 + N, 1)
End Function

Sub SwapValzRg(Cell1 As Range, Cell2 As Range)
Dim A: A = RgRC(Cell1, 1, 1).Value
RgRC(Cell1, 1, 1).Value = RgRC(Cell2, 1, 1).Value
RgRC(Cell2, 1, 1).Value = A
End Sub
Function CellAbove(Cell As Range, Optional Above = 1) As Range
Set CellAbove = RgRC(Cell, 1 - Above, 1)
End Function

Function CellRight(A As Range, Optional Right = 1) As Range
Set CellRight = RgRC(A, 1, 1 + Right)
End Function
Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Function

Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = WszRg(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function

Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, NColRg(A))
End Function

Function ResiRg(At As Range, Sq()) As Range
If Si(Sq) = 0 Then Set ResiRg = A1zRg(At): Exit Function
Set ResiRg = RgRCRC(At, 1, 1, NRowzSq(Sq), NColzSq(Sq))
End Function

Function SqzRg(A As Range) As Variant()
If A.Columns.Count = 1 Then
    If A.Rows.Count = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = A.Value
        SqzRg = O
        Exit Function
    End If
End If
SqzRg = A.Value
End Function

Function WbzRg(A As Range) As Workbook
Set WbzRg = WbzWs(WszRg(A))
End Function
Function WszRg(A As Range) As Worksheet
Set WszRg = A.Parent
End Function

Private Sub ZZ()
Z_RgzMoreBelow
MXls_Z_Rg:
End Sub
Function DrszWL(Ws As Worksheet, LoNm$) As Drs
DrszWL = DrszLo(Ws.ListObjects(LoNm))
End Function
Function StrColzLC(Lo As ListObject, C$) As String()
StrColzLC = StrColzSqc(SqzRg(Lo.DataBodyRange))
End Function
Function StrColzWsLC(Ws As Worksheet, LoNm$, C$) As String()
StrColzWsLC = StrColzLC(Ws.ListObjects(LoNm), C)
End Function
Function VbarAy(A As Range) As Variant()
Ass IsVbarRg(A)
'VbarAy = ColzSq(RgzSq(A), 1)
End Function

Function VbarIntAy(A As Range) As Integer()
'VbarIntAy = AyIntAy(VbarAy(A))
End Function

Function VbarSy(A As Range) As String()
VbarSy = SyzAy(VbarAy(A))
End Function

Function RgzBelowCell(Cell As Range)
Dim Ws As Worksheet: Set Ws = WszRg(Cell)
Set RgzBelowCell = WsCRR(Ws, Cell.Column, Cell.Row, LasRno(Ws))
End Function

Function DrszFxq(Fx, Q) As Drs
DrszFxq = DrszArs(CnzFx(Fx).Execute(Q))
End Function
Function TmpDbzFx(Fx) As Database
Set TmpDbzFx = TmpDbzFxWny(Fx, Wny(Fx))
End Function

Function TmpDbzFxWny(Fx, Wny$()) As Database
Dim O As Database
   Set O = TmpDb
Dim W
For Each W In Itr(Wny)
    LnkFx O, ">" & W, Fx, W
Next
Set TmpDbzFxWny = O
End Function

Function Wb(Fx) As Workbook
Set Wb = Xls.Workbooks(Fx)
End Function
Function WbzFx(Fx) As Workbook
Set WbzFx = Wb(Fx)
End Function

Function WszFxw(Fx, Optional Wsn$ = "Data") As Worksheet
Set WszFxw = WszWb(WbzFx(Fx), Wsn)
End Function

Function ArszFxwf(Fx, W$, F$) As AdoDb.Recordset
Set ArszFxwf = ArsCnq(CnzFx(Fx), SqlSel_F_T(F, W & "$"))
End Function

Function WsCdNyzFx(Fx) As String()
Dim Wb As Workbook
Set Wb = WbzFx(Fx)
WsCdNyzFx = WsCdNy(Wb)
Wb.Close False
End Function

Function DtzFxw(Fx, Optional Wsn0$) As Dt
Dim N$: N = DftWsn(Wsn0, Fx)
Dim Q$: Q = FmtQQ("Select * from [?$]", N)
DtzFxw = DtzDrs(DrszFxq(Fx, Q), N)
End Function

Function IntAyFxwf(Fx, W$, F$) As Integer()
IntAyFxwf = IntAyzArs(ArszFxwf(Fx, W, F))
End Function

Function SyzFxwf(Fx, W$, F$) As String()
SyzFxwf = SyzArs(ArszFxwf(Fx, W, F))
End Function

Private Sub ZZ_Wny()
Const Fx$ = "Users\user\Desktop\Invoices 2018-02.xlsx"
D Wny(Fx)
End Sub

Private Sub Z_FstWsn()
Dim Fx$
Fx = SampFxzKE24
Ept = "Sheet1"
GoSub T1
Exit Sub
T1:
    Act = FstWsn(Fx)
    C
    Return
End Sub




Private Sub Z_TmpDbzFx()
Dim Db As Database: Set Db = TmpDbzFx("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
DmpAy Tny(Db)
Db.Close
End Sub

Private Sub Z_Wny()
Dim Fx$
'GoTo ZZ
GoSub T1
Exit Sub
T1:
    Fx = SampFxzKE24
    Ept = SyzSS("Sheet1 Sheet21")
    GoSub Tst
    Return
Tst:
    Act = Wny(Fx)
    C
    Return
ZZ:
    DmpAy Wny(SampFxzKE24)
End Sub

Function ChkFxww(Fx, Wsnn$, Optional FxKd$ = "Excel file") As String()
Dim W$, I
'If Not HasFfn(Fx) Then ChkFxww = MsgzMisFfn(Fx, FxKd): Exit Function
For Each I In Ny(Wsnn)
    W = I
    PushIAy ChkFxww, ChkWs(Fx, W, FxKd)
Next
End Function

Function ChkWs(Fx, Wsn, FxKd$) As String()
If HasFxw(Fx, Wsn) Then Exit Function
Dim M$
M = FmtQQ("? does not have expected worksheet", FxKd)
ChkWs = LyzFunMsgNap(CSub, M, "Folder File Expected-Worksheet Worksheets-in-file", Pth(Fx), Fn(Fx), Wsn, Wny(Fx))
End Function

Function ChkFxw(Fx, Wsn, Optional FxKd$ = "Excel file") As String()
ChkFxw = ChkHasFfn(Fx, FxKd): If Si(ChkFxw) > 0 Then Exit Function
ChkFxw = ChkWs(Fx, Wsn, FxKd)
End Function
Function ChkLnkWs(A As Database, T, Fx, Wsn, Optional FxKd$ = "Excel file") As String()
Const CSub$ = CMod & "ChkLnkWs"
Dim O$()
    O = ChkFxw(Fx, Wsn, FxKd)
    If Si(O) > 0 Then
        ChkLnkWs = O
        Exit Function
    End If
On Error GoTo X
LnkFx A, T, Fx, Wsn
Exit Function
X: ChkLnkWs = _
    LyzMsgNap("Error in linking Xls file", "Er LnkFx LnkWs ToDb AsTbl", Err.Description, Fx, Wsn, Dbn(A), T)
End Function



Function WszAyab(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As Worksheet
Set WszAyab = WszDrs(DrszAyab(A, B, N1, N2))
End Function
Function WszCd(WsCdn$) As Worksheet
Dim Ws As Worksheet
For Each Ws In CWb.Sheets
    If Ws.CodeName = WsCdn Then Set WszCd = Ws: Exit Function
Next
End Function
Function WszDic(Dic As Dictionary, Optional InclDicValOptTy As Boolean) As Worksheet
Set WszDic = WszDrs(DrszDic(Dic, InclDicValOptTy))
End Function

Function WszDt(A As Dt) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
LozDrs DrszDt(A), A1zWs(O)
WszDt = O
End Function

Function NyzFml(Fml$) As String()
NyzFml = NyzMacro(Fml)
End Function

Function WszLy(Ly$(), Optional Wsn$ = "Sheet1") As Worksheet
Set WszLy = WszRg(RgzAyV(Ly, A1zWs(NewWs(Wsn))))
End Function

Property Get MaxWsCol&()
Static C&, Y As Boolean
If Not Y Then
    Y = True
    C = IIf(Xls.Version = "16.0", 16384, 255)
End If
MaxWsCol = C
End Property

Property Get MaxWsRow&()
Static R&, Y As Boolean
If Not Y Then
    Y = True
    R = IIf(Xls.Version = "16.0", 1048576, 65535)
End If
MaxWsRow = R
End Property

Function SqHzN(N%) As Variant()
Dim O()
ReDim O(1 To 1, 1 To N)
SqHzN = O
End Function

Function SqVzN(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
SqVzN = O
End Function

Function N_ZerFill$(N, NDig&)
N_ZerFill = Format(N, String(NDig, "0"))
End Function

Function WszS1S2s(A As S1S2s, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set WszS1S2s = WszSq(SqzS1S2s(A, Nm1, Nm2))
End Function

Private Sub Z_AyabWs()
GoTo ZZ
Dim A, B
ZZ:
    A = SyzSS("A B C D E")
    B = SyzSS("1 2 3 4 5")
    ShwWs WszAyab(A, B)
End Sub

Private Sub Z_WbFbOupTbl()
GoTo ZZ
ZZ:
    Dim W As Workbook
    'Set W = WbFbOupTbl(WFb)
    'ShwWb W
    Stop
    'W.Close False
    Set W = Nothing
End Sub

Function LasCnozRg%(R As Range)
LasCnozRg = R.Column + R.Columns.Count - 1
End Function

Function LasRnozRg&(R As Range)
LasRnozRg = R.Row + R.Rows.Count - 1
End Function

