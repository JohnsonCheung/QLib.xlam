Attribute VB_Name = "QXls_Xls_XlsOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_Wb."
Private Const Asm$ = "QXls"
Sub BrwFx(Fx)
If HasFfn(Fx) Then Debug.Print "No Fx:" & Fx
ShwWb WbzFx(Fx)
End Sub

Sub CrtFx(Fx)
WbSavAs(NewWb, Fx).Close
End Sub

Function EnsFx$(Fx)
If Not HasFfn(Fx) Then CrtFx Fx
EnsFx = Fx
End Function

Function OpnFx(Fx) As Workbook
ThwIf_FfnNotExist Fx, CSub
Set OpnFx = Xls.Workbooks.Open(Fx)
End Function
Sub ClrWsNm(Ws As Worksheet)
Dim N
For Each N In Itn(Ws.Names)
    Ws.Names(N).Delete
Next
End Sub
Sub EnsHypLnkzFollowNm(Rg As Range, NmPfx$) 'Do 3 things: _
1. RmvExcess HypLnk _
2. Chg Val to Fml _
3. Crt HypLnk.
Dim P$: P = NmPfx & "_"
'1. RmvExcess HypLnk
    Dim HL() As Excel.Hyperlink
    Dim H As Hyperlink
    For Each H In Rg.Hyperlinks
        If Not HasPfx(H.Address, NmPfx) Then PushObj HL, H
    Next
    DoOyMth HL, "Delete"    '<== Rmv
'2. Chg Val to Fml
    Dim R As Range, Ny$(), V, F$
    Ny = RmvPfxzAy(AywPfx(Itn(WbzRg(Rg).Names), P), P)
    For Each R In Rg
        V = R.Value
        If IsStr(V) Then
            If HasEle(Ny, V) Then
                F = "=" & P & V
                If R.Formula <> F Then
                    R.Formula = F '<== Changed
                End If
            End If
        End If
    Next
'3. Crt HL
    For Each R In Rg
        With R.Hyperlinks
        V = R.Value
        Select Case True
        Case Not IsStr(V)
        Case Not HasEle(Ny, V)
        Case .Count > 0
        Case Else
            .Add Anchor:=R, Address:="", SubAddress:=P & R.Value
        End Select
        End With
    Next
   
End Sub
Sub EnsWbNmzLcPfx(Ws As Worksheet, LoNm$, Col$, NmPfx$)
Dim P$:                               P = NmPfx & "_"
Dim Rg As Range:                 Set Rg = Ws.ListObjects(LoNm).ListColumns(Col).DataBodyRange
Dim OldNm As New Dictionary:  Set OldNm = DicNmAdrzWsNmPfx(Ws, P)
Dim NewNm As New Dictionary:  Set NewNm = AddPfxToKey(P, DicValToWbAdrzRg(Rg))
Dim Add As Dictionary:          Set Add = MinusDic(NewNm, OldNm)
Dim Rmv$():                         Rmv = SyzDicKey(MinusDic(OldNm, NewNm))
Dim Upd As Dictionary:          Set Upd = DicAzDifVal(NewNm, OldNm)
'Add
    Dim Nm
    For Each Nm In Add.Keys
        WbzWs(Ws).Names.Add Nm, "=" & Add(Nm)
    Next
'Upd
    For Each Nm In Upd.Keys
        Ws.Names(Nm).RefersTo = Upd(Nm)
    Next
'Rmv
    Dim I
    For Each I In Itr(Rmv)
        Ws.Names(I).Delete
    Next
End Sub
Sub RmvWsIf(Fx, Wsn$)
If HasFxw(Fx, Wsn) Then
   Dim B As Workbook: Set B = WbzFx(Fx)
   WszWb(B, Wsn).Delete
   SavWb B
   ClsWbNoSav B
End If
End Sub

Function LozAyH(Ay, Wb As Workbook, Optional Wsn$, Optional Lon$) As ListObject
Set LozAyH = LozRg(RgzSq(Sqh(Ay), A1zWb(Wb, Wsn)), Lon)
End Function

Private Sub Z_SetWsCdNm()
Dim A As Worksheet: Set A = NewWs
SetWsCdNm A, "XX"
ShwWs A
Stop
End Sub
Sub MgeBottomCell(VBar As Range)
Ass IsVbarRg(VBar)
Dim R2: R2 = VBar.Rows.Count
Dim R1
    Dim Fnd As Boolean
    For R1 = R2 To 1 Step -1
        If Not IsEmpty(RgRC(VBar, R1, 1)) Then Fnd = True: GoTo Nxt
    Next
Nxt:
    If Not Fnd Then Stop
If R2 = R1 Then Exit Sub
Dim R As Range: Set R = RgCRR(VBar, 1, R1, R2)
R.Merge
R.VerticalAlignment = XlVAlign.xlVAlignTop
End Sub

Sub SetWsCdNm(A As Worksheet, CdNm$)
CmpzWs(A).Name = CdNm
End Sub

Sub SetWsCdNmAndLoNm(A As Worksheet, Nm$)
CmpzWs(A).Name = Nm
SetLoNm FstLo(A), Nm
End Sub
Function RgzDbtzByWc(Db As Database, T, At As Range) As Range

End Function

Sub BdrRgAy(A() As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
Dim I
For Each I In Itr(A)
    BdrRg CvRg(I), Ix, Wgt
Next
End Sub

Sub BdrRg(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub

Sub BdrRgAround(A As Range)
BdrRgLeft A
BdrRgRight A
BdrRgTop A
BdrRgBottom A
End Sub

Sub BdrRgBottom(A As Range)
BdrRg A, xlEdgeBottom
BdrRg A, xlEdgeTop
End Sub

Sub BdrRgInner(A As Range)
BdrRg A, xlInsideHorizontal
BdrRg A, xlInsideVertical
End Sub

Sub BdrRgInside(A As Range)
BdrRgInner A
End Sub
Sub BdrRgAlign(A As Range, H As XlHAlign)
Select Case H
Case XlHAlign.xlHAlignLeft: BdrRgLeft A
Case XlHAlign.xlHAlignRight: BdrRgRight A
End Select
End Sub
Sub BdrRgLeft(A As Range)
BdrRg A, xlEdgeLeft
If A.Column > 1 Then
    BdrRg RgC(A, 0), xlEdgeRight
End If
End Sub

Sub BdrRgRight(A As Range)
BdrRg A, xlEdgeRight
If A.Column < MaxWsCol Then
    BdrRg RgC(A, A.Columns.Count + 1), xlEdgeLeft
End If
End Sub

Sub BdrRgTop(A As Range)
BdrRg A, xlEdgeTop
If A.Row > 1 Then
    BdrRg RgR(A, 0), xlEdgeBottom
End If
End Sub


Sub MgeRg(A As Range)
A.MergeCells = True
A.HorizontalAlignment = XlHAlign.xlHAlignCenter
A.VerticalAlignment = XlVAlign.xlVAlignCenter
End Sub


Function LozSq(Sq(), At As Range, Optional Lon$) As ListObject
Set LozSq = LozRg(RgzSq(Sq(), At), Lon)
End Function


Function LozRg(Rg As Range, Optional Lon$) As ListObject
Dim O As ListObject: Set O = WszRg(Rg).ListObjects.Add(xlSrcRange, Rg, , xlYes)
BdrRgAround Rg
Rg.EntireColumn.AutoFit
SetLoNm O, Lon
Set LozRg = O
End Function

Function RgzDbt(Db As Database, T, At As Range) As Range
Set RgzDbt = RgzSq(SqzT(Db, T), At)
End Function

Sub AddWszT1(A As Workbook, Db As Database, T, Optional Wsn0$, Optional AddgWay As EmAddgWay)
If AddgWay = EiSqWay Then AddWszT A, Db, T, Wsn0, AddgWay: Exit Sub
Dim Wsn$: Wsn = DftStr(Wsn0, T)
Dim Sq(): Sq = SqzT(Db, T)
Dim A1 As Range: Set A1 = A1zWs(AddWs(A, Wsn))
End Sub

Sub PutTbl(A As Database, T, At As Range, Optional AddgWay As EmAddgWay)
Select Case AddgWay
Case EmAddgWay.EiSqWay: PutSq SqzT(A, T), At
Case EmAddgWay.EiWcWay: AddLo At, A.Name, T
Case Else: Thw CSub, "Invalid AddgWay"
End Select
End Sub

Sub AddWszTny(A As Workbook, Db As Database, Tny$(), Optional AddgWay As EmAddgWay)
Dim T$, I
For Each I In Tny
    T = I
    AddWszT A, Db, T, , AddgWay
Next
End Sub

Function WszWbDt(A As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = AddWs(A, Dt.DtNm)
LozDrs DrszDt(Dt), A1(O)
Set WszWbDt = O
End Function

Function AddWc(ToWb As Workbook, FmFb, T) As WorkbookConnection
Set AddWc = ToWb.Connections.Add2(T, T, CnStrzFbzForWc(FmFb), T, XlCmdType.xlCmdTable)
End Function

Sub ThwWbMisOupNy(A As Workbook, OupNy$())
Dim O$(), N$, B$(), Wny$()
Wny = WsCdNy(A)
O = MinusAy(AddPfxzAy(OupNy, "WsO"), Wny)
If Si(O) > 0 Then
    N = "OupNy":  B = OupNy:  GoSub Dmp
    N = "WbCdNy": B = Wny: GoSub Dmp
    N = "Mssing": B = O:      GoSub Dmp
    Stop
    Exit Sub
End If
Exit Sub
Dmp:
Debug.Print UnderLin(N)
Debug.Print N
Debug.Print UnderLin(N)
DmpAy B
Return
End Sub

Sub ClsWbNoSav(A As Workbook)
A.Close False
End Sub

Sub DltWc(A As Workbook)
Dim Wc As Excel.WorkbookConnection
For Each Wc In A.Connections
    Wc.Delete
Next
End Sub


Function WbMax(A As Workbook) As Workbook
A.Application.WindowState = xlMaximized
Set WbMax = A
End Function

Function NewA1Wb(A As Workbook, Optional Wsn$) As Range
'Set NewA1Wb = A1zWs(AddWs(A, Wsn))
End Function

Sub WbQuit(A As Workbook)
QuitXls A.Application
End Sub

Function SavWb(A As Workbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.Save
A.Application.DisplayAlerts = Y
Set SavWb = A
End Function

Function WbSavAs(A As Workbook, Fx, Optional Fmt As XlFileFormat = xlOpenXMLWorkbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.SaveAs Fx, Fmt
A.Application.DisplayAlerts = Y
Set WbSavAs = A
End Function

Sub SetWcFcsv(A As Workbook, Fcsv$)
'Set first Wb TextConnection to Fcsv if any
Dim T As TextConnection: Set T = TxtWc(A)
Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub

Function HasWs(A As Workbook, WsIx) As Boolean
If IsNumeric(WsIx) Then
    HasWs = IsBet(WsIx, 1, A.Sheets.Count)
    Exit Function
End If
HasWs = HasItn(A.Sheets, WsIx)
End Function

Private Sub ZZ_WbWcsy()
'D WcStrAyWbOLE(WbzFx(TpFx))
End Sub

Private Sub ZZ_LozAyH()
'D NyOy(LozAyH(TpWb))
End Sub

Private Sub Z_TxtWcCnt()
Dim O As Workbook: 'Set O = WbzFx(Vbe_MthFx)
Ass TxtWcCnt(O) = 1
O.Close
End Sub

Private Sub Z_SetWcFcsv()
Dim Wb As Workbook
'Set Wb = WbzFx(Vbe_MthFx)
Debug.Print TxtWcStr(Wb)
SetWcFcsv Wb, "C:\ABC.CSV"
Ass TxtWcStr(Wb) = "TEXT;C:\ABC.CSV"
Wb.Close False
Stop
End Sub

Private Sub ZZ()
Dim A
Dim B As WorkbookConnection
Dim C As Workbook
Dim D$
Dim E As Database
Dim F As Boolean
Dim G As Dt
Dim H$()
Dim I()
Dim XX
CvWb A
TxtCnzWc B
FstWs C
FxzWb C
LasWs C
MainWs C
TxtWcCnt C
TxtWcStr C
Wsny C
WszCdNm C, D
WszCdNm C, D
AddWs C, D, F, F, D, D
ThwWbMisOupNy C, H
ClsWbNoSav C
DltWc C
WbSavAs C, A
ShwWb C
XX = CWb()
End Sub


Function AddWs(A As Workbook, Optional Wsn$, Optional AtBeg As Boolean, Optional AtEnd As Boolean, Optional BefWsn$, Optional AftWsn$) As Worksheet
Dim O As Worksheet
DltWsIf A, Wsn
Select Case True
Case AtBeg:        Set O = A.Sheets.Add(FstWs(A))
Case AtEnd:        Set O = A.Sheets.Add(LasWs(A))
Case BefWsn <> "": Set O = A.Sheets.Add(A.Sheets(BefWsn))
Case AftWsn <> "": Set O = A.Sheets.Add(, A.Sheets(AftWsn))
Case Else:         Set O = A.Sheets.Add
End Select
SetWsn O, Wsn
Set AddWs = O
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



Sub DltWsIf(A As Workbook, WsIx)
If HasWs(A, WsIx) Then DltWs A, WsIx
End Sub








Sub SavAszAndCls(Wb As Workbook, Fx)
Wb.SaveAs Fx
Wb.Close
End Sub


Function WbnzWs$(A As Worksheet)
WbnzWs = WbzWs(A).FullName
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


Function PtCpyToLo(A As PivotTable, At As Range) As ListObject
Dim R1, R2, C1, C2, NC, NR
    R1 = A.RowRange.Row
    C1 = A.RowRange.Column
    R2 = LasRowRg(A.DataBodyRange)
    C2 = LasColRg(A.DataBodyRange)
    NC = C2 - C1 + 1
    NR = R2 - C1 + 1
WsRCRC(WszPt(A), R1, C1, R2, C2).Copy
At.PasteSpecial xlPasteValues

Set PtCpyToLo = LozRg(RgRCRC(At, 1, 1, NR, NC))
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


Sub ThwHasWbzWs(Wb As Workbook, Wsn$, Fun$)
If HasWs(Wb, Wsn) Then
    Thw Fun, "Wb should have not have Ws", "Wb Ws", Wb.FullName, Wsn
End If
End Sub

Sub SetPtWdt(A As PivotTable, Colss$, ColWdt As Byte)
If ColWdt <= 1 Then Stop
Dim C
For Each C In Itr(SyzSS(Colss))
    EntColzPt(A, C).ColumnWidth = ColWdt
Next
End Sub

Sub SetPtOutLin(A As PivotTable, Colss$, Optional Lvl As Byte = 2)
If Lvl <= 1 Then Stop
Dim F, C As VBComponent
For Each C In Itr(SyzSS(Colss))
    EntColzPt(A, F).OutlineLevel = Lvl
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

Function PutSq(Sq(), At As Range) As Range
Dim O As Range
Set O = ResiRg(At, Sq)
LozRg O
O.Value = Sq
Set PutSq = O
End Function

Function NewA1(Optional Wsn$) As Range
Set NewA1 = A1zWs(NewWs(Wsn))
End Function

Function NewWb(Optional Wsn$) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Add
Set NewWb = WbzWs(SetWsn(FstWs(O), Wsn))
End Function

Function NewWs(Optional Wsn$) As Worksheet
Set NewWs = SetWsn(FstWs(NewWb), Wsn)
End Function

Function NewXls() As Excel.Application
Set NewXls = CreateObject("Excel.Application") ' Don't use New Excel.Application
End Function


Sub QuitXls(A As Excel.Application)
Stamp "QuitXls: Start"
Stamp "QuitXls: ClsAllWb":    ClsAllWb A
Stamp "QuitXls: Quit":        A.Quit
Stamp "QuitXls: Set nothing": Set A = Nothing
Stamp "QuitXls: Done"
End Sub
Sub ClsAllWb(A As Excel.Application)
Dim W As Workbook
For Each W In A.Workbooks
    W.Close False
Next
End Sub
Private Sub ClsWc(A As WorkbookConnection)
If IsNothing(A.OLEDBConnection) Then Exit Sub
CvCn(A.ODBCConnection.Connection).Close
End Sub

Private Sub ClsWczWb(Wb As Workbook)
Dim Wc As WorkbookConnection
For Each Wc In Wb.Connections
    ClsWc Wc
Next
End Sub

Private Sub SetWczFb(A As WorkbookConnection, ToUseFb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
Dim Cn$
#Const A = 2
#If A = 1 Then
    Dim S$
    S = A.OLEDBConnection.Connection
    Cn = RplBet(S, ToUseFb, "Data Source=", ";")
#End If
#If A = 2 Then
    Cn = CnStrzFbzAsAdoOle(ToUseFb)
#End If
A.OLEDBConnection.Connection = Cn
End Sub

Private Sub RfhWc(A As WorkbookConnection, ToUseFb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
SetWczFb A, ToUseFb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub

Private Sub RfhPc(A As PivotCache)
A.MissingItemsLimit = xlMissingItemsNone
A.Refresh
End Sub

Sub RfhFx(Fx, Fb)
RfhWb(WbzFx(Fx), Fb).Close SaveChanges:=True
End Sub

Private Sub RfhWs(A As Worksheet)
Dim Q As QueryTable: For Each Q In A.QueryTables: Q.BackgroundQuery = False: Q.Refresh: Next
Dim P As PivotTable: For Each P In A.PivotTables: P.Update: Next
Dim L As ListObject: For Each L In A.ListObjects: L.Refresh: Next
End Sub

Function RfhWb(Wb As Workbook, Fb) As Workbook
RplLozFb Wb, Fb
Dim C As WorkbookConnection
Dim P As PivotCache, W As Worksheet
'For Each C In Wb.Connections: RfhWc C, Fb:                                          Next
For Each P In Wb.PivotCaches: P.MissingItemsLimit = xlMissingItemsNone: P.Refresh:  Next
For Each W In Wb.Sheets:      RfhWs W:                                              Next
StdFmtLozWb Wb
ClsWczWb Wb
DltWc Wb
Set RfhWb = Wb
End Function

Private Sub RplLozFb(Wb As Workbook, Fb)
Dim Ws As Worksheet, D As Database
Set D = Db(Fb)
For Each Ws In Wb.Sheets
    RplLozWs Ws, D
Next
D.Close
End Sub

Private Sub RplLozWs(Ws As Worksheet, D As Database)
Dim Lo As ListObject
For Each Lo In Ws.ListObjects
    RplLozT Lo, D, "@" & Mid(Lo.Name, 3)
Next
End Sub

Private Sub RplLozT(A As ListObject, Db As Database, T)
Dim Fny1$(): Fny1 = Fny(Db, T)
Dim Fny2$(): Fny2 = FnyzLo(A)
If Not IsSamAy(Fny1, Fny2) Then
    Thw CSub, "LoFny and TblFny are not same", "LoFny TblNm TblFny Db", Fny2, T, Fny1, Dbn(A)
End If
Dim Sq()
    Dim R As DAO.Recordset
    Set R = Rs(A, SqlSel_Fny_T(Fny2, T))
    Sq = AddSngQuotezSq(SqzRs(R))
MinxLo A
RgzSq Sq, A.DataBodyRange
End Sub


Sub PutWc(Wc As WorkbookConnection, At As Range)
Dim Lo As ListObject
Set Lo = WszRg(At).ListObjects.Add(SourceType:=0, Source:=Wc.OLEDBConnection.Connection, Destination:=At)
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = Wc.Name
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .ListObject.DisplayName = Lon(Wc.Name)
    .Refresh BackgroundQuery:=False
End With
End Sub

Sub AddWczTT(ToWb As Workbook, FmFb, TT$)
Dim T$, I
For Each I In Ny(TT)
    T = I
    AddWc ToWb, FmFb, T
Next
End Sub

Private Sub Z_CrtFxzOupTbl()
Dim Fx$: Fx = TmpFx
CrtFxzOupTbl Fx, SampFbzDutyDta
OpnFx Fx
End Sub
Sub CrtFxzOupTbl(Fx, Fb, Optional AddgWay As EmAddgWay)
SavAszAndCls NewWbzOupTbl(Fb, AddgWay), Fx
End Sub



Function ShwWb(A As Workbook) As Workbook
ShwXls A.Application
Set ShwWb = A
End Function

Function ShwXls(A As Excel.Application) As Excel.Application
If Not A.Visible Then A.Visible = True
Set ShwXls = A
End Function

Function ShwRg(A As Range) As Range
ShwXls A.Application
Set ShwRg = A
End Function

Function ShwLo(A As ListObject) As ListObject
ShwXls A.Application
Set ShwLo = A
End Function

Function ShwWs(A As Worksheet) As Worksheet
ShwXls A.Application
Set ShwWs = A
End Function


Function WbzDs(A As Ds) As Workbook
Dim O As Workbook
Set O = NewWb
With FstWs(O)
   .Name = "Ds"
   .Range("A1").Value = A.DsNm
End With
Dim J%, Ay() As Dt
For J = 0 To A.N - 1
    'WszWbDt O, Ay(J)
Next
Set WbzDs = O
End Function
Sub PutSeqDown(A As Range, FmNum&, ToNum&)
'AyRgV LngSeq(FmNum, ToNum), A
End Sub



Sub DltSheet1(Wb As Workbook)
DltWs Wb, "Sheet1"
End Sub

Sub DltWs(Wb As Workbook, WsIx)
Wb.Application.DisplayAlerts = False
If Wb.Sheets.Count = 1 Then Exit Sub
If HasWs(Wb, WsIx) Then WszWb(Wb, WsIx).Delete
End Sub

Sub ClrDown(A As Range)
VbarRgAt(A, AtLeastOneCell:=True).Clear
End Sub


Sub MgeCellAbove(Cell As Range)
'If Not IsEmpty(A.Value) Then Exit Sub
If Cell.MergeCells Then Exit Sub
If Cell.Row = 1 Then Exit Sub
If RgRC(Cell, 0, 1).MergeCells Then Exit Sub
MgeRg RgCRR(Cell, 1, 0, 1)
End Sub


Sub FillSeqH(HBar As Range)
Dim Sq()
Sq = SqVzN(HBar.Rows.Count)
ResiRg(HBar, Sq).Value = Sq
End Sub
Sub ClrCellBelow(Cell As Range)
RgzBelowCell(Cell).Clear
End Sub

Sub FillSeqV(VBar As Range)
Dim Sq()
Sq = SqHzN(VBar.Rows.Count)
ResiRg(VBar, Sq).Value = Sq
End Sub

Sub FillWsny(At As Range)
RgzAyV Wsny(WbzRg(At)), At
End Sub

