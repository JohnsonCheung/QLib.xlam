Attribute VB_Name = "MxXlsOp"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsOp."
Enum EmWsPos
    EiEnd
    EiBeg
    EiRfWs
End Enum

Sub BrwFx(Fx)
If HasFfn(Fx) Then Debug.Print "No Fx:" & Fx
ShwWb WbzFx(Fx)
End Sub

Sub CrtFx(Fx)
WbSavAs(NewWb, Fx).Close
End Sub

Function EnsFx$(Fx)
'Ret : crt an emp @Fx if not exist @@
If NoFfn(Fx) Then CrtFx Fx
EnsFx = Fx
End Function

Sub InsRow(OCell As Range, Optional N = 1)
'Fm OCell : due to xls ins row will the rg used to ins row, so @OCell will be rest to A1-of-original-@OCell.
'Ret      : ins @N emp rows below the fst row of @OCell
Dim Ws As Worksheet: Set Ws = WszRg(OCell)
Dim R&:                   R = OCell.Row
Dim C%:                   C = OCell.Column
Dim OIns:                     RgCRR(OCell, 1, 1, N).EntireRow.Insert ' @OCell will be changed after insert  ! <==
                  Set OCell = WsRC(Ws, R, C)                         ' reset @OCell to A1-of-orignal-@OCell ! <==
End Sub

Sub RmvRowEmp(ORg As Range)
'Fm ORg : ! the rect may have emp row.  Aft rmv, @ORg will chg to pointing to the remaining rect.
'Ret    : ! ##RmvRow-Wht:Emp-Fm:ORg Rmv all emp row ept keeping one emp row.  Rows rmv is ent row.
'         ! All non emp row will keep no chg and may shf up.
'         ! set the @ORg pointing to
Dim Ws As Worksheet: Set Ws = WszRg(ORg)     '                        ! <--
Dim R&:                   R = ORg.Row        ' A1-Rno-of-@ORg         ! <--
Dim C%:                   C = ORg.Column     ' A1-Cno-of-@ORg         ! <--
Dim NR%:                 NR = ORg.Rows.Count
Dim C2%:                 C2 = C2zRg(ORg)     ' N-Columns-of-@ORg      ! <--
Dim Sq():                Sq = SqzRgNo(ORg)
Dim RRmv&():           RRmv = RxyzEmpRow(Sq) ' Rny to be rmv '        ! <--
Dim NRLeft&:         NRLeft = NR - Si(R)     ' R2-of-the-remaining-rg
Dim R2&:                 R2 = R + NRLeft - 1 '                        ! <--
'=======================================================================================================================
EntRows(Ws, RRmv).Remove
Set ORg = WsRCRC(Ws, R, C, R2, C2)
End Sub

Function EntRowsStr$(Rny&())
Dim O$()
Dim Rno: For Each Rno In Rny
    PushI O, Rno & ":" & Rno
Next
EntRowsStr = JnComma(O)
End Function

Function EntRows(Ws As Worksheet, Rny&()) As Range
Set EntRows = Ws.Range(EntRowsStr(Rny))
End Function

Sub RmvRow(OCell As Range, Optional N = 1)
'Fm OCell : due the xls rmv row, the range used to rmv row will be changed, so @OCell will be rest to A1-of-original-@OCell.
'Ret      : rmv the @N-rows fm the fst of @OCell.
Dim Bef$: Bef = RgA1(OCell).Address    ' Original A1-of-@OCell
RgCRR(OCell, 1, 1, N).EntireRow.Insert ' @OCell will be changed after insert
Set OCell = RgRC(OCell, -N + 1, 1)     ' reset @OCell to A1-of-orignal-@OCell
Dim Aft$: Aft = OCell.Address          ' Aft rmv row, the adr-of-@OCell
If Bef <> Aft Then Stop
End Sub

Function OpnFx(Fx) As Workbook
ThwIf_NoFfn Fx, CSub
Set OpnFx = Xls.Workbooks.Open(Fx)
End Function

Sub OpnFcsv(Fcsv)
Xls.Workbooks.OpenText Fcsv
End Sub

Sub ClrWsNm(Ws As Worksheet)
Dim N
For Each N In Itn(Ws.Names)
    Ws.Names(N).Delete
Next
End Sub

Private Sub XEnsHyp(Cell As Range, W$())
'Fm W : Ws ny
'Ret : Ens the @Cell is pointing of one of @W if the @cell val is in @W else rmv the hyp lnk @@
Dim V:  V = Cell.Value
Dim Wsn$: If IsStr(V) Then If HasEle(W, V) Then Wsn = V  ' The wsn pointed by the @Cell.Value, if the @Cell.Value is str and in @W
Dim A$                                                   ' The adr of A1 of %Wsn
    If Wsn <> "" Then A = WszWb(WbzRg(Cell), Wsn).Range("A1").Address(External:=True)
Dim L As Hyperlink: For Each L In Cell.Hyperlinks
    If L.SubAddress <> A Then L.Delete
Next
'Need to add if #A<>"" and @Cell hyp lnk cnt = 0
With Cell.Hyperlinks
    If A <> "" And .Count = 0 Then
        .Add Anchor:=Cell, Address:="", SubAddress:=A
    End If
End With
End Sub
Private Sub Z_EnsHyp()
Dim Rg As Range
GoSub Crt
GoSub T0
GoSub Rmv
Exit Sub
T0:
    Set Rg = CWs.Range("A1:A2")
    EnsHyp Rg
    'Stop
    Return
Crt:
    Dim Wb As Workbook: Set Wb = NewWb("A")
    AddWs Wb, "A 1"
    AddWs Wb, "A 2"
    WszWb(Wb, "A").Activate
    CWs.Range("A1").Value = "A 1"
    CWs.Range("A2").Value = "A 2"
    ShwWb Wb
    Return
Rmv:
    Wb.Close False
    Return
End Sub
Sub EnsHyp(Rg As Range)
'Ret: Ens the hyp lnk in @Rg is pointing ws A1 if the @Rg val is a ws name. @@
Dim W$(): W = Wsny(WbzRg(Rg))   ' Ws ny
Dim Cell As Range: For Each Cell In Rg
    XEnsHyp Cell, W 'Ens the @Cell is pointing of one of @W if the @cell val is in @W else rmv the hyp lnk
Next
End Sub

Sub RmvWsIf(Fx, Wsn$)
'Ret : Ret @Wsn in @Fx if exists @@
If HasFxw(Fx, Wsn) Then
   Dim B As Workbook: Set B = WbzFx(Fx)
   WszWb(B, Wsn).Delete
   SavWb B
   ClsWbNoSav B
End If
End Sub

Function LozAyH(Ay, Wb As Workbook, Optional Wsn$, Optional Lon$) As ListObject
Set LozAyH = LozRg(RgzSq(SqH(Ay), A1zWb(Wb, Wsn)), Lon)
End Function

Private Sub Z_SetWsCdNm()
Dim A As Worksheet: Set A = NewWs
SetWsCdNm A, "XX"
ShwWs A
Stop
End Sub
Sub MgeBottomCell(VBar As Range)
Ass IsRgVBar(VBar)
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
    BDr CvRg(I), Ix, Wgt
Next
End Sub

Function BdrLoAround(A As ListObject)
Dim R As Range
Set R = RgMoreTop(A.DataBodyRange)
If A.ShowTotals Then Set R = RgMoreBelow(R)
BdrAround R
End Function

Sub BdrEr(R As Range, Optional ColrNm$ = "Red")
R.BorderAround xlContinuous, xlMedium, Color:=Colr(ColrNm)
End Sub

Sub BdrErzAy(RgAy() As Range, Optional ColrNm$ = "Red")
Dim R: For Each R In Itr(RgAy)
    BdrEr CvRg(R), ColrNm
Next
End Sub

Sub BDr(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub
Sub BdrNone(A As Range, Ix As XlBordersIndex)
A.Borders(Ix).LineStyle = xlLineStyleNone
End Sub

Sub BdrAround(A As Range)
BdrLeft A
BdrRight A
BdrTop A
BdrBottom A
End Sub
Function CvCmt(A) As Comment
Set CvCmt = A
End Function

Function IsCell(A As Range) As Boolean
If A.Rows.Count > 1 Then Exit Function
If A.Columns.Count > 1 Then Exit Function
IsCell = True
End Function

Function HasCell(A As Range, Cell As Range) As Boolean
If Not IsCell(Cell) Then Exit Function
If Not IsBet(Cell.Row, A.Row, R2zRg(A)) Then Exit Function
If Not IsBet(Cell.Column, A.Column, C2zRg(A)) Then Exit Function
HasCell = True
End Function
Function CmtAyzRg(A As Range) As Comment()
Dim C As Comment: For Each C In WszRg(A).Comments
    Dim CmtRg As Range: Set CmtRg = C.Parent
    If HasCell(A, CmtRg) Then PushObj CmtAyzRg, C
Next
End Function
Sub DltCmtzRg(A As Range)
Dim Cmt() As Comment: Cmt = CmtAyzRg(A)
Dim I: For Each I In Itr(Cmt)
    CvCmt(I).Delete
Next
End Sub
Sub BdrAroundNone(A As Range)
BdrNone A, xlInsideHorizontal
BdrNone A, xlInsideVertical
BdrNone A, xlEdgeLeft
BdrNone A, xlEdgeRight
BdrNone A, xlEdgeBottom
BdrNone A, xlEdgeTop
End Sub

Function RowzBelow(Rg As Range) As Range
Dim R&: R = R2zRg(Rg) + 1
Dim C1%: C1 = Rg.Column
Dim C2%: C2 = C2zRg(Rg)
Set RowzBelow = RgR(Rg, R)
End Function

Function RowzAbove(R As Range) As Range
Set RowzAbove = RgR(R, R.Rows.Count + 1)
End Function

Function ColzLeft(R As Range) As Range
Set ColzLeft = RgC(R, R.Column - 1)
End Function

Function ColzRight(R As Range) As Range
Set ColzRight = RgC(R, R.Columns.Count + 1)
End Function

Sub BdrBottom(A As Range)
BDr A, xlEdgeBottom
BDr RowzBelow(A), xlEdgeTop
End Sub

Sub BdrInside(A As Range)
BDr A, xlInsideHorizontal
BDr A, xlInsideVertical
End Sub

Sub BdrLeft(A As Range)
BDr A, xlEdgeLeft
BDr ColzRight(A), xlEdgeRight
End Sub

Sub BdrRight(A As Range)
BDr A, xlEdgeRight
If A.Column < MaxWsCol Then
    BDr RgC(A, A.Columns.Count + 1), xlEdgeLeft
End If
End Sub

Sub BdrTop(A As Range)
BDr A, xlEdgeTop
If A.Row > 1 Then
    BDr RgR(A, 0), xlEdgeBottom
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
BdrAround Rg
Rg.EntireColumn.AutoFit
SetLoNm O, Lon
Set LozRg = O
End Function

Function RgzDbt(Db As Database, T, At As Range) As Range
Set RgzDbt = RgzSq(SqzT(Db, T), At)
End Function

Function AddWszSq(A As Workbook, Sq(), Optional Wsn$) As Worksheet
Dim A1 As Range: Set A1 = A1zWs(AddWs(A, Wsn))
LozSq Sq, A1
Set AddWszSq = WszRg(A1)
End Function

Function AddWszT1(A As Workbook, Db As Database, T, Optional Wsn$, Optional AddgWay As EmAddgWay) As Worksheet
If AddgWay = EiSqWay Then
    Set AddWszT1 = AddWszT(A, Db, T, Wsn, AddgWay)
Else
    Set AddWszT1 = AddWszSq(A, SqzT(Db, T), Wsn)
End If
End Function

Function AddWszDrs(A As Workbook, B As Drs, Optional Wsn$) As Worksheet
Set AddWszDrs = AddWszSq(A, SqzDrs(B), Wsn)
End Function

Sub PutTbl(D As Database, T, At As Range, Optional AddgWay As EmAddgWay)
Select Case AddgWay
Case EmAddgWay.EiSqWay: PutSq SqzT(D, T), At
Case EmAddgWay.EiWcWay: AddLo At, D.Name, T
Case Else: ThwMsg CSub, "Invalid AddgWay"
End Select
End Sub

Sub AddWszTny(A As Workbook, Db As Database, Tny$(), Optional AddgWay As EmAddgWay)
Dim T$, I
For Each I In Tny
    T = I
    AddWszT A, Db, T, , AddgWay
Next
End Sub

Function AddWszDt(Wb As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = AddWs(Wb, Dt.DtNm)
LozDrs DrszDt(Dt), A1zWs(O)
Set AddWszDt = O
End Function

Function AddWc(ToWb As Workbook, FmFb, T) As WorkbookConnection
Set AddWc = ToWb.Connections.Add2(T, T, WcCnStrzFb(FmFb), T, XlCmdType.xlCmdTable)
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
Debug.Print ULin(N)
Debug.Print N
Debug.Print ULin(N)
DmpAy B
Return
End Sub

Sub ClsWbNoSav(A As Workbook)
A.Close False
End Sub

Sub DltWc(A As Workbook)
Dim WC As Excel.WorkbookConnection
For Each WC In A.Connections
    WC.Delete
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

Function HasWsCd(WsCdn$, Optional IsInf As Boolean) As Boolean
Dim Ws As Worksheet
For Each Ws In CWb.Sheets
    If Ws.CodeName = WsCdn Then HasWsCd = True: Exit Function
Next
If IsInf Then
    MsgBox RplVBar("No worksheet code||" & WsCdn), vbCritical
End If
End Function
Function HasWs(A As Workbook, WsIx) As Boolean
If IsNumeric(WsIx) Then
    HasWs = IsBet(WsIx, 1, A.Sheets.Count)
    Exit Function
End If
Dim Ws As Worksheet
For Each Ws In A.Worksheets
    If Ws.Name = WsIx Then HasWs = True: Exit Function
Next
End Function

Private Sub Z_WbWcsy()
'D WcStrAyWbOLE(WbzFx(TpFx))
End Sub

Private Sub Z_LozAyH()
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

Private Sub Z()
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
ThwWbMisOupNy C, H
ClsWbNoSav C
DltWc C
WbSavAs C, A
ShwWb C
XX = CWb()
End Sub

Function AddWs(A As Workbook, Optional Wsn$, Optional Pos As EmWsPos, Optional Aft$, Optional Bef$) As Worksheet
Dim O As Worksheet
DltWsIf A, Wsn
Select Case True
Case Pos = EiBeg:  Set O = A.Sheets.Add(FstWs(A))
Case Pos = EiEnd:  Set O = A.Sheets.Add(, LasWs(A))
Case Pos = EiRfWs And Aft <> "": Set O = A.Sheets.Add(, A.Sheets(Aft))
Case Pos = EiRfWs And Bef <> "": Set O = A.Sheets.Add(A.Sheets(Bef))
Case Else: Stop
End Select
SetWsn O, Wsn
Set AddWs = O
End Function
Private Sub Z_ClrLoRow()
DltLoRow CWs.ListObjects("T_SrcCd")
End Sub
Sub DltLoRow(A As ListObject)
Dim R As Range: Set R = A.DataBodyRange
If IsNothing(R) Then Exit Sub
R.ClearContents
Set R = RgA1(A.ListColumns(1).Range)
Dim R1 As Range: Set R1 = RgRR(R, 1, 2)
A.Resize R1
End Sub

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

Sub SavAsCls(Wb As Workbook, Fx)
Wb.SaveAs Fx
Wb.Close
End Sub

Sub SavAsFxm(Wb As Workbook, Fxm)
Wb.SaveAs Fxm, XlFileFormat.xlOpenXMLWorkbookMacroEnabled
End Sub

Function SavAsTmpFxm$(Wb As Workbook)
Dim O$: O = TmpFxm
SavAsFxm Wb, O
SavAsTmpFxm = O
End Function

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
    R2 = LasRnozRg(A.DataBodyRange)
    C2 = LasCnozRg(A.DataBodyRange)
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
Sub PutCd(Cd$(), CdLo As ListObject)
If ChkCdLo(CdLo) Then Exit Sub
DltLoRow CdLo
PutAyV Cd, A1zLo(CdLo)
End Sub
Private Function ChkCdLo(Lo As ListObject) As Boolean
If IsCdLo(Lo) Then Exit Function
MsgBox "Given Lo is not CdLo", "Lo-Name", Lo.Name
ChkCdLo = True
End Function
Private Function IsCdLo(A As ListObject) As Boolean
If A.ListColumns.Count <> 1 Then Exit Function
If A.ListColumns(1).Name <> "SrcCd" Then Exit Function
IsCdLo = True
End Function
Function PutSq(Sq(), At As Range) As Range
Dim O As Range
If NRowzSq(Sq) = 0 Then
    Set PutSq = RgA1(At)
    Exit Function
End If
Set O = ResiRg(At, Sq)
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

Private Sub Z_NewXls()
Dim X As Excel.Application
Set X = NewXls
Stop
End Sub

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
Dim WC As WorkbookConnection
For Each WC In Wb.Connections
    ClsWc WC
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
    Cn = OleCnStrzFb(ToUseFb)
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
FmtLoBStd Wb
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

Private Sub RplLozT(A As ListObject, D As Database, T)
Dim Fny1$(): Fny1 = Fny(D, T)
Dim Fny2$(): Fny2 = FnyzLo(A)
If Not IsAySam(Fny1, Fny2) Then
    Thw CSub, "LoFny and TblFny are not same", "LoFny T TblFny Db", Fny2, T, Fny1, D.Name
End If
Dim Sq()
    Dim R As dao.Recordset
    Set R = Rs(A, SqlSel_Fny_T(Fny2, T))
    Sq = AddSngQtezSq(SqzRs(R))
MinxLo A
RgzSq Sq, A.DataBodyRange
End Sub


Sub PutWc(WC As WorkbookConnection, At As Range)
Dim Lo As ListObject
Set Lo = WszRg(At).ListObjects.Add(SourceType:=0, Source:=WC.OLEDBConnection.Connection, Destination:=At)
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = WC.Name
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
    .ListObject.DisplayName = Lon(WC.Name)
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
SavAsCls NewWbzOupTbl(Fb, AddgWay), Fx
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
    'WszDt O, Ay(J)
Next
Set WbzDs = O
End Function
Sub PutSeqDown(At As Range, FmNum&, ToNum&)
PutAyV LngSeq(FmNum, ToNum), At
End Sub

Sub DltSheet1(Wb As Workbook)
DltWs Wb, "Sheet1"
End Sub
Sub ActWs(Ws As Worksheet)
If IsEqObj(Ws, CWs) Then Exit Sub
Ws.Activate
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
CellBelow(Cell).Clear
End Sub

Sub FillSeqV(VBar As Range)
Dim Sq()
Sq = SqHzN(VBar.Rows.Count)
ResiRg(VBar, Sq).Value = Sq
End Sub

Sub FillWsny(At As Range)
RgzAyV Wsny(WbzRg(At)), At
End Sub


Sub FillAtV(At As Range, Ay)
FillSq SqV(Ay), At
End Sub

Sub FillLc(Lo As ListObject, ColNm$, Ay)
If Lo.DataBodyRange.Rows.Count <> Si(Ay) Then Thw CSub, "Lo-NRow <> Si(Ay)", "Lo-NRow ,Si(Ay)", NRowzLo(Lo), Si(Ay)
Dim At As Range, C As ListColumn, R As Range
'DmpAy FnyzLo(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
FillAtV At, Ay
End Sub
Sub FillSq(Sq(), At As Range)
ResiRg(At, Sq).Value = Sq
End Sub
Sub FillAtH(Ay, At As Range)
FillSq SqH(Ay), At
End Sub

Sub RunFxqByCn(Fx, Q)
CnzFx(Fx).Execute Q
End Sub
Function DKVzKSet(KSet As Dictionary) As Drs
Dim K, Dy(): For Each K In KSet.Keys
    Dim S As Aset: Set S = KSet(K)
    Dim V: For Each V In S.Itms
        PushI Dy, Array(K, V)
    Next
Next
DKVzKSet = DrszFF("K V", Dy)
End Function
Private Sub Z_DKVzLoFilter()
Dim Lo As ListObject: Set Lo = FstLo(CWs)
BrwDrs DKVzLoFilter(Lo)
End Sub
Function DKVzLoFilter(L As ListObject) As Drs
DKVzLoFilter = DKVzKSet(KSetzLoFilter(L))
End Function

Function KSetzKyAsetAy(Ky$(), AsetAy() As Aset) As Dictionary
Set KSetzKyAsetAy = New Dictionary
Dim K, J&: For Each K In Itr(Ky)
    KSetzKyAsetAy.Add K, AsetAy(J)
    J = J + 1
Next
End Function
Sub SetOnFilter(L As ListObject)
On Error GoTo X
Dim M As Boolean: M = L.AutoFilter.FilterMode ' If filter is on, it will have no err, otherwise, there is err
Exit Sub
X:
L.Range.AutoFilter 'Turn on
End Sub
Function KSetzLoFilter(L As ListObject) As Dictionary
'Ret : KSet
Dim O As Dictionary: Set O = New Dictionary
SetOnFilter L
Dim Fny$(): Fny = FnyzLo(L)
Dim F As Filter, J%: For Each F In L.AutoFilter.Filters
    Dim K$: K = Fny(J)
    KSetzLoFilter__Add O, K, F
    J = J + 1
Next
Set KSetzLoFilter = O
End Function

Private Sub KSetzLoFilter__Add(OKSet As Dictionary, K$, F As Filter)
If Not F.On Then Exit Sub
If F.Operator <> xlFilterValues Then Exit Sub
Dim S As Aset: Set S = AsetzAy(RmvPfxzAy(F.Criteria1, "="))
OKSet.Add K, S
End Sub

Function RgzDrs(A As Drs, At As Range) As Range
Set RgzDrs = RgzSq(SqzDrs(A), At)
End Function

Function LozDrs(A As Drs, At As Range, Optional Lon$) As ListObject
Set LozDrs = LozRg(RgzDrs(A, At), Lon)
End Function

Function WszAy(Ay, Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet, R As Range
Set O = NewWs(Wsn)
O.Range("A1").Value = "Array"
Set R = RgzSq(SqV(Ay), O.Range("A2"))
LozRg RgMoreTop(R)
Set WszAy = O
End Function

Function WszDrs(A As Drs, Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(Wsn)
LozDrs A, O.Range("A1")
Set WszDrs = O
End Function

Function RgzAyV(Ay, At As Range) As Range
Set RgzAyV = RgzSq(SqV(Ay), At)
End Function

Function RgzAyH(Ay, At As Range) As Range
Set RgzAyH = RgzSq(SqH(Ay), At)
End Function

Function RgzDy(Dy(), At As Range) As Range
Set RgzDy = RgzSq(SqzDy(Dy), At)
End Function

Function WszDy(Dy(), Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(Wsn)
RgzDy Dy, A1zWs(O)
Set WszDy = O
End Function


Function WszDs(A As Ds) As Worksheet
Dim O As Worksheet: Set O = NewWs
A1zWs(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = WsRC(O, 2, 1)
Dim BelowN&, Dt As Dt, Ay() As Dt
Ay = A.Ay
For J = 0 To A.N - 1
    Dt = Ay(J)
    LozDt Dt, At
    BelowN = 2 + Si(Dt.Dy)
    Set At = CellBelow(At, BelowN)
Next
Set WszDs = O
End Function

Function RgzDt(A As Dt, At As Range, Optional DtIx%)
Dim Pfx$: If DtIx > 0 Then Pfx = QteBkt(CStr(DtIx))
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


Function RgzSq(Sq(), At As Range) As Range
If Si(Sq) = 0 Then
    Set RgzSq = RgA1(At)
    Exit Function
End If
Dim O As Range
Set O = ResiRg(At, Sq)
O.MergeCells = False
O.Value = Sq
Set RgzSq = O
End Function

Private Sub Z_WszDs()
ShwWs WszDs(SampDs)
End Sub

Function EnsWbzXls(Xls As Excel.Application, Wbn$) As Workbook
Dim O As Workbook
Const FxFn$ = "Insp.xlsx"
If HasWbn(Xls, FxFn) Then
    Set O = Xls.Workbooks(FxFn)
Else
    Set O = Xls.Workbooks.Add
    O.SaveAs InstFdr("Insp") & "Insp.xlsx"
End If
Set EnsWbzXls = ShwWb(O)
End Function


Sub StopXls()
Const Ps1Str$ = "Stop-Process -Id{try{(Get-Process -Name Excel).Id}finally{@()}}.invoke()"
Dim F$: F = TmpHom & "StopXls.ps1"
Static X As Boolean
If Not X Then
    X = True
    If NoFfn(F) Then
        WrtStr Ps1Str, F
    End If
End If
'PowerRun F
End Sub
