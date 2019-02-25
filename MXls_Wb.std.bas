Attribute VB_Name = "MXls_Wb"
Option Explicit
Property Get CurWb() As Workbook
Set CurWb = Xls.ActiveWorkbook
End Property

Function CvWb(A) As Workbook
Set CvWb = A
End Function

Function TxtCnWc(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set TxtCnWc = A.TextConnection
End Function

Function FstWs(A As Workbook) As Worksheet
Set FstWs = A.Sheets(1)
End Function

Function FxWb$(A As Workbook)
Dim F$
F = A.FullName
If F = A.Name Then Exit Function
FxWb = F
End Function

Function LasWs(A As Workbook) As Worksheet
Set LasWs = A.Sheets(A.Sheets.Count)
End Function


Function LoItr(A As Workbook)
LoItr = Itr(LoAy(A))
End Function

Function LozAyH(Ay, Wb As Workbook, Optional Wsn$, Optional LoNm$) As ListObject
Set LozAyH = LozRg(RgvSqH(SqzAyV(Ay), A1Wb(Wb, Wsn)), LoNm)
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
Set MainWs = WbzWsCdNm(A, "WsOMain")
End Function

Function OupLoAy(A As Workbook) As ListObject()
OupLoAy = OywNmPfx(LoAy(A), "T_")
End Function

Function Wbs(A As Workbook) As Workbooks
Set Wbs = A.Parent
End Function

Function PtNy(A As Workbook) As String()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushIAy PtNy, PtNyzWs(Ws)
Next
End Function

Function TxtWc(A As Workbook) As TextConnection
Dim N%: N = TxtWcCnt(A)
If N <> 1 Then
    Stop
    Exit Function
End If
Dim C As WorkbookConnection
For Each C In A.Connections
    If Not IsNothing(TxtCnWc(C)) Then
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
    If Not IsNothing(TxtCnWc(C)) Then Cnt = Cnt + 1
Next
TxtWcCnt = Cnt
End Function

Function TxtWcStr$(A As Workbook)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = TxtWc(A)
If IsNothing(T) Then Exit Function
TxtWcStr = T.Connection
End Function

Function OleWcAy(A As Workbook) As OLEDBConnection()
Dim O() As OLEDBConnection, Wc As WorkbookConnection
For Each Wc In A.Connections
    PushObjExlNothing O, Wc.OLEDBConnection
Next
OleWcAy = OyeNothing(IntozItrPrp(A.Connections, "OLEDBConnection", OleWcAy))
End Function

Function WcNyWb(A As Workbook) As String()
WcNyWb = Itn(A.Connections)
End Function

Function WcStrAyWbOLE(A As Workbook) As String()
WcStrAyWbOLE = SyOyP(OleWcAy(A), "Connection")
End Function

Function WsWb(A As Workbook, Wsn) As Worksheet
Set WsWb = A.Sheets(Wsn)
End Function

Function WsNyzRg(A As Range) As String()
WsNyzRg = WsNy(WbzRg(A))
End Function

Function WsNy(A As Workbook) As String()
WsNy = Itn(A.Sheets)
End Function

Function WbzWsCdNm(A As Workbook, CdNm$) As Worksheet
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws.CodeName = CdNm Then Set WbzWsCdNm = Ws: Exit Function
Next
End Function

Function WsCdNyWb(A As Workbook) As String()
WsCdNyWb = SyItrPrp(A.Sheets, "CodeName")
End Function

Function WbAddDbt(A As Workbook, Db As Database, T$, Optional UseWc As Boolean) As Workbook
'Set WbAddDbt = WbLo(RplLoCnzDbt(Db, T, A1zWs(A, T), UseWc))
End Function
Function WbFullNm$(A As Workbook)
On Error Resume Next
WbFullNm = A.FullName
End Function

Function AddWczWbDbtt(A As Workbook, Db As Database, TT, Optional UseWc As Boolean) As Workbook
DoAyPPXP CvNy(TT), "WbAddDbt", A, Db, UseWc
Set AddWczWbDbtt = A
End Function

Function AddWszDt(A As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = AddWs(A, Dt.DtNm)
LozDrs DrszDt(Dt), A1(O)
Set AddWszDt = O
End Function

Function AddWczWbFb(A As Workbook, LnkToFb$, WcNm) As WorkbookConnection
Set AddWczWbFb = A.Connections.Add2(WcNm, WcNm, CnStrzFbForWbCn(LnkToFb), WcNm, XlCmdType.xlCmdTable)
End Function

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
Set AddWs = SetWsNm(O, Wsn)
End Function

Sub ThwWbMisOupNy(A As Workbook, OupNy$())
Dim O$(), N$, B$(), WsCdNy$()
WsCdNy = WsCdNyWb(A)
O = AyMinus(AyAddPfx(OupNy, "WsO"), WsCdNy)
If Sz(O) > 0 Then
    N = "OupNy":  B = OupNy:  GoSub Dmp
    N = "WbCdNy": B = WsCdNy: GoSub Dmp
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

Sub DltWs(A As Workbook, Wsn)
If HasWbzWs(A, Wsn) Then
    A.Application.DisplayAlerts = False
    WsWb(A, Wsn).Delete
    A.Application.DisplayAlerts = True
End If
End Sub

Function WbMax(A As Workbook) As Workbook
A.Application.WindowState = xlMaximized
Set WbMax = A
End Function

Function NewA1Wb(A As Workbook, Optional Wsn$) As Range
'Set NewA1Wb = A1zWs(AddWs(A, Wsn))
End Function

Sub WbQuit(A As Workbook)
XlsQuit A.Application
End Sub

Function WbSav(A As Workbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.Save
A.Application.DisplayAlerts = Y
Set WbSav = A
End Function

Function SavWbAs(A As Workbook, Fx, Optional Fmt As XlFileFormat = xlOpenXMLWorkbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.SaveAs Fx, Fmt
A.Application.DisplayAlerts = Y
Set SavWbAs = A
End Function

Sub SetWbFcsvCn(A As Workbook, Fcsv$)
'Set first Wb TextConnection to Fcsv if any
Dim T As TextConnection: Set T = TxtWc(A)
Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub

Function WbVis(A As Workbook) As Workbook
XlsVis A.Application
Set WbVis = A
End Function

Function HasWbzWs(A As Workbook, Wsn) As Boolean
HasWbzWs = HasItn(A.Sheets, Wsn)
End Function

Private Sub ZZ_WbWcSy()
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

Private Sub Z_SetWbFcsvCn()
Dim Wb As Workbook
'Set Wb = WbzFx(Vbe_MthFx)
Debug.Print TxtWcStr(Wb)
SetWbFcsvCn Wb, "C:\ABC.CSV"
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
TxtCnWc B
FstWs C
FxWb C
LasWs C
LozWb C, D
MainWs C
OupLoAy C
Wbs C
TxtWc C
TxtWcCnt C
TxtWcStr C
OleWcAy C
WsCdNyWb C
WcStrAyWbOLE C
WsWb C, A
WsNy C
WbzWsCdNm C, D
WbzWsCdNm C, D
AddWczWbFb C, D, D
AddWs C, D, F, F, D, D
ThwWbMisOupNy C, H
ClsWbNoSav C
DltWc C
DltWs C, A
SavWbAs C, A
SetWbFcsvCn C, D
WbVis C
HasWbzWs C, A
XX = CurWb()
End Sub

Private Sub Z()
Z_TxtWcCnt
Z_SetWbFcsvCn
End Sub
