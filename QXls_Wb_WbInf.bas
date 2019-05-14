Attribute VB_Name = "QXls_Wb_WbInf"
Option Explicit
Private Const CMod$ = "MXls_Wb."
Private Const Asm$ = "QXls"
Property Get CurWb() As Workbook
Set CurWb = Xls.ActiveWorkbook
End Property

Function LoAy(A As Workbook) As ListObject()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushObjzItr LoAy, Ws.ListObjects
Next
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

Function LasWs(A As Workbook) As Worksheet
Set LasWs = A.Sheets(A.Sheets.Count)
End Function

Function LoItr(A As Workbook)
LoItr = Itr(LoAy(A))
End Function

Function CrtLozAyH(Ay, Wb As Workbook, Optional Wsn$, Optional LoNm$) As ListObject
Set CrtLozAyH = CrtLozRg(RgzSq(Sqv(Ay), A1zWb(Wb, Wsn)), LoNm)
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

Private Sub Z_SetWsCdNm()
Dim A As Worksheet: Set A = NewWs
SetWsCdNm A, "XX"
ShwWs A
Stop
End Sub

Sub SetWsCdNm(A As Worksheet, CdNm$)
CmpzWs(A).Name = CdNm
End Sub

Sub SetWsCdNmAndLoNm(A As Worksheet, Nm$)
CmpzWs(A).Name = Nm
SetLoNm FstLo(A), Nm
End Sub

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
Function RgzDbtzByWc(Db As Database, T, At As Range) As Range

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

Sub DltWs(A As Workbook, WsIx)
A.Application.DisplayAlerts = False
WszWb(A, WsIx).Delete
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

Function HasWs(A As Workbook, W) As Boolean
If IsNumeric(W) Then
    HasWs = IsBet(W, 1, A.Sheets.Count)
    Exit Function
End If
HasWs = HasItn(A.Sheets, CStr(W))
End Function

Private Sub ZZ_WbWcsy()
'D WcStrAyWbOLE(WbzFx(TpFx))
End Sub

Private Sub ZZ_CrtLozAyH()
'D NyOy(CrtLozAyH(TpWb))
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
Wbs C
TxtWc C
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
XX = CurWb()
End Sub

