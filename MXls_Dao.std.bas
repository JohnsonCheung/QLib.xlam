Attribute VB_Name = "MXls_Dao"
Option Explicit

Function RplLoCnzDbt(A As ListObject, Db As Database, T) As ListObject
Dim Sq(), Drs As Drs, R As DAO.Recordset
Set R = Rsz(Db, T)
If Not IsEqAy(FnyzRs(R), FnyzLo(A)) Then
    Debug.Print "--"
    Debug.Print "Rs"
    Debug.Print "--"
    DmpAy FnyzRs(R)
    Debug.Print "--"
    Debug.Print "A"
    Debug.Print "--"
    DmpAy FnyzLo(A)
    Stop
End If
Sq = SqAddSngQuote(SqzRs(R))
MinxLo A
'RgzSq Sq, A.DataBodyRange
Set RplLoCnzDbt = A
End Function

Function CvCn(A) As ADODB.Connection
Set CvCn = A
End Function
Sub RplOleWcFb(Wc As WorkbookConnection, Fb)
CvCn(Wc.OLEDBConnection.ADOConnection).ConnectionString = CnStrzFbAdo(Fb)
End Sub
Sub RplLoCnzFbt(Lo As ListObject, Fb As Database, T)
With Lo.QueryTable
    RplOleWcFb .Connection, Fb '<==
    .CommandType = xlCmdTable
    .CommandText = T '<==
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .ListObject.DisplayName = LoNm(T) '<==
    .Refresh BackgroundQuery:=False
End With
End Sub

Function WbzDbtt(Db As Database, TT) As Workbook
Dim O As Workbook, T
Set O = NewWb
For Each T In TermAy(TT)
    AddWszDbt O, Db, T
Next
Set WbzDbtt = O
End Function
Function SetWsn(Ws As Worksheet, Nm$) As Worksheet
If Nm = "" Then Exit Function
Ws.Name = Nm
Set SetWsn = Ws
End Function

Sub AddWszDbt(Wb As Workbook, Db As Database, T, Optional Wsn0$)
Dim Wsn$: Wsn = Dft(Wsn0, T)
AddWs Wb, Wsn
PutDbtAt Db, T, A1zWs(LasWs(Wb))
End Sub

Sub AddWszT(Wb As Workbook, T, Optional Wsn0$)
AddWszDbt Wb, CDb, T, Wsn0$
End Sub

Function WbzOupTblDb(Db As Database) As Workbook
Set WbzOupTblDb = WbzDbtt(Db, OupTnyz(Db))
End Function

Function WbzT(Db As Database, T, Optional Wsn$ = "Data", Optional LoNm$, Optional Vis As Boolean) As Workbook
Set WbzT = WszRg(AtAddDbt(NewA1(Wsn, Vis), Db, T, LoNm))
End Function
Function AtAddDbt(At As Range, Db As Database, T, Optional LoNm$) As Range
'LozRg AtAddSq(At, Dbt(Db, T).Sq), LoNm
Set AtAddDbt = At
End Function
Sub PutDbtWs(A As Database, T, Ws As Worksheet)
PutDbtAt A, T, A1zWs(Ws)
End Sub

Sub PutDbtAt(A As Database, T, At As Range, Optional LoNm$)
'LozRg AtAddSq(At, Dbt(A, T).Sq), Dft(LoNm, LoNm(T))
End Sub
Sub SetQtFbt(Qt As QueryTable, Fb$, T)
With Qt
    .CommandType = xlCmdTable
    .Connection = CnStrzFbAdoOle(Fb) '<--- Fb
    .CommandText = T '<-----  T
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
    .Refresh BackgroundQuery:=False
End With
End Sub
Sub PutFbtAt(Fb$, T, At As Range, Optional LoNm0$)
Dim O As ListObject
Set O = WszRg(At).ListObjects.Add(SourceType:=XlSourceType.xlSourceWorkbook, Destination:=At)
SetLoNm O, Dft(LoNm0, LoNm(T))
SetQtFbt O.QueryTable, Fb, T
End Sub
Sub CrtFxDbtt(Fx$, Db As Database, TT)
WbDbtt(Db, TT).SaveAs Fx
End Sub

Function WbDbtt(Db As Database, TT, Optional UseWc As Boolean) As Workbook
Dim O As Workbook
Set O = NewWb
Set WbDbtt = AddWczWbDbtt(O, Db, TT, UseWc)
WsWb(O, "Sheet1").Delete
Set WbDbtt = O
End Function

Sub CrtFxTT(Fx$, TT$)
CrtFxDbtt Fx, CDb, TT
End Sub

Function WszDbt(Db As Database, T, Optional Wsn$) As Worksheet
Dim Sq(): Sq = SqzDbt(Db, T)
Dim A1 As Range: Set A1 = NewA1(Wsn)
Set WszDbt = WsLo(LozSq(Sq, A1))
End Function
Function WszT(T, Optional Wsn$ = "Data") As Worksheet
Set WszT = WszDbt(CDb, T, Wsn)
End Function
