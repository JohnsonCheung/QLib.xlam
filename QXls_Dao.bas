Attribute VB_Name = "QXls_Dao"
Option Explicit
Private Const CMod$ = "MXls_Dao."
Private Const Asm$ = "QXls"

Function CvCn(A) As AdoDb.Connection
Set CvCn = A
End Function

Sub RplOleWcFb(Wc As WorkbookConnection, Fb$)
CvCn(Wc.OLEDBConnection.ADOConnection).ConnectionString = CnStrzFbzAsAdo(Fb)
End Sub

Sub RplLozFbzFbt(Lo As ListObject, Fb$, T$)
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

Function WbzFb(Fb$, Optional Vis As Boolean) As Workbook
Dim D As Database: Set D = Db(Fb$)
Set WbzFb = SetViszWb(WbzTny(D, Tny(D)), Vis)
End Function

Sub PutTny(A As Database, Tny$(), ToWb As Workbook, Optional AddgWay As EmAddgWay)
For Each I In Tny
    T = I
    PutTbl A, T, A1, AddgWay
Next
End Sub

Sub PutTbl(A As Database, T$, At As Range, Optional AddgWay As EmAddgWay)
Select Case AddgWay
Case EmAddgWay.EiSqWay: PutSq SqzT(A, T), At
Case EmAddgWay.EiWcWay: AddLo At, A.Name, T
Case Else: Thw CSub, "Invalid AddgWay"
End Select
End Sub

Function SetWsn(Ws As Worksheet, Nm$) As Worksheet
Set SetWsn = Ws
If Nm = "" Then Exit Function
If HasWs(WbzWs(Ws), Nm) Then
    Dim Wb As Workbook: Set Wb = WbzWs(Ws)
    Thw CSub, FmtQQ("Wsn exists in Wb", "Wsn WbNm Wny-in-Wb", Nm, WbNm(Wb), WnyzWb(Wb))
End If
Ws.Name = Nm
End Function
Sub AddLozSamp()
'    Application.CutCopyMode = False
'    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
'        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Users\user\Desktop\SAPAccessReports\DutyPrepay5\DutyP" _
'        , _
'        "repay5_Data.mdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:D" _
'        , _
'        "atabase Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Glob" _
'        , _
'        "al Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=Fals" _
'        , _
'        "e;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Supp" _
'        , _
'        "ort Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceFiel" _
'        , "d Validation=False"), Destination:=Range("$H$4")).QueryTable
'        .CommandType = xlCmdTable
'        .CommandText = Array("@RptM")
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = True
'        .SourceDataFile = _
'        "C:\Users\user\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
'        .ListObject.DisplayName = "Table_DutyPrepay5_Data_1"
'        .Refresh BackgroundQuery:=False
'    End With

End Sub
Function AddLo(At As Range, Fb$, T$) As ListObject
Dim Ws As Worksheet: Set Ws = WszRg(At)
Dim Lo As ListObject: Set Lo = Ws.ListObjects.Add(xlSrcExternal, CnStrzFbzAsAdoOle(Fb), Destination:=At)
Dim Qt As QueryTable: Set Qt = Lo.QueryTable
With Qt
    .CommandType = xlCmdTable
    .CommandText = T
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = False
    .RefreshPeriod = 0
    .PreserveColumnInfo = True
    .ListObject.DisplayName = LoNmzTblNm(T)
    .Refresh BackgroundQuery:=False
End With
LoAutoFit Lo
End Function
Sub LoAutoFit(A As ListObject)
A.DataBodyRange.EntireColumn.AutoFit
End Sub
Function AddWszT(Wb As Workbook, Db As Database, T$, Optional Wsn0$, Optional AddgWay As EmAddgWay) As Worksheet
Dim O As Worksheet: Set O = AddWs(Wb, StrDft(Wsn0, T))
Dim A1 As Range: Set A1 = A1zWs(O)
PutTbl Db, T, A1, AddgWay
End Function

Function NewWbzOupTbl(Fb$, Optional AddgWay As EmAddgWay) As Workbook '
Dim O As Workbook, D As Database
Set O = NewWb
Set D = Db(Fb)
AddWszTny O, D, OupTny(D), AddgWay
DltWsIf O, "Sheet1"
Set NewWbzOupTbl = O
End Function

Function WbzT(A As Database, T$, Optional Wsn$ = "Data", Optional LoNm$, Optional Vis As Boolean) As Workbook
Set WbzT = WszRg(AtAddDbt(NewA1(Wsn, Vis), A, T, LoNm))
End Function
Function AtAddDbt(At As Range, Db As Database, T$, Optional LoNm$) As Range
'CrtLozRg PutSq(At, Dbt(Db, T).Sq), LoNm
Set AtAddDbt = At
End Function
Sub PutDbtWs(A As Database, T$, Ws As Worksheet)
PutDbtAt A, T, A1zWs(Ws)
End Sub

Sub PutDbtAt(A As Database, T$, At As Range, Optional AddgWay As EmAddgWay)
CrtLozRg PutSq(SqzDbt(A, T), At), LoNm(T)
End Sub
Sub SetQtFbt(Qt As QueryTable, Fb$, T$)
With Qt
    .CommandType = xlCmdTable
    .Connection = CnStrzFbzAsAdoOle(Fb$) '<--- Fb
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
Sub PutFbtAt(Fb$, T$, At As Range, Optional LoNm0$)
Dim O As ListObject
Set O = WszRg(At).ListObjects.Add(SourceType:=XlSourceType.xlSourceWorkbook, Destination:=At)
SetLoNm O, Dft(LoNm0, LoNm(T))
SetQtFbt O.QueryTable, Fb, T
End Sub
Sub FxzTny(Fx$, Db As Database, Tny$())
WbzTny(Db, Tny).SaveAs Fx
End Sub

Function WszT(A As Database, T$, Optional Wsn$) As Worksheet
'set Wszt = WszT(NewWb(
Dim Sq(): Sq = SqzT(A, T)
Dim A1 As Range: Set A1 = NewA1(Wsn)
Set WszT = WszLo(CrtLozSq(Sq(), A1))
End Function

