Attribute VB_Name = "QXls_Dao"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_Dao."
Private Const Asm$ = "QXls"

Function CvCn(A) As AdoDb.Connection
Set CvCn = A
End Function

Sub RplOleWcFb(WC As WorkbookConnection, Fb)
CvCn(WC.OLEDBConnection.ADOConnection).ConnectionString = CnStrzFbzAsAdo(Fb)
End Sub

Sub RplLozFbzFbt(Lo As ListObject, Fb, T)
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
    .ListObject.DisplayName = Lon(T) '<==
    .Refresh BackgroundQuery:=False
End With
End Sub
Function TmpInpTny(A As Database) As String()
TmpInpTny = AywPfx(Tny(A), "#I")
End Function

Private Sub ZZ_LoIxSq()
Dim WB As Workbook: Set WB = NewWb
AddWszzWbSq WB, SampSq
AddWszzWbSq WB, SampSq1
BrwSq LoIxSq(WB)
End Sub

Sub AddWszzWbSq(WB As Workbook, Sq())
LozSq Sq, A1zWs(AddWs(WB))
End Sub

Function LoIxSq(WB As Workbook) As Variant()
Dim Ws As Worksheet, M(), O(), Fnd As Boolean
For Each Ws In WB.Sheets
    M = LoIxSqzWs(Ws)
    If Si(M) > 0 Then
        If Fnd Then
            PushSq O, M
        Else
            O = M
        End If
    End If
Next
LoIxSq = O
End Function

Private Function LoIxSqzWs(Ws As Worksheet) As Variant()
Dim Lo As ListObject, R&, NR&
NR = Ws.ListObjects.Count
ReDim O(1 To NR, 1 To 4)
For Each Lo In Ws.ListObjects
    R = R + 1
    SetSqr O, LoDr(Lo), R
Next
End Function

Private Sub ZZ_LoDr()
Dim Lo As ListObject: Set Lo = SampLo
D LoDr(Lo)
ClsWbNoSav WbzLo(Lo)
End Sub

Private Function LoDr(A As ListObject) As Variant()
Dim WN$: WN = WsnzLo(A)
Dim LN$:: LN = A.Name
Dim NR&: NR = NRowzLo(A)
Dim Nc&: Nc = A.ListColumns.Count
LoDr = Array(WN, LN, NR, Nc)
End Function

Sub AddLoIx(At As Range)
LozSq LoIxSq(WbzRg(At)), At
End Sub

Function WbzTmpInp(A As Database) As Workbook
Set WbzTmpInp = WbzTny(A, TmpInpTny(A))
End Function

Function WbzTny(A As Database, Tny$()) As Workbook
Dim T, O As Workbook
Set O = NewWb
For Each T In Itr(Tny)
    AddWszT O, A, CStr(T)
Next
DltSheet1 O
Set WbzTny = O
End Function
Function WbzFb(Fb) As Workbook
Dim D As Database: Set D = Db(Fb)
Set WbzFb = ShwWb(WbzTny(D, Tny(D)))
End Function

Function SetWsn(Ws As Worksheet, Nm$) As Worksheet
Set SetWsn = Ws
If Nm = "" Then Exit Function
If HasWs(WbzWs(Ws), Nm) Then
    Dim WB As Workbook: Set WB = WbzWs(Ws)
    Thw CSub, FmtQQ("Wsn exists in Wb", "Wsn Wbn Wny-in-Wb", Nm, Wbn(WB), WnyzWb(WB))
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
Function AddLo(At As Range, Fb, T) As ListObject
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
    .ListObject.DisplayName = LoNmzT(T)
    .Refresh BackgroundQuery:=False
End With
LoAutoFit Lo
End Function
Sub LoAutoFit(A As ListObject)
A.DataBodyRange.EntireColumn.AutoFit
End Sub

Function AddWszT(WB As Workbook, Db As Database, T, Optional Wsn0$, Optional AddgWay As EmAddgWay) As Worksheet
Dim O As Worksheet: Set O = AddWs(WB, StrDft(Wsn0, T))
Dim A1 As Range: Set A1 = A1zWs(O)
PutTbl Db, T, A1, AddgWay
End Function

Function NewWbzOupTbl(Fb, Optional AddgWay As EmAddgWay) As Workbook '
Dim O As Workbook, D As Database
Set O = NewWb
Set D = Db(Fb)
AddWszTny O, D, OupTny(D), AddgWay
DltWsIf O, "Sheet1"
Set NewWbzOupTbl = O
End Function

Function WbzT(A As Database, T, Optional Wsn$ = "Data") As Workbook
Set WbzT = WszRg(AddWszT(NewWb, A, T, Wsn))
End Function

Sub PutDbtWs(A As Database, T, Ws As Worksheet)
PutDbtAt A, T, A1zWs(Ws)
End Sub
Sub ClrLo(A As ListObject)
If A.ListRows.Count = 0 Then Exit Sub
A.DataBodyRange.Delete xlShiftUp
End Sub
Sub PutAyAtV(Ay, At As Range)
PutSq SqzAyV(Ay), At
End Sub
Function CrtLo(Ws As Worksheet, FF$, Optional Lon$) As ListObject
Set CrtLo = LozRg(RgzAyH(SyzSS(FF), A1zWs(Ws)), Lon)
End Function
Sub PutAyAtH(Ay, At As Range)
PutSq SqzAyH(Ay), At
End Sub
Sub PutDbtAt(A As Database, T, At As Range, Optional AddgWay As EmAddgWay)
LozRg PutSq(SqzT(A, T), At), Lon(T)
End Sub
Sub SetQtFbt(Qt As QueryTable, Fb, T)
With Qt
    .CommandType = xlCmdTable
    .Connection = CnStrzFbzAsAdoOle(Fb) '<--- Fb
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
Sub PutFbtAt(Fb, T$, At As Range, Optional LoNm0$)
Dim O As ListObject
Set O = WszRg(At).ListObjects.Add(SourceType:=XlSourceType.xlSourceWorkbook, Destination:=At)
SetLoNm O, Dft(LoNm0, Lon(T))
SetQtFbt O.QueryTable, Fb, T
End Sub
Sub FxzTny(Fx, Db As Database, Tny$())
WbzTny(Db, Tny).SaveAs Fx
End Sub

Function WszT(A As Database, T, Optional Wsn$) As Worksheet
Dim Sq(): Sq = SqzT(A, T)
Dim A1 As Range: Set A1 = NewA1(Wsn)
Set WszT = WszLo(LozSq(Sq(), A1))
End Function

