Attribute VB_Name = "QXls_F_Dao"
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
Function TmpInpTny(D As Database) As String()
TmpInpTny = AwPfx(Tny(D), "#I")
End Function

Private Sub Z_LoIxSq()
Dim Wb As Workbook: Set Wb = NewWb
AddWszzWbSq Wb, SampSq
AddWszzWbSq Wb, SampSq1
BrwSq LoIxSq(Wb)
End Sub

Sub AddWszzWbSq(Wb As Workbook, Sq())
LozSq Sq, A1zWs(AddWs(Wb))
End Sub

Function LoIxSq(Wb As Workbook) As Variant()
Dim Ws As Worksheet, M(), O(), Fnd As Boolean
For Each Ws In Wb.Sheets
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

Private Sub Z_LoDr()
Dim Lo As ListObject: Set Lo = SampLo
D LoDr(Lo)
ClsWbNoSav WbzLo(Lo)
End Sub

Private Function LoDr(A As ListObject) As Variant()
Dim WN$: WN = WsnzLo(A)
Dim LN$:: LN = A.Name
Dim NR&: NR = NRowzLo(A)
Dim NC&: NC = A.ListColumns.Count
LoDr = Array(WN, LN, NR, NC)
End Function

Sub AddLoIx(At As Range)
LozSq LoIxSq(WbzRg(At)), At
End Sub

Function WbzTmpInp(D As Database) As Workbook
Set WbzTmpInp = WbzTny(D, TmpInpTny(D))
End Function

Function WbzTny(D As Database, Tny$()) As Workbook
Dim T, O As Workbook
Set O = NewWb
For Each T In Itr(Tny)
    AddWszT O, D, CStr(T)
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
If Ws.Name = Nm Then Exit Function
If HasWs(WbzWs(Ws), Nm) Then
    Dim Wb As Workbook: Set Wb = WbzWs(Ws)
    Thw CSub, "Wsn exists in Wb", "Wsn Wbn Wny-in-Wb", Nm, Wbn(Wb), WnyzWb(Wb)
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
Dim QT As QueryTable: Set QT = Lo.QueryTable
With QT
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

Function AddWszT(Wb As Workbook, Db As Database, T, Optional Wsn0$, Optional AddgWay As EmAddgWay) As Worksheet
Dim O As Worksheet: Set O = AddWs(Wb, StrDft(Wsn0, T))
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

Function WbzT(D As Database, T, Optional Wsn$ = "Data") As Workbook
Set WbzT = WszRg(AddWszT(NewWb, D, T, Wsn))
End Function

Sub PutDbtWs(D As Database, T, Ws As Worksheet)
PutDbtAt D, T, A1zWs(Ws)
End Sub
Sub ClrLo(A As ListObject)
If A.ListRows.Count = 0 Then Exit Sub
A.DataBodyRange.Delete xlShiftUp
End Sub
Sub PutAyV(AyV, At As Range)
PutSq SqzAyV(AyV), At
End Sub
Function CrtLo(Ws As Worksheet, FF$, Optional Lon$) As ListObject
Set CrtLo = LozRg(RgzAyH(SyzSS(FF), A1zWs(Ws)), Lon)
End Function

Sub PutSSH(SSH$, At As Range)
PutAyH SyzSS(SSH), At
End Sub

Sub PutSSV(SSV$, At As Range)
PutAyV SyzSS(SSV), At
End Sub

Sub PutAyH(AyH, At As Range)
PutSq SqzAyH(AyH), At
End Sub
Sub PutDbtAt(D As Database, T, At As Range, Optional AddgWay As EmAddgWay)
LozRg PutSq(SqzT(D, T), At), Lon(T)
End Sub
Sub SetQtFbt(QT As QueryTable, Fb, T)
With QT
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

Function WszT(D As Database, T, Optional Wsn$) As Worksheet
Dim Sq(): Sq = SqzT(D, T)
Dim A1 As Range: Set A1 = NewA1(Wsn)
Set WszT = WszLo(LozSq(Sq(), A1))
End Function


'

