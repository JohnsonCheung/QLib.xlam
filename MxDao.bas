Attribute VB_Name = "MxDao"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxDao."

Function CvCn(A) As ADODB.Connection
Set CvCn = A
End Function

Sub Rpl_LoCn_ByFbt(Lo As ListObject, Fb, T)
With Lo.QueryTable
    Rpl_Wc_ByFb .Connection, Fb '<==
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

Sub Z_LoIxSq()
Dim Wb As Workbook: Set Wb = NewWb
AddWszzWbSq Wb, SampSq
AddWszzWbSq Wb, SampSq1
BrwSq LoIxSq(Wb)
End Sub

Sub AddWszzWbSq(Wb As Workbook, Sq())
CrtLoAtzSq Sq, A1zWs(AddWs(Wb))
End Sub

Function LoIxSq(Wb As Workbook) As Variant()
Dim Ws As Worksheet, M(), O(), Fnd As Boolean
For Each Ws In Wb.Sheets
    M = LoIxDtaSq(Ws)
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

Function LoIxDtaSq(Ws As Worksheet) As Variant()
Dim Lo As ListObject, R&, NR&
NR = Ws.ListObjects.Count
ReDim O(1 To NR, 1 To 4)
For Each Lo In Ws.ListObjects
    R = R + 1
    SetSqr O, LoDr(Lo), R
Next
End Function

Sub Z_LoDr()
Dim Lo As ListObject: Set Lo = SampLo
D LoDr(Lo)
ClsWbNoSav WbzLo(Lo)
End Sub

Function LoDr(A As ListObject) As Variant()
Dim WN$: WN = WsnzLo(A)
Dim LN$:: LN = A.Name
Dim NR&: NR = NRowzLo(A)
Dim NC&: NC = A.ListColumns.Count
LoDr = Array(WN, LN, NR, NC)
End Function

Sub Add_Lo_At_FmFbtIx(At As Range)
CrtLoAtzSq LoIxSq(WbzRg(At)), At
End Sub

Function Crt_Wb_FmDbtmpInp(D As Database) As Workbook
Set Crt_Wb_FmDbtmpInp = Crt_Wb_FmDbtny(D, TmpInpTny(D))
End Function

Function Crt_Wb_FmDbtny(D As Database, Tny$()) As Workbook
Dim T, O As Workbook
Set O = NewWb
For Each T In Itr(Tny)
    Add_Ws_ToWb_FmDbt O, D, CStr(T)
Next
DltSheet1 O
Set Crt_Wb_FmDbtny = O
End Function
Function WbzFb(Fb) As Workbook
Dim D As Database: Set D = Db(Fb)
Set WbzFb = ShwWb(Crt_Wb_FmDbtny(D, Tny(D)))
End Function

Function SetWsn(Ws As Worksheet, Nm$) As Worksheet
Set SetWsn = Ws
If Nm = "" Then Exit Function
If Ws.Name = Nm Then Exit Function
If HasWs(WbzWs(Ws), Nm) Then
    Dim Wb As Workbook: Set Wb = WbzWs(Ws)
    Thw CSub, "Wsn exists in Wb", "Wsn Wbn Wny-in-Wb", Nm, Wbn(Wb), Wny(Wb)
End If
Ws.Name = Nm
End Function
Sub Add_Lo_At_FmFbtzSamp()
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
Function Add_Lo_At_FmFbt(At As Range, Fb, T) As ListObject
Dim Ws As Worksheet: Set Ws = WszRg(At)
Dim Lo As ListObject: Set Lo = Ws.ListObjects.Add(xlSrcExternal, OleCnStrzFb(Fb), Destination:=At)
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
    .ListObject.DisplayName = LonzT(T)
    .Refresh BackgroundQuery:=False
End With
AutoFit_Lo Lo
End Function
Sub AutoFit_Lo(A As ListObject)
A.DataBodyRange.EntireColumn.AutoFit
End Sub

Function Add_Ws_ToWb_FmDbt(Wb As Workbook, Db As Database, T, Optional Wsn0$, Optional AddgWay As EmAddgWay) As Worksheet
Dim O As Worksheet: Set O = AddWs(Wb, StrDft(Wsn0, T))
Dim A1 As Range: Set A1 = A1zWs(O)
PutTbl Db, T, A1, AddgWay
End Function

Function Crt_Wb_FmFb_OupTbl(Fb, Optional AddgWay As EmAddgWay) As Workbook '
Dim O As Workbook, D As Database
Set O = NewWb
Set D = Db(Fb)
Add_Ws_ToWb_FmDbtny O, D, OupTny(D), AddgWay
DltWsIf O, "Sheet1"
Set Crt_Wb_FmFb_OupTbl = O
End Function

Function Crt_Wb_FmDbt(D As Database, T, Optional Wsn$ = "Data") As Workbook
Set Crt_Wb_FmDbt = WszRg(Add_Ws_ToWb_FmDbt(NewWb, D, T, Wsn))
End Function

Sub Put_Dbt_ToWs(D As Database, T, Ws As Worksheet)
Put_Dbt_At D, T, A1zWs(Ws)
End Sub
Sub ClrLo(A As ListObject)
If A.ListRows.Count = 0 Then Exit Sub
A.DataBodyRange.Delete xlShiftUp
End Sub
Sub Put_Dbt_At(D As Database, T, At As Range, Optional AddgWay As EmAddgWay)
CrtLo Put_Sq_At(SqzT(D, T), At), Lon(T)
End Sub
Sub SetQtFbt(Qt As QueryTable, Fb, T)
With Qt
    .CommandType = xlCmdTable
    .Connection = OleCnStrzFb(Fb) '<--- Fb
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
Sub PutFbtAt(Fb, T$, At As Range, Optional Lon0$)
Dim O As ListObject
Set O = WszRg(At).ListObjects.Add(SourceType:=XlSourceType.xlSourceWorkbook, Destination:=At)
SetLon O, Dft(Lon0, Lon(T))
SetQtFbt O.QueryTable, Fb, T
End Sub
Sub FxzTny(Fx, Db As Database, Tny$())
Crt_Wb_FmDbtny(Db, Tny).SaveAs Fx
End Sub

Function NewWs_FmDbt(D As Database, T, Optional Wsn$) As Worksheet
Dim Sq(): Sq = SqzT(D, T)
Dim A1 As Range: Set A1 = NewA1(Wsn)
Set NewWs_FmDbt = WszLo(CrtLoAtzSq(Sq(), A1))
End Function
