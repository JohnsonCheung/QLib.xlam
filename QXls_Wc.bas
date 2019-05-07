Attribute VB_Name = "QXls_Wc"
Option Explicit
Private Const CMod$ = "MXls_Wc."
Private Const Asm$ = "QXls"
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
RgzWc Wc, A1zWs(Ws)
Set WszWc = Ws
End Function

Sub RgzWc(Wc As WorkbookConnection, At As Range)
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
    .ListObject.DisplayName = LoNm(Wc.Name)
    .Refresh BackgroundQuery:=False
End With
End Sub

Sub AddWczTT(ToWb As Workbook, FmFb$, TT$)
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
Sub CrtFxzOupTbl(Fx$, Fb$, Optional AddgWay As EmAddgWay)
SavAszAndCls NewWbzOupTbl(Fb, AddgWay), Fx
End Sub

