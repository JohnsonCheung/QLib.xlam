Attribute VB_Name = "MXls_Wc"
Option Explicit

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

Sub AddWcTpWFb()
'AddWcFxFbtt TpFx, WFb(Apn), TnyzFb(WFb)
End Sub

Sub AddWcFxFbtt(Fx, LnkFb$, TT)
Dim Wb As Workbook, T
Set Wb = WbzFx(Fx)
For Each T In NyzNN(TT)
    WczWbFb Wb, LnkFb, T
Next
Wb.Close True
End Sub

Private Function WbzFbOupTbl(Fb) As Workbook
Dim O As Workbook
Set O = NewWb
DoAyABX OupTnyzFb(Fb), "WczWbFb", O, Fb
DoItrFun O.Connections, "NewWsC"
RfhWb O, Fb
Set WbzFbOupTbl = O
End Function

Sub CrtFxzFbOupTbl(Fx$, Fb$)
WbzFbOupTbl(Fb).SaveAs Fx
End Sub

