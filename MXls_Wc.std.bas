Attribute VB_Name = "MXls_Wc"
Option Explicit
Function NewWsWc(Wc As WorkbookConnection) As Worksheet
Dim Wb As Workbook, Ws As Worksheet
Set Wb = Wc.Parent
Set Ws = AddWs(Wb, Wc.Name)
PutWcAt Wc, A1zWs(Ws)
Set NewWsWc = Ws
End Function

Sub PutWcAt(A As WorkbookConnection, At As Range)
Dim Lo As ListObject
Set Lo = WszRg(At).ListObjects.Add(SourceType:=0, Source:=A.OLEDBConnection.Connection, Destination:=At)
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = A.Name
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
    .ListObject.DisplayName = LoNm(A.Name)
    .Refresh BackgroundQuery:=False
End With
End Sub

Sub AddWcTpWFb()
'AddWcFxFbtt TpFx, WFb(Apn), TnyzFb(WFb)
End Sub

Sub AddWcFxFbtt(Fx, LnkFb$, TT)
Dim Wb As Workbook, T
Set Wb = WbzFx(Fx)
For Each T In FnyzFF(TT)
    AddWczWbFb Wb, LnkFb, T
Next
Wb.Close True
End Sub

Function WbzFbOupTbl(Fb) As Workbook
Dim O As Workbook
Set O = NewWb
DoAyABX OupTnyzFb(Fb), "AddWczWbFb", O, Fb
DoItrFun O.Connections, "NewWsC"
RfhWb O, Fb
Set WbzFbOupTbl = O
End Function

Sub RplLoCn(Wb As Workbook, Fb)
Dim I, Lo As ListObject, D As Database
Set D = Db(Fb)
For Each I In OupLoAy(Wb)
    Set Lo = I
    RplLoCnzDbt Lo, D, "@" & Mid(Lo.Name, 3)
Next
D.Close
Set D = Nothing
End Sub

Sub CrtFxzFbOupTbl(Fb$, Fx$)
WbzFbOupTbl(Fb).SaveAs Fx
End Sub

