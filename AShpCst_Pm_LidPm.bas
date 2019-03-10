Attribute VB_Name = "AShpCst_Pm_LidPm"
Option Explicit
Property Get RptAppFb$()
RptAppFb = "C:\Users\user\Documents\Projects\Vba\ShpCst\ShpCstApp.accdb"
End Property

Private Sub Z_RptLidPrmSrc()
Brw FmtAyPm(RptLidPmSrc, "[LidPm Apn AppFb Fil Ws WsCol] [2:LidPm AppFb Fil] [3:Ws] 4:WsCol")
End Sub
Property Get RptLidPmSrc() As String()
RptLidPmSrc = RptLidPmSrczAppFb(RptAppFb)
End Property

Private Function RptLidPmSrczAppFb(AppFb$) As String()
Erase XX
X "LidPm"
X "Apn ShpCst"
X "AppFb " & AppFb
Dim A$(): A = LidFilLinAy(AppFb)
X A(0)
X A(1)
X A(2)
'X "Fil MB52 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\MB52 2018-07-30.xls"
'X "Fil UOM  C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\sales text.xlsx"
'X "Fil ZHT1 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\ZHT1.XLSX"
X "Ws ZHT18701  ZHT1 8601"
X "Ws ZHT18601  ZHT1 8701"
X "Ws UOM       UOM  Sheet1"
X "Ws MB52      MB52 Sheet1"
X "WsCol ZHT18701 ZHT1   M Brand"
X "WsCol ZHT18701 RateSc D Amount"
X "WsCol ZHT18701 VdtFm  M Valid From"
X "WsCol ZHT18701 VdtTo  M Valid to"
X "WsCol ZHT18601 ZHT1   M Brand"
X "WsCol ZHT18601 RateSc D Amount"
X "WsCol ZHT18601 VdtFm  M Valid From"
X "WsCol ZHT18601 VdtTo  M Valid to"
X "WsCol UOM Sku     M Material"
X "WsCol UOM Des     M Material Description"
X "WsCol UOM Sc_U    M SC "
X "WsCol UOM StkUom  M Base Unit of Measure"
X "WsCol UOM Topaz   M Topaz Code"
X "WsCol UOM ProdH   M Product hierarchy"
X "WsCol MB52 Sku    M Material"
X "WsCol MB52 Whs    M Plant"
X "WsCol MB52 QInsp  D In Quality Insp#"
X "WsCol MB52 QUnRes D Unrestricted"
X "WsCol MB52 QBlk   D Blocked"
RptLidPmSrczAppFb = XX
Erase XX
End Function

Property Get RptLidPm() As LidPm
Set RptLidPm = RptLidPmzAppFb(RptAppFb)
End Property

Private Function RptLidPmzAppFb(AppFb$) As LidPm
Set RptLidPmzAppFb = LidPm(RptLidPmSrczAppFb(AppFb))
End Function

Private Function RptLidFilLinAy() As String()
RptLidFilLinAy = LidFilLinAy(SampFbzShpCst)
End Function

Private Function LidFilLinAy(AppFb$) As String()
Dim D As Database: Set D = Db(AppFb)
Dim O$(2)
O(0) = LidFilLin(D, "MB52")
O(1) = LidFilLin(D, "UOM")
O(2) = LidFilLin(D, "ZHT1")
LidFilLinAy = O
D.Close
End Function

Private Function LidFilLin$(AppDb As Database, Itm$)
Dim Ffn$: Ffn = PnmFfn(AppDb, Itm)
LidFilLin = FmtQQ("Fil ? ?", Itm, Ffn)
End Function

