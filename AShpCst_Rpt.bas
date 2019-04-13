Attribute VB_Name = "AShpCst_Rpt"
Option Explicit
Const CMod$ = "AShpCst_Rpt."
Function RptAppDb() As Database
'Set RptAppDb = AppDb(RptApn)
End Function

Sub GenOupTbl(Apn$)
WOpn Apn
GenORate
GenOMain
WCls
End Sub

Function OupFxzShpCst$()
OupFxzShpCst = OupFxzLidPm(RptLidPm)
End Function

Sub BrwRptPm()
ShpCstLiPm.Brw
End Sub

Sub DocUOM _
(A, _
B)
'InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
'Oup : UOM        Sku      SkuUOM                 Des                    Sc_U

'Note on [Sales text.xls]
'Col  Xls Title            FldName     Means
'F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
'J    Unit per case        Sc_U        how many unit per AC
'K    SC                   SC_U        how many unit per SC   ('no need)
'L    COL per case         AC_B        how many bottle per AC
'-----
'Letter meaning
'B = Bottle
'AC = act case
'SC = standard case
'U = Unit  (Bottle(COL) or Set (PCE))

' "SC              as SC_U," & _  no need
' "[COL per case]  as AC_B," & _ no need
End Sub

Private Sub GenOMain()
'Inp: #IMB52
'     #IUom
'     @Rate

WDrp "@Main"
WRun "Select Distinct Whs,Sku,Sum(QUnRes+QBlk+QInsp) As OH into [@Main] from [#IMB52] Group by Whs,Sku"
'Des StkUom Sc_U OH_Sc
WRun "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10),Sc_U Int, OH_Sc Double"
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"
WRun "Update [@Main] set OH_Sc=OH/Sc_U where Sc_U>0"

'Stream ProdH F2 M32 M35 M37 Topaz ZHT1 RateSc Z2 Z5 Z7
WRun "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH text(7), F2 Text(2), M32 text(2), M35 text(5), M37 text(7), ZHT1 Text(7), Z2 text(2), Z5 text(5), Z7 text(7), RateSc Currency, Amt Currency"

'ProdH Topaz
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"

'F2 M32 M35 M37
WRun "Update [@Main] set F2=Left(ProdH,2),M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M37=Mid(ProdH,3,7)"

'ZHT1 RateSc
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M37=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'Stream
WRun "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"

'Z2 Z5 Z7
WRun "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z7=Left(ZHT1,7) where not ZHT1 is null"

'Amt
WRun "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
End Sub

Private Sub GenORate()
'VdtFm & VdtTo format DD.MM.YYYY
'1: #IZHT1 VdtFm VdtTo L3 RateSc
'2: #IUom     SKu Sc_U
'O: @Rate  ZHT1 RateSc
WDrp "#Cpy1 #Cpy2 #Cpy @Rate"
WRun "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
WRun "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

WRun "Select * into [#Cpy] from [#Cpy1] where False"
WRun "Insert into [#Cpy] select * from [#Cpy1]"
WRun "Insert into [#Cpy] select * from [#Cpy2]"

WRun "Alter Table [#Cpy] Add Column VdtFmDte Date,VdtToDte Date,IsCur YesNo"
WRun "Update [#Cpy] Set" & _
" VdtFmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" VdtToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
WRun "Update [#Cpy] set IsCur = true where Now between VdtFmDte and VdtToDte"

WRun "Select Whs,ZHT1,RateSc into [@Rate] from [#Cpy]"
WDrp "#Cpy #Cpy1 #Cpy2"
End Sub

Private Function ErzMB52MissingWhs8601Or8701(FxMB52$, Wsn$) As String()
Const CSub$ = CMod & "ErzMB52MissingWhs8601Or8701"
Const M$ = "Column-[Plant] must have value 8601 or 8701"
If NReczFxw(FxMB52, Wsn, "Plant in ('8601','8701')") = 0 Then
    ErzMB52MissingWhs8601Or8701 = _
        LyzFunMsgNap(CSub, M, "MB52-File Worksheet", FxMB52, Wsn)
End If
End Function

Function PnmStkDte(AppDb As Database) As Date
PnmStkDte = CDate(Mid(PnmVal(AppDb, "MB52Fn"), 6, 10))
End Function

Function PnmStkYYMD$(AppDb As Database)
PnmStkYYMD = Format(PnmStkDte(AppDb), "YYYY-MM-DD")
End Function

Sub ShpCstBrwLiAct()
BrwLiAct ShpCstLiAct
End Sub

Property Get ShpCstLiAct() As LiAct
Set ShpCstLiAct = LiAct(ShpCstLiPm)
End Property

