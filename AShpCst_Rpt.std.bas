Attribute VB_Name = "AShpCst_Rpt"
Option Explicit
Public Const ShpCstApn$ = "ShpCst"
Sub ShpCstGenSamp()
End Sub

Sub GenRpt()
Dim P As LidPm: Set P = RptLidPm
Set CDb = LnkImpDbzLidPm(P)
GenORate
GenOMain
Dim OupFx$: OupFx = TmpFx
GenOupFx P.Apn, OupFx
End Sub

Sub ShpCstBrwPm()
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

Drp "@Main"
RunQ "Select Distinct Whs,Sku,Sum(QUnRes+QBlk+QInsp) As OH into [@Main] from [#IMB52] Group by Whs,Sku"

'Des StkUom Sc_U OH_Sc
RunQ "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10),Sc_U Int, OH_Sc Double"
RunQ "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"
RunQ "Update [@Main] set OH_Sc=OH/Sc_U where Sc_U>0"

'Stream ProdH F2 M32 M35 M37 Topaz ZHT1 RateSc Z2 Z5 Z7
RunQ "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH text(7), F2 Text(2), M32 text(2), M35 text(5), M37 text(7), ZHT1 Text(7), Z2 text(2), Z5 text(5), Z7 text(7), RateSc Currency, Amt Currency"

'ProdH Topaz
RunQ "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"

'F2 M32 M35 M37
RunQ "Update [@Main] set F2=Left(ProdH,2),M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M37=Mid(ProdH,3,7)"

'ZHT1 RateSc
RunQ "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M37=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
RunQ "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
RunQ "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'Stream
RunQ "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"

'Z2 Z5 Z7
RunQ "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z7=Left(ZHT1,7) where not ZHT1 is null"

'Amt
RunQ "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
End Sub

Private Sub GenORate()
'VdtFm & VdtTo format DD.MM.YYYY
'1: #IZHT1 VdtFm VdtTo L3 RateSc
'2: #IUom     SKu Sc_U
'O: @Rate  ZHT1 RateSc
Drp "#Cpy1 #Cpy2 #Cpy @Rate"
RunQ "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
RunQ "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

RunQ "Select * into [#Cpy] from [#Cpy1] where False"
RunQ "Insert into [#Cpy] select * from [#Cpy1]"
RunQ "Insert into [#Cpy] select * from [#Cpy2]"

RunQ "Alter Table [#Cpy] Add Column VdtFmDte Date,VdtToDte Date,IsCur YesNo"
RunQ "Update [#Cpy] Set" & _
" VdtFmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" VdtToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
RunQ "Update [#Cpy] set IsCur = true where Now between VdtFmDte and VdtToDte"

RunQ "Select Whs,ZHT1,RateSc into [@Rate] from [#Cpy]"
Drp "#Cpy #Cpy1 #Cpy2"
End Sub


Private Property Get MB52_8601_8701_Missing() As String()
'Const CSub$ = CMod & "MB52_8601_8701_Missing"
'Const M$ = "Column-[Plant] must have value 8601 or 8701"
'Const Wh$ = "Plant in ('8601','8701')"
'Dim Fx$
'Fx = IFxMB52
''LnkFxDb W, "#A", Fx
''If NRecDT(W, "#A", Wh) = 0 Then
'    MB52_8601_8701_Missing = _
'        LyzFunMsgNap(CSub, M, "MB52-File Worksheet", Fx, "Sheet1")
''End If
'WDrp "#A"
End Property

Sub OpnRptMB52(): OpnFx RptIFilMB52: End Sub
Sub OpnRptUOM():  OpnFx RptIFilUOM: End Sub
Sub OpnRptZHT1(): OpnFx RptIFilZHT1: End Sub

Property Get RptIFilMB52$()
RptIFilMB52 = RptWPth & PnmFn("MB52")
End Property

Property Get RptWPth$()
RptWPth = WPth(ShpCstApn)
End Property

Property Get RptIFilUOM$()
RptIFilUOM = RptWPth & PnmFn("UOM")
End Property

Property Get RptIFilZHT1$()
RptIFilZHT1 = RptWPth & PnmFn("ZHT1")
End Property

Property Get PnmStkDte() As Date
PnmStkDte = CDate(Mid(PnmVal("MB52Fn"), 6, 10))
End Property

Property Get PnmStkYYMD$()
PnmStkYYMD = Format(PnmStkDte, "YYYY-MM-DD")
End Property

Sub ShpCstBrwLiAct()
BrwLiAct ShpCstLiAct
End Sub

Property Get ShpCstLiAct() As LiAct
Set ShpCstLiAct = LiAct(ShpCstLiPm)
End Property


