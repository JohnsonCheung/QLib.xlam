Attribute VB_Name = "ATaxExpCmp"
Option Explicit
Private Type Pm
 OupPth As String: OupFxFn As String: InpPth As String: InpFxFnGLDutyFx As String: InpFxFnGLAnp As String: InpFxFnMB51 As String
 Tolerence As Long
End Type
Type XlsLnkInf
    IsXlsLnk As Boolean
    Fx As String
    WsNm As String
End Type
Type Fxw
    Fx As String
    Wsn As String
End Type
Private Const OupTn_Main$ = "@Main"
Const OupFld_Main$ = _
    "Flg RecTy Amt Key Uom MovTy Qty BchRateUX RateTy Bch Las GL |" & _
    " Flg IsAlert Is3p IsRepack Lvl1 Lvl2 Lvl3 Rank |" & _
    " Key PstMth PstDte Sku |" & _
    " Bch BchRateU BchNo    BchPermitDate BchPermit BchPermitD|" & _
    " Las LasRateU LasBchNo LasPermitDate LasPermit LasPermitD|" & _
    " GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
    " Uom Des StkUom Ac_U"
Const TblFld_PermitD$ = "Duty.PermitD "
Const TblFld_Permit$ = "Duty.Permit  "
Const TblFld_Repack$ = "StockHld.SkuRepackMuti SkuNew SkuFm FmQty "
Const TblFld_SkuTaxBy3rdParty$ = "TaxCmp.SkuTaxBy3rdParty      Sku RateU"
Const TblFld_SkuNoLongerTax$ = "TaxCmp.SkuNoLongerTax        Sku"
Const WsFld_GLDuty$ = "#IGLDuty Sku PstDte Amt GLDocNo GLDocDte GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef"
Const WsFld_GLAnp$ = "#IGLAnp  Sku PstDte Amt GLDocNo GLDocDte GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef"
Const WsFld_MB51$ = "#IMB51   Sku PstDte MovTy Qty BchNo"
Const WsFld_UOM$ = "#IUom    Sku Des Ac_U StkUom"
Public Const LnkSpec_GLDuty$ = ">GLDuty |" & _
    "Sku       Txt Material |" & _
    "PstDte    Dte [Posting Date] |" & _
    "Amt       Dbl [Amount in local currency] |" & _
    "GLDocNo   Txt [Document Number] |" & _
    "GLDocDte  Dte [Document Date] |" & _
    "GLAsg     Txt Assignment |" & _
    "GLDocTy   Txt [Document Type] |" & _
    "GLLin     Txt [Line item] |" & _
    "GLPstKy   Txt [Posting Key] |" & _
    "GLPc      Txt [Profit Center] |" & _
    "GLAc      Txt [G/L Account] |" & _
    "GLAcTy    Txt [Account Type] |" & _
    "GLBusA    Txt [Business Area]|" & _
    "GLRef     Txt [Reference] |" & _
    "Where not [Posting Date] is null"

Public Const LnkSpec_GLAnp$ = ">GLAnp |" & _
    "Sku       Txt Material |" & _
    "PstDte    Dte [Posting Date] |" & _
    "Amt       Dbl [Amount in local currency] |" & _
    "GLDocNo   Txt [Document Number] |" & _
    "GLDocDte  Dte [Document Date] |" & _
    "GLAsg     Txt Assignment |" & _
    "GLDocTy   Txt [Document Type] |" & _
    "GLLin     Txt [Line item] |" & _
    "GLPstKy   Txt [Posting Key] |" & _
    "GLPc      Txt [Profit Center] |" & _
    "GLAc      Txt [G/L Account] |" & _
    "GLAcTy    Txt [Account Type] |" & _
    "GLBusA    Txt [Business Area]|" & _
    "GLRef     Txt [Reference] |" & _
    "Where not [Posting Date] is null"

Const LnkSpec_MB51$ = ">MB51 |" & _
    "Whs    Txt Plant |" & _
    "Loc    Txt [Storage Location]|" & _
    "Sku    Txt Material |" & _
    "PstDte Txt [Posting Date] |" & _
    "MovTy  Txt [Movement Type]|" & _
    "Qty    Txt Quantity|" & _
    "BchNo  Txt Batch |" & _
    "Where (Plant='8601' and [Storage Location]='0002' and [Movement Type] like '6*' and not [Movement Type] in ('632','633'))" & _
    " or   (Plant='8601' and [Movement Type] in ('633','634'))"
Const LnkSpec_Uom$ = ">Uom |" & _
    "Sku    Txt Material |" & _
    "Des    Txt [Material Description] |" & _
    "Ac_U   Txt [Unit per case] |" & _
    "StkUom Txt [Base Unit of Measure] |" & _
    "ProdH  Txt [Product hierarchy] |" & _
    "Where Plant='8601'"

'Material    Amount  Document Number Posting Date  Document Date   Assignment  Document Type   Line item   Posting Key Profit Center   G/L Account Business Area   Reference   Account Type
'1074574 -3684.17    800000004       2/1/2018      2/1/2018        20180102    CA              4           50          86BL0A4GRM      E353101C    UD00            0802023295  S
'1051982 -2107.20    800000007       2/1/2018      2/1/2018        20180102    CA              4           50          86JB020GRM      E353101C    UD00            0802023298  S
'1055642 -1774.92    800000007       2/1/2018      2/1/2018        20180102    CA              11          50          86HY040GRM      E353101C    HY00            0802023298  S
Const Apn$ = "TaxExpCmp"
Private W As Database
Property Get ATaxExpCmp() As RptPm
Static X As Boolean, XGenr As IGenr, XWbFmtr As IWbFmtr
With ATaxExpCmp
Set .Genr = XGenr
'.ImpWsSqy = EmpSy
'.InpFbAy = EmpSy
'.InpFbTny = SySsl("")
End With
End Property
Property Get App() As App
Static X As App
If IsNothing(X) Then
    Set X = New App
    X.Init "TaxExpCmp", "1_3"
End If
Set App = X
End Property

Private Function DSpecNm$(DSpec$)
DSpecNm = AftDotOrAll(T1(DSpec))
End Function


Private Sub WMB51Opn()
OpnFx IFx_WMB51
End Sub
Private Function LnkSpec_Ay() As String()
LnkSpec_Ay = Sy(LnkSpec_GLDuty, LnkSpec_GLAnp, LnkSpec_MB51, LnkSpec_Uom)
End Function
Private Function IMB51Fny() As String()
IMB51Fny = Fny(W, ">MB51")
End Function
Private Function IGLAnpFny() As String()
IGLAnpFny = Fny(W, ">GLAnp")
End Function
Private Function IGLDutyFny() As String()
IGLDutyFny = Fny(W, ">MB51")
End Function
Private Sub MsgSet(A$)
'Form_Main.MsgSet A
End Sub
Private Sub MsgClr()
'Form_Main.MsgClr
End Sub
Private Function DSpecAy() As String()
DSpecAy = Sy(TblFld_PermitD, TblFld_Permit, TblFld_Repack, TblFld_SkuTaxBy3rdParty, TblFld_SkuNoLongerTax)
End Function
Sub ImpzLnkSpec_Ay(A As Database, LnkSpec_Ay$())

End Sub
Private Sub Rpt_Imp()
MsgSet "Import the Excel files ....."
ImpzLnkSpec_Ay W, LnkSpec_Ay
ImpTT W, "Permit PermitD SkuRepackMulti SkuTaxBy3rdParty SkuNoLongerTax"
W.Execute FmtQQ("Alter Table [?] drop column Whs,Loc", TmpInpTn_MB51)
End Sub
Private Sub ImpTT(A As Database, TT)
'ImpTT W, TT
End Sub

Private Sub Rpt_Tmp()
MsgSet "Running query ($Sku) .....": Tmp_Sku
MsgSet "Running query ($MB51) .....": Tmp_MB51
MsgSet "Running query ($Rate) .....": Tmp_Rate
MsgSet "Running query ($T1T2) .....": Tmp_T1T2
MsgSet "Running query ($T3 & $T5) .....": Tmp_T3_and_5RepackOup
MsgSet "Running query ($T4) .....": Tmp_T4
End Sub

Private Sub Tmp_Rate()
WDrp "$Rate #A"
W.Execute "select Distinct Sku into [#A] from [$MB51]"

W.Execute "select Sku,PermitDate,PermitDate as PermitDateEnd,BchNo,Rate as BchRateU,PermitD,x.Permit" & _
" into [$Rate] from [#IPermitD] x" & _
" inner join [#IPermit] a on x.Permit=a.Permit" & _
" where Sku in (Select Distinct Sku from [#A])" & _
" Order By Sku,PermitDate"
UpdEndDte W, "$Rate", "PermitDateEnd", "Sku", "PermitDate"
W.Execute "Create Index Sku on [$Rate] (Sku,BchNo)"

WDrp "#A"

End Sub
Private Sub Tmp_T4()
WDrp "$T4 #A #B"
W.Execute "Select * into [#A] from [#ISkuTaxBy3rdParty]"
W.Execute "Select Id,Sku,Is3p into [#B] from [$MB51] where Is3p"
W.Execute "Select Id,x.Sku,RateU into [$T4] from [#B] x left join [#A] a on x.Sku=a.Sku"
End Sub
Private Sub Tmp_T3_and_5RepackOup()
WDrp "$T3 #A #B #C #D #E #F @Repack1 @Repack2 @Repack3 @Repack4 @Repack5 @Repack6"
W.Execute "Select * into [#A] from [#ISkuRepackMulti]"
W.Execute "Create Index Pk on [#A] (SkuNew,SkuFm) with Primary"

W.Execute "Select Id,Sku as MB51SkuNew, PstDte into [#B] from [$MB51] where IsRepack"
W.Execute "Create Index Pk on [#B] (Id) with Primary"

W.Execute "Select Id,MB51SkuNew,PstDte,SkuFm,FmSkuQty into [#C] from [#A] x inner join [#B] a on x.SkuNew=a.MB51SkuNew"
W.Execute "Create Index Pk on [#C] (Id,SkuFm)"

W.Execute "Select Distinct MB51SkuNew,PstDte,SkuFm,FmSkuQty into [#D] from [#C]"
W.Execute "Create Index Pk on [#D] (MB51SkuNew,PstDte,SkuFm)"

W.Execute "Select MB51SkuNew,PstDte,SkuFm,FmSkuQty,BchRateU,PermitDate,BchNo,Permit,CLng(a.PermitD) as PermitD" & _
    " into [#E] from [#D] x,[$Rate] a where false"
W.Execute "Insert into [#E] select * from [#D]"
W.Execute "Update [#E] x inner join [$Rate] a on x.SkuFm=a.Sku set " & _
" x.BchRateU  =a.BchRateU  ," & _
" x.PermitDate=a.PermitDate," & _
" x.BchNo     =a.BchNo     ," & _
" x.Permit    =a.Permit    ," & _
" x.PermitD   =a.PermitD    " & _
" where PstDte between a.PermitDate and a.PermitDateEnd"
W.Execute "Alter Table [#E] add column Amt Currency"
W.Execute "Update [#E] set Amt=FmSkuQty*BchRateU"
W.Execute "Create Index Pk on [#E] (MB51SkuNew,PstDte,SkuFm)"

W.Execute "Select Distinct MB51SkuNew,PstDte,Sum(x.Amt) as PackRateU, Count(*) as FmSkuCnt,Sum(x.FmSkuQty) as FmQty into [#F] from [#E] x group by MB51SkuNew,PstDte"
W.Execute "Create Index Pk on [#F] (MB51SkuNew,PstDte)"

W.Execute "Select Id,x.MB51SkuNew,x.PstDte,PackRateU,FmSkuCnt,FmQty into [$T3] from [#B] x inner join [#F] a on x.MB51SkuNew=a.MB51SkuNew and x.PstDte=a.PstDte"
W.Execute "Create Index Pk on [$T3] (Id)"

W.Execute "Drop Index Pk on [#B]"
W.Execute "Drop Index Pk on [#C]"

W.Execute "ALter Table [#B] drop column Id"
W.Execute "ALter Table [#C] drop column Id"

WReOpn
WRenTbl "#A", "@Repack1"
WRenTbl "#B", "@Repack2"
WRenTbl "#C", "@Repack3"
WRenTbl "#D", "@Repack4"
WRenTbl "#E", "@Repack5"
WRenTbl "#F", "@Repack6"

End Sub
Private Sub WRenTbl(Fm, ToTbl)

End Sub
Private Sub WReOpn()

End Sub

Private Sub Tmp_T1T2()
WDrp "#M ##M ##R #A #B $T1 $T2"

'M
W.Execute "select Id,Sku,BchNo,PstDte into [#M] from [$MB51]"
W.Execute "Update [#M] set BchNo='' where BchNo is null"

'T1
W.Execute "Select Id,x.Sku,x.BchNo,PermitD,Permit,PermitDate,BchRateU into [#A] from [#M] x inner join [$Rate] a on x.Sku=a.Sku and x.BchNo=a.BchNo"
W.Execute "Select Distinct Id,Sku,BchNo,Max(x.PermitD) as PermitD into [#B] from [#A] x group by Id,Sku,BchNo"
W.Execute "Select x.Id,x.Sku,x.BchNo,x.PermitD,Permit,PermitDate,BchRateU into [$T1] from [#B] x  inner join [#A] a on x.Id=a.Id and x.Sku=a.Sku and x.BchNo=a.BchNo and x.PermitD=a.PermitD"
WDrp "#A #B"

'T2
W.Execute "Select x.* into [##M] from [#M] x left join [$T1] a on x.Id=a.Id where a.Id is null"
W.Execute "Select x.* into [##R] from [$Rate] x where Sku in (Select Sku from [##M])"
W.Execute "select Id,x.Sku,PstDte,PermitD,Permit,PermitDate,x.BchNo,BchRateU,a.BchNo As LasBchNo into [$T2] from [##M] x inner join [##R] a on x.Sku=a.Sku where PstDte between PermitDate and PermitDateEnd"

WDrp "#M ##M ##R #A #B"
End Sub
Private Sub Oup_ORate()
WDrp "@Rate"
'Add RecTy to $Rate
W.Execute "Select * into [@Rate] from [$Rate]"
W.Execute "Alter Table [@Rate] add column RecTy Text(4)"
W.Execute "Update [@Rate] x inner join [$T1] a on x.PermitD=a.PermitD set RecTy='*Bch'"
W.Execute "Update [@Rate] x inner join [$T2] a on x.PermitD=a.PermitD set RecTy='*Las'"

End Sub
Private Sub Tmp_MB51()
WDrp "$MB51"
W.Execute "Select CLng(0) as Id,IsTax,IsRepack,Is3p,IsNoLongerTax,x.* into [$MB51] from [#IMB51] x left join [$Sku] a on x.Sku=a.Sku where IsImportFmMB51 or a.Sku is null"
'UpdSeqFld W, "$MB51", "Id",
End Sub

Private Sub Rpt_Lnk()
MsgSet "Linking the Excel files ....."
Dim A$(), B$(), C$(), D$(), O$(), E$(), F$()
A = LnkFxw(W, ">MB51", IFx_WMB51)
B = LnkFxw(W, ">GLDuty", IFx_WGLDuty)
B = LnkFxw(W, ">GLAnp", IFx_WGLAnp)
C = LnkFxw(W, ">Uom", IFx_WUom)
D = LnkFbtt(W, "Permit PermitD", IFb_Duty)
E = LnkFbtt(W, "SkuRepackMulti SkuNoLongerTax SkuTaxBy3rdParty", IFb_StkHld)
O = AyAddAp(A, B, C, D, E)
If Si(O) > 0 Then Thw CSub, "There are error in linking tables", "Er", O
End Sub
Private Function XFfn$(PmNm$)
XFfn = FfnzPm(App.Db, PmNm)
End Function
Private Function IFb_Duty$()
IFb_Duty = XFfn("Duty")
Exit Function
'Const N$ = "N:\SAPAccessReports\DutyPrepay5\DutyPrepay5_data.mdb"
'If IsDev Then
'    Dim L$
'    L = CurDbPth & "PgmObj\Sample\TaxExpCmp_InpTbl.mdb"
'    IFb_Duty = L
'Else
'    IFb_Duty = N
'End If
End Function

Private Function IFb_StkHld$()
IFb_StkHld = XFfn("StkHld")
Exit Function
Const N$ = "N:\SAPAccessReports\StockHolding6\StockHolding6_Data.mdb"
If IsDev Then
    Dim L$
'    L = CurDbPth & "PgmObj\Sample\TaxExpCmp_InpTbl.mdb"
    IFb_StkHld = L
Else
    IFb_StkHld = N
End If
End Function
Private Function WPth$()
WPth = App.WPth
End Function
Private Sub WOpn()

End Sub
Private Sub Rpt()
Dim Tp$, AppNm$, AppDb As Database, OupFx$, OupWb As Workbook, WFb$, Pm As Pm
ClrMainMsg
Rpt_Cpy
Rpt_Lnk
Rpt_Imp
Rpt_Tmp
Rpt_Oup Pm.Tolerence
Rpt_Gen OupFx, AppDb, AppNm, Tp, OupWb, WFb, Pm
End Sub

Private Sub Rpt_Cpy()
SetMainMsg "Copying 4 Excel files to C: temp folder ...."
CpyFfnSyzIfDif Sy(IFx_MB51, IFx_Duty, IFx_Anp, IFx_Uom), WPth
End Sub

Private Function Gen_OupPth$()

End Function

Private Function Gen_OupFx$()
Dim A$, B$
A = Gen_OupPth & FmtQQ("TaxExpCmp ?.xlsx", Format(Now, "YYYY-MM-DD HHMM"))
Gen_OupFx = A
End Function

Private Sub Gen_Crt(OupFx$, AppDb As Database, AppNm$, Tp$)
ExpAtt AppDb, AppNm, Tp
CpyFfnzIfDif Tp, OupFx, True
End Sub
Private Sub Rpt_Gen(OupFx$, AppDb As Database, AppNm$, Tp$, OupWb As Workbook, WFb$, Pm As Pm)
SetMainMsg "Export to Excel ....."
Gen_Crt OupFx, AppDb, AppNm, Tp
Gen_Rfh OupWb, WFb
Gen_Fmt OupWb, Pm
End Sub
Private Sub Gen_Rfh(OupWb As Workbook, WFb$)
RfhWb OupWb, WFb
End Sub
Private Sub Gen_AddWs_FmInp_ToOupWb_ForMB51(OupWb As Workbook, MB51Ws As Worksheet)
MB51Ws.Copy , LasWs(OupWb)
LasWs(OupWb).Name = "Inp MB51"
End Sub
Private Sub Fmt_OupPm(OupWb As Workbook, Pm As Pm)
With Pm
RgzAyV Array(.OupPth, .OupFxFn, .InpPth, .InpFxFnGLDutyFx, .InpFxFnGLAnp, .InpFxFnMB51), WszCdNm(OupWb, "WsCtl").Range("C3")
End With
End Sub
Private Sub Gen_Fmt(OupWb As Workbook, Pm As Pm)
Fmt_OupPm OupWb, Pm
End Sub

Private Sub IOpnFxMB51()
OpnFx IFx_MB51
End Sub

Private Sub IOpnFbStkHld()
BrwFb IFb_StkHld
End Sub
Private Sub IOpnFbDuty()
BrwFb IFb_Duty
End Sub
Private Sub IOpnFxAnp()
OpnFx IFx_Anp
End Sub
Private Sub IOpnFxUom()
OpnFx IFx_Uom
End Sub

Private Sub IOpnFxDuty()
OpnFx IFx_Duty
End Sub

Private Function IFx_MB51$()
IFx_MB51 = App.FfnzPm("MB51")
End Function

Private Function IFx_Duty$()
IFx_Duty = App.FfnzPm("GLDuty")
End Function
Private Function IFx_Anp$()
IFx_Anp = App.FfnzPm("GLAnp")
End Function
Private Function IFx_Uom$()
IFx_Uom = App.FfnzPm("Uom")
End Function

Private Function IFx_WGLDuty$()
IFx_WGLDuty = WPth & Fn(IFx_Duty)
End Function
Private Function IFx_WGLAnp$()
IFx_WGLAnp = WPth & Fn(IFx_Anp)
End Function
Private Function IFx_WUom$()
IFx_WUom = WPth & Fn(IFx_Uom)
End Function

Private Function IFx_WMB51$()
IFx_WMB51 = WPth & Fn(IFx_MB51)
End Function

Private Sub Rpt_Oup(Tolerence&)
SetMainMsgzQnm "@Main": Oup_OMain Tolerence
SetMainMsgzQnm "@Rate": Oup_ORate
SetMainMsgzQnm "@Sku":  Oup_OSku
End Sub
Private Sub Oup_OSku()
WDrp "@Sku"
W.Execute "Select * into [@Sku] from [$Sku]"
End Sub
Private Sub Tmp_Sku()
WDrp "$Sku"
WDrp "#A #B #C #D ##"
W.Execute "Select Distinct Sku           into [#A] from [#IPermitD]"
W.Execute "Select Distinct SkuNew As Sku into [#B] from [#ISkuRepackMulti]"
W.Execute "Select Distinct Sku           into [#C] from [#ISkuTaxBy3rdParty]"
W.Execute "Select Distinct Sku           into [#D] from [#ISkuNoLongerTax]"
W.Execute "Select Sku,Des,StkUom,Ac_U into [##] from [#IUom]"
W.Execute "Alter Table [##] add column IsTax YesNo, IsRepack YesNo, Is3p YesNo, IsNoLongerTax YesNo, IsImportFmMB51 YesNo"
W.Execute "Update [##] x inner join [#A] a on x.Sku = a.Sku set IsTax=True"
W.Execute "Update [##] x inner join [#B] a on x.Sku = a.Sku set IsRepack=True"
W.Execute "Update [##] x inner join [#C] a on x.Sku = a.Sku set Is3p=True"
W.Execute "Update [##] x inner join [#D] a on x.Sku = a.Sku set IsNoLongerTax=True"
W.Execute "Update [##] set IsImportFmMB51=(IsTax or Is3p or IsRepack) and Not IsNoLongerTax"
W.Execute "Select * into [$Sku] from [##]"
W.Execute "Create Index Pk on [$Sku] (Sku) with Primary"
WDrp "#A #B #C #D ##"
End Sub
Private Sub Oup_OMain(Tolerence&)
OMain_1_Crt_WithTmpT1234
OMain_2_AddGL
OMain_3_AddCol_IsAlert Tolerence
OMain_4_MdyCol_PstDte
OMain_5_AddCol_PstMth
OMain_6_AddCol_Uom
OMain_7_AddCol_Rank
OMain_8_ReSeqFld
End Sub
Private Sub OMain_1_Crt_WithTmpT1234()
WDrp "@Main"

W.Execute "Select * into [@Main] from [$MB51]"
'----------------------------------
'Use $T1 $T2 $T3 $T4
'to update the addition columns @Main
W.Execute "Alter table [@Main] add column " & _
    "BchRateUX Currency, RateTy Text(4)," & _
    "LasRateU Currency, LasPermitD Long, LasPermit Long, LasPermitDate Date, LasBchNo text(50)," & _
    "BchRateU Currency, BchPermitD Long, BchPermit Long, BchPermitDate Date," & _
    "PackRateU Currency, PackFmSkuCnt Int, PackFmQty Int," & _
    "TaxBy3pRateU Currency"

Const S1$ = "set RateTy='*Bch' ,x.BchRateUX=a.BchRateU ,x.BchRateU=a.BchRateU,x.BchPermitD=a.PermitD,x.BchPermit=a.Permit, x.BchPermitDate=a.PermitDate"
Const S2$ = "set RateTy='*Las' ,x.BchRateUX=a.BchRateU ,x.LasRateU=a.BchRateU,x.LasPermitD=a.PermitD,x.LasPermit=a.Permit, x.LasPermitDate=a.PermitDate,x.LasBchNo=a.LasBchNo"
Const s3$ = "set RateTy='*Pac' ,x.BchRateUX=a.PackRateU,x.PackRateU=a.PackRateU,x.PackFmSkuCnt=a.FmSkuCnt,x.PackFmQty=a.FmQty"
Const s4$ = "set RateTy='*3p'  ,x.BchRateUX=a.RateU    ,x.TaxBy3pRateU=a.RateU"
W.Execute "Update [@Main] x inner join [$T1] a on x.Id=a.Id " & S1
W.Execute "Update [@Main] x inner join [$T2] a on x.Id=a.Id " & S2
W.Execute "Update [@Main] x inner join [$T3] a on x.Id=a.Id " & s3
W.Execute "Update [@Main] x inner join [$T4] a on x.Id=a.Id " & s4
W.Execute "Alter Table [@Main] drop column Id"
W.Execute "Alter Table [@Main] add column RecTy Text(8), Amt Currency"
W.Execute "Update [@Main] set RecTy='*MB51', Amt = BchRateUX*Qty"
End Sub
Private Sub OMain_2_AddGL()
W.Execute "Alter Table [@Main] add column" & _
" GLDocNo  Text(255)," & _
" GLDocDte Text(255)," & _
" GLAsg    Text(255)," & _
" GLDocTy  Text(255)," & _
" GLLin    Text(255)," & _
" GLPstKy  Text(255)," & _
" GLPc     Text(255)," & _
" GLAc     Text(255)," & _
" GLBusA   Text(255)," & _
" GLRef    Text(255)," & _
" GLAcTy   Text(255)"
W.Execute _
"Insert into [@Main] (Amt,Sku,PstDte,GLDocNo,GLDocDte,GLAsg,GLDocTy,GLLin,GLPstKy,GLPc,GLAc,GLBusA,GLRef,GLAcTy,RecTy)" & _
         " select     Amt,Sku,PstDte,GLDocNo,GLDocDte,GLAsg,GLDocTy,GLLin,GLPstKy,GLPc,GLAc,GLBusA,GLRef,GLAcTy,'*GLDuty' as RecTy from [#IGLDuty]"
W.Execute _
"Insert into [@Main] (Amt,Sku,PstDte,GLDocNo,GLDocDte,GLAsg,GLDocTy,GLLin,GLPstKy,GLPc,GLAc,GLBusA,GLRef,GLAcTy,RecTy)" & _
         " select     Amt,Sku,PstDte,GLDocNo,GLDocDte,GLAsg,GLDocTy,GLLin,GLPstKy,GLPc,GLAc,GLBusA,GLRef,GLAcTy,'*GLAnp' as RecTy from [#IGLAnp]"
W.Execute "Update [@Main] x inner join [$Sku] a on x.Sku=a.Sku" & _
" Set x.IsTax=a.IsTax, x.IsRepack=a.IsRepack, x.Is3p=a.Is3p, x.IsNoLongerTax=a.IsNoLongerTax" & _
" where RecTy in ('*GLDuty','*GLAnp')"
End Sub
Private Sub OMain_4_MdyCol_PstDte()
'DbtfChgDteToTxt W, OupTn_Main, "PstDte"
End Sub
Private Sub OMain_5_AddCol_PstMth()
W.Execute FmtQQ("Alter Table [?] add column PstMth Text(7)", OupTn_Main)
W.Execute FmtQQ("Update [?] set PstMth = Left(PstDte,7)", OupTn_Main)
End Sub
Private Sub OMain_3_AddCol_IsAlert(Tolerance&)
'---- Add Column IsAlert
WDrp "#A #B #C #D #E #F #G #H #Z"
W.Execute "Select Distinct PstDte,Sku,RecTy,Sum(x.Amt) as Amt into [#A] from [@Main] x group by PstDte,Sku,RecTy"
W.Execute "Update [#A] set RecTy='*GL' where RecTy in ('*GLDuty','*GLAnp')"
W.Execute "Select Distinct Sku,PstDte,RecTy,Sum(x.Amt) As Amt into [#B] from [#A] x group by Sku,PstDte,RecTy"
W.Execute "Select Distinct Sku,PstDte,Count(*) As RecTyCnt, Sum(x.Amt) as Amt into [#C] from [#B] x group by Sku,PstDte"
If HasReczQ(W, "Select * from [#C] where RecTyCnt>2") Then PgmEr CSub, "Table [#C] should have at most 2 different RecTy (*GLDuty | *GLAnp)"
W.Execute "Select x.Sku,x.PstDte,RecTy into [#D] from [#C] x inner join [#B] a on x.Sku=a.Sku and x.PstDte=a.PstDte where RecTyCnt=1"

W.Execute "Select x.Sku,x.PstDte,'*Only ' & Mid(x.RecTy,2) As IsAlert into [#Z] from [#D] x inner join [#B] a on x.Sku=a.Sku and a.PstDte=x.PstDte"

W.Execute "Select x.Sku,x.PstDte,Amt into [#E] from [#C] x where RecTyCnt=2"
W.Execute "Select x.Sku,x.PstDte,Amt,'#Match' as IsAlert into [#G] from [#C] x where Abs(Amt)<=" & Tolerance
W.Execute "Select x.Sku,x.PstDte,Amt,'*GL+'   as IsAlert into [#F] from [#C] x where Amt > +" & Tolerance
W.Execute "Select x.Sku,x.PstDte,Amt,'*GL-'   as IsAlert into [#H] from [#C] x where Amt < -" & Tolerance
W.Execute "Insert into [#Z] (Sku,PstDte,IsAlert) select Sku,PstDte,IsAlert from [#G]"
W.Execute "Insert into [#Z] (Sku,PstDte,IsAlert) select Sku,PstDte,IsAlert from [#F]"
W.Execute "Insert into [#Z] (Sku,PstDte,IsAlert) select Sku,PstDte,IsAlert from [#H]"

W.Execute "Alter table [@Main] add column IsAlert Text(10)"
'If Not DbtHasFld(W, "@Main", "IsAlert") Then
'    W.Execute "Alter table [@Main] add column IsAlert Text(10)"
'Else
'    W.Execute "Update [@Main] set IsAlert = Null"
'End If
W.Execute "Update [@Main] x inner join [#Z] a on x.Sku=a.Sku and x.PstDte=a.PstDte set x.IsAlert=a.IsAlert"
WDrp "#A #B #C #D #E #F #G #H #Z"
End Sub
Private Sub OMain_6_AddCol_Uom()
W.Execute "Alter Table [@Main] add column Des Text(255),StkUom Text(3),Ac_U Int, Lvl1 Text(2), Lvl2 Text(4), Lvl3 Text(7)"
W.Execute "update [@Main] x inner join [#IUom] a on x.Sku=a.Sku" & _
" set" & _
" x.Des = a.Des, x.StkUom=a.StkUom, x.AC_U=a.AC_U," & _
" x.Lvl1 = Left(ProdH,2)," & _
" x.Lvl2 = Left(ProdH,4)," & _
" x.Lvl3 = Left(ProdH,7)"
End Sub
Private Sub OMain_7_AddCol_Rank()
'#Rank
WDrp "#Rank #Rank1"
W.Execute "Select CInt(0) as Rank, Sku,PstDte,Sum(x.Amt) as Amt into [#Rank] from [@Main] x group by Sku,PstDte"
W.Execute "Update [#Rank] set Amt = Abs(Round(Nz(Amt,0),2))"
W.Execute "Select Distinct CInt(0) as Rank, Amt into [#Rank1] from [#Rank] order by Amt Desc"
UpdSeqFld W, "#Rank1", "Rank", "", ""
W.Execute "Update [#Rank] x inner join [#Rank1] a on x.Amt=a.Amt set x.Rank=a.Rank"

W.Execute "Alter Table [@Main] Add column Rank Int"
W.Execute "Update [@Main] x inner join [#Rank] a on x.Sku=a.Sku and x.PstDte=a.PstDte set x.Rank=a.Rank"
WDrp "#Rank #Rank1"
End Sub
Sub ReSeqFld(A As Database, T, Oup)

End Sub
Private Sub OMain_8_ReSeqFld()
ReSeqFld W, OupTn_Main, OupFld_Main
End Sub
Private Sub Brw_LnkSpec_GLAnp()
'LnkSpec_Dmp LnkSpec_GLAnp
End Sub
Private Sub Brw_LnkSpec_GLDuty()
'LnkSpec_Dmp LnkSpec_GLDuty
End Sub
Private Sub TblFld_PermitDDmp()
Debug.Print TblFld_PermitD
D Fny(W, "Permit")
End Sub
Private Sub TblFld_PermitDmp()
D Fny(W, "PermitD")
End Sub
Private Sub Z_TmpInpTn_Anp()
D TmpInpTn_GLAnp
End Sub
Private Function TmpInpTn_GLAnp$()
TmpInpTn_GLAnp = T1(WsFld_GLAnp)
End Function
Private Function TmpInpTn_GLDuty$()
TmpInpTn_GLDuty = T1(WsFld_GLDuty)
End Function
Private Function TmpInpTn_UOM$()
TmpInpTn_UOM = T1(WsFld_UOM)
End Function
Private Function TmpInpTn_MB51$()
TmpInpTn_MB51 = T1(WsFld_MB51)
End Function

Private Sub Brw_LnkSpec_Uom()
Brw_LnkSpec LnkSpec_Uom
End Sub
Private Sub Brw_LnkSpec(LnkSpec$)
Brw SplitVBar(LnkSpec)
End Sub
Private Sub Brw_LnkSpec_MB51()
Brw_LnkSpec LnkSpec_MB51
End Sub
Private Sub AddWcToTp()
AddWcToWbFmFbtt App.TpWb, App.WFb, "@Main @Rate @Sku @Repack1 @Repack2 @Repack2 @Repack3 @Repack4 @Repack5 @Repack6"
End Sub
Private Property Get IsDev() As Boolean
Stop '
'IsDev = FstChr(CurDbPth) = "C"
End Property
'''================================
'''================================
'''================================
'''================================
'Private Function FxDaoCnStr$(A)
''Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
''INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
''Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=C:\Users\sium\Desktop\TaxRate\sales text.xlsx;TABLE=Sheet1$
'Dim O$
'Select Case FfnExt(A)
'Case ".xlsx":: O = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & A & ";"
'Case ".xls": O = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & A & ";"
'Case Else: Stop
'End Select
'FxDaoCnStr = O
'End Function
'Private Sub WbAddWc(A As Workbook, Fb$, Nm$)
'A.Connections.Add2 Nm, Nm, FbWbCnStr(Fb$), Nm
'End Sub
'Private Function SplitCrLf(A) As String()
'SplitCrLf = Split(A, vbCrLf)
'End Function
'Private Function AyKeepLasN(A, N)
'Dim O, J&, I&, U&, Fm&, NewU&
'U = UB(A)
'If U < N Then AyKeepLasN = A: Exit Function
'O = A
'Fm = U - N + 1
'NewU = N - 1
'For J = Fm To U
'    Asg O(J), O(I)
'    I = I + 1
'Next
'ReDim Preserve O(NewU)
'AyKeepLasN = O
'End Function
'Private Sub ZZ_LinesKeepLasN()
'Dim Ay$(), A$, J%
'For J = 0 To 9
'Push Ay, "Line " & J
'Next
'A = Join(Ay, vbCrLf)
'Debug.Print LinesKeepLasN(A, 3)
'End Sub
'Private Function LinesKeepLasN$(A$, N%)
'Dim Ay$()
'Ay = SplitCrLf(A)
'LinesKeepLasN = JnCrLf(AyKeepLasN(Ay, N))
'End Function
'Private Function FbDaoCn(A) As Dao.Connection
'Set FbDaoCn = DBEngine.OpenConnection(A)
'End Function
'Private Function CvCtl(A) As Access.Control
'Set CvCtl = A
'End Function
'Private Function CvBtn(A) As Access.CommandButton
'Set CvBtn = A
'End Function
'Private Function IsBtn(A) As Boolean
'IsBtn = TypeName(A) = "CommandButton"
'End Function
'Private Function IsTglBtn(A) As Boolean
'IsTglBtn = TypeName(A) = "ToggleButton"
'End Function
'Private Function CvTgl(A) As Access.ToggleButton
'Set CvTgl = A
'End Function
'Private Sub CmdTurnOffTabStop(AcsCtl)
'Dim A As Access.Control
'Set A = AcsCtl
'If Not HasPfx(A.Name, "Cmd") Then Exit Sub
'Select Case True
'Case IsBtn(A): CvBtn(A).TabStop = False
'Case IsTglBtn(A): CvTgl(A).TabStop = False
'End Select
'End Sub
'Private Sub FrmSetCmdNotTabStop(A As Access.Form)
'ItrDo A.Controls, "CmdTurnOffTabStop"
'End Sub
'Private Function FxAdoCnStr$(A)
'FxAdoCnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;Extended Properties=""Excel 12.0;HDR=YES""", A)
'End Function
'Private Function FbAdoCnStr$(A)
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
'FbAdoCnStr = FmtQQ(C, A)
'End Function
'Private Function AdoCnStr_Cn(A) As ADODB.Connection
'Dim O As New ADODB.Connection
'O.Open A
'Set AdoCnStr_Cn = O
'End Function
'Private Function FxCn(A) As ADODB.Connection
'Set FxCn = AdoCnStr_Cn(FxAdoCnStr(A))
'End Function
'Private Function FbCn(A) As ADODB.Connection
'Set FbCn = AdoCnStr_Cn(FbAdoCnStr(A))
'End Function
'
'Private Function FxCat(A) As Catalog
'Set FxCat = CnCat(FxCn(A))
'End Function
'
'Private Function CnCat(Cn As AdoDb.Connection) As Catalog
'Dim O As New Catalog
'Set O.ActiveConnection = A
'Set CnCat = O
'End Function
'
'Private Function FbTny(A) As String()
'FbTny = CvSy(AyWhPredXPNot(CatTny(FbCat(A)), "HasPfx", "MSys"))
'End Function
'
'Private Function AyCln(A)
'Dim O
'O = A
'Erase O
'AyCln = O
'End Function
'Private Function AyWhPredXPNot(A, PredXP$, P)
'If Si(A) = 0 Then AyWhPredXPNot = AyCln(A): Exit Function
'Dim O, X
'O = AyCln(A)
'For Each X In A
'    If Not Run(PredXP, X, P) Then
'        Push O, X
'    End If
'Next
'AyWhPredXPNot = O
'End Function
'Private Function AyWhPredXP(A, PredXP$, P)
'If Si(A) = 0 Then AyWhPredXP = AyCln(A): Exit Function
'Dim O, X
'O = AyCln(A)
'For Each X In A
'    If Run(PredXP, X, P) Then
'        Push O, X
'    End If
'Next
'AyWhPredXP = O
'End Function
'Private Function FbCat(A) As Catalog
'Set FbCat = CnCat(FbCn(A))
'End Function
'Private Function CatTny(A As Catalog) As String()
'CatTny = ItrNy(A.Tables)
'End Function
'Private Function FxWsNy(A) As String()
'FxWsNy = CatTny(FxCat(A))
'End Function
'Private Function FxHasWs(A, WsNm$) As Boolean
'FxHasWs = AyHas(FxWsNy(A), WsNm)
'End Function
'
'Private Sub DbtImpTbl(A as Database, Tny0)
'Dim Tny$(), J%, S$
'Tny = DftNy(Tny0)
'For J = 0 To UB(Tny)
'    DbtDrp A, "#I" & Tny(J)
'    S = FmtQQ("Select * into [#I?] from [?]", Tny(J), Tny(J))
'    A.Execute S
'Next
'End Sub
'Private Function LnkColStr_Ly(A$) As String()
'Dim A1$(), A2$(), Ay() As LnkCol
'Ay = LnkColStr_LnkColAy(A)
'A1 = LnkColAy_Ny(Ay)
'A2 = AyAlignL(SyQuoteSqBkt(LnkColAy_ExtNy(Ay)))
'Dim J%, O$()
'For J = 0 To UB(A1)
'    Push O, A2(J) & "  " & A1(J)
'Next
'LnkColStr_Ly = O
'End Function
'Private Function AyLasEle(A)
'Asg A(UB(A)), AyLasEle
'End Function
'
'Private Function AscIsDig(A%) As Boolean
'AscIsDig = &H30 <= A And A <= &H39
'End Function
'
'Private Property Get LnkCol(Nm$, Ty As Dao.DataTypeEnum, Extnm$) As LnkCol
'Dim O As New LnkCol
'Set LnkCol = O.Init(Nm, Ty, Extnm)
'End Property
'
'Private Function LnkColStr_LnkColAy(A) As LnkCol()
'Dim Emp() As LnkCol, Ay$()
'Ay = SplitVBar(A): If Si(Ay) = 0 Then Stop
'LnkColStr_LnkColAy = AyMapInto(Ay, "LinLnkCol", Emp)
'End Function
'
'Private Function SplitVBar(A) As String()
'SplitVBar = Split(A, "|")
'End Function
'
'Private Function RmvSqBkt$(A)
'If IsSqBktQuoted(A) Then
'    RmvSqBkt = RmvFstLasChr(A)
'Else
'    RmvSqBkt = A
'End If
'End Function
'Private Sub ZZ_LinLnkCol()
'Dim A$, Act As LnkCol, Exp As LnkCol
'A = "AA Txt XX"
'Exp = LnkCol("AA", dbText, "AA")
'GoSub Tst
'Exit Sub
'Tst:
'Act = LinLnkCol(A)
'Debug.Assert LnkColIsEq(Act, Exp)
'Return
'End Sub
'Private Function LnkColIsEq(A As LnkCol, B As LnkCol) As Boolean
'With A
'    If .Extnm <> B.Extnm Then Exit Function
'    If .Ty <> B.Ty Then Exit Function
'    If .Nm <> B.Nm Then Exit Function
'End With
'LnkColIsEq = True
'End Function
'Private Function LinLnkCol(A) As LnkCol
'Dim Nm$, ShtTy$, Extnm$, Ty As Dao.DataTypeEnum
'LinTTRstAsg A, Nm, ShtTy, Extnm
'Extnm = RmvSqBkt(Extnm)
'Ty = DaoShtTy_Ty(ShtTy)
'Set LinLnkCol = LnkCol(Nm, Ty, IIf(Extnm = "", Nm, Extnm))
'End Function
'Private Function RmvFstLasChr$(A)
'RmvFstLasChr = RmvFstChr(RmvLasChr(A))
'End Function
'Private Function CnStrzT$(A as Database, T)
'CnStrzT = A.TableDefs(T).Connect
'End Function
'Private Sub DbtImpMap(A as Database, T, LnkColStr$, Optional WhBExpr$)
'If FstChr(T) <> ">" Then
'    Debug.Print "FstChr of T must be >"
'    Stop
'End If
''Assume [>?] T exist
''Create [#I?] T
'Dim S$
'S = LnkColStr_ImpSql(LnkColStr, T, WhBExpr)
'DbtDrp A, "#I" & Mid(T, 2)
'A.Execute S
'End Sub
'
'Private Function LnkColStr_ImpSql$(A$, T, Optional WhBExpr$)
'Dim Ay() As LnkCol
'Ay = LnkColStr_LnkColAy(A)
'LnkColStr_ImpSql = LnkColAy_ImpSql(Ay, T, WhBExpr)
'End Function
'
'Private Function IsSqBktQuoted(A) As Boolean
'If FstChr(A) <> "[" Then Exit Function
'If LasChr(A) <> "]" Then Exit Function
'IsSqBktQuoted = True
'End Function
'
'Private Function FstChr$(A)
'FstChr = Left(A, 1)
'End Function
'
'Private Function LasChr$(A)
'LasChr = Right(A, 1)
'End Function
'Private Property Get Drs(Fny0, Dry()) As Drs
'Dim O As New Drs
'Drs = O.Init(DftNy(Fny0), Dry)
'End Property
'Private Function ApSy(ParamArray Ap()) As String()
'Dim Av(): Av = Ap
'Dim O$(), J%, U&
'U = UB(Av)
'ReDim O(U)
'For J = 0 To U
'    O(J) = Av(J)
'Next
'ApSy = O
'End Function
'Private Function DbtHasFld(A as Database, T, F$) As Boolean
'DbtHasFld = ItrHasNm(A.TableDefs(T).Fields, F)
'End Function
'
'Private Function DbDrpTbl(A as Database, Tny0)
'AyDoPX DftNy(Tny0), "DbtDrp", A
'End Function
'Private Sub SavRec()
'DoCmd.RunCommand acCmdSaveRecord
'End Sub
'
'Private Sub AyDoPX(A, PXFunNm$, P)
'If Si(A) = 0 Then Exit Sub
'Dim I
'For Each I In A
'    Run PXFunNm, P, I
'Next
'End Sub
'Private Function DbqRs(A as Database, Sql) As Dao.Recordset
'Set DbqRs = A.OpenRecordset(Sql)
'End Function
'Private Function Acs() As Access.Application
'Static X As Boolean, Y As Access.Application
'On Error GoTo X
'If X Then
'    Set Y = New Access.Application
'    X = True
'End If
'If Y.Application.Name = "Microsoft Access" Then
'    Set Acs = Y
'    Exit Function
'End If
'X:
'    Set Y = New Access.Application
'    Debug.Print "Acs: New Acs instance is crreated."
'Set Acs = Y
'End Function
'
'Private Sub AcsVis(A As Access.Application)
'If Not A.Visible Then A.Visible = True
'End Sub
'
'Private Function IsNothing(A) As Boolean
'IsNothing = TypeName(A) = "Nothing"
'End Function
'Private Function SyAddPfx(A, Pfx) As String()
'If Si(A) = 0 Then Exit Function
'Dim O$(), U&, J&
'U = UB(A)
'ReDim O(U)
'For J = 0 To U
'    O(J) = Pfx & A(J)
'Next
'SyAddPfx = O
'End Function
'Private Function IsObjAy(A) As Boolean
'IsObjAy = VarType(A) = vbArray + vbObject
'End Function
'Private Function SyRmvEleAt(A, Optional At&)
'Dim O, J&, U&
'U = UB(A)
'O = A
'Select Case True
'Case U = 0
'    Erase O
'    SyRmvEleAt = O
'    Exit Function
'Case IsObjAy(A)
'    For J = At To U - 1
'        Set O(J) = O(J + 1)
'    Next
'Case Else
'    For J = At To U - 1
'        O(J) = O(J + 1)
'    Next
'End Select
'ReDim Preserve O(U - 1)
'SyRmvEleAt = O
'End Function
'Private Sub ZZZ_AyShift()
'Dim Ay(), Exp, Act, ExpAyAft()
'Ay = Array(1, 2, 3, 4)
'Exp = 1
'ExpAyAft = Array(2, 3, 4)
'GoSub Tst
'Exit Sub
'Tst:
'Act = AyShift(Ay)
'Debug.Assert IsEq(Exp, Act)
'Debug.Assert AyIsEq(Ay, ExpAyAft)
'Return
'End Sub
'Private Function AyShift(Ay)
'AyShift = Ay(0)
'Ay = SyRmvEleAt(Ay)
'End Function
'Private Sub ZZZ_PfxSsl_Sy()
'Dim A$, Exp$()
'A = "A B C D"
'Exp = SslSy("AB AC AD")
'GoSub Tst
'Exit Sub
'Tst:
'Dim Act$()
'Act = PfxSsl_Sy(A)
'Debug.Assert AyIsEq(Act, Exp)
'Return
'End Sub
'Private Function ItrFstPrpEq(A, PrpNm$, V)
'Dim I, OP
'For Each I In A
'    OP = Prp(I, PrpNm)
'    If OP = V Then Asg I, ItrFstPrpEq: Exit Function
'Next
'Debug.Print PrpNm, V
'For Each I In A
'    Debug.Print Prp(I, PrpNm)
'Next
'Stop
'End Function
'Private Function Prp(A, PrpNm$)
'On Error GoTo X
'Dim V
'V = CallByName(A, PrpNm, VbGet)
'Asg V, Prp
'Exit Function
'X:
'Debug.Print "Prp: " & Err.Description
'End Function
'Private Function ItrPrpSy(A, PrpNm$) As String()
'ItrPrpSy = ItrPrpInto(A, PrpNm, EmpSy)
'End Function
'Private Function ItrPrpInto(A, PrpNm$, OInto)
'Dim O, I
'O = OInto
'Erase O
'For Each I In A
'    Push O, Prp(I, PrpNm)
'Next
'ItrPrpInto = O
'End Function
'Private Function WbWsCdNy(A As Workbook) As String()
'WbWsCdNy = ItrPrpSy(A.Sheets, "CodeName")
'End Function
'Private Function FxWsCdNy(A) As String()
'Dim Wb As Workbook
'Set Wb = FxWb(A)
'FxWsCdNy = WbWsCdNy(Wb)
'Wb.Close False
'End Function
'Private Function PfxSsl_Sy(A) As String()
'Dim Ay$(), Pfx$
'Ay = SslSy(A)
'Pfx = AyShift(Ay)
'PfxSsl_Sy = SyAddPfx(Ay, Pfx)
'End Function
'Private Function ApnWAcs(A$)
'Dim O As Access.Application
'AcsOpn O, ApnWFb(A)
'Set ApnWAcs = O
'End Function
'Private Function ApnAcs(A$) As Access.Application
'AcsOpn Acs, ApnWFb(A)
'Set ApnAcs = Acs
'End Function
'Private Sub AcsOpn(A As Access.Application, Fb$)
'Select Case True
'Case IsNothing(A.CurrentDb)
'    A.OpenCurrentDatabase Fb
'Case A.CurrentDb.Name = Fb
'Case Else
'    A.CurrentDb.Close
'    A.OpenCurrentDatabase Fb
'End Select
'End Sub
'Private Sub ApnBrwWDb(A$)
'Dim Fb$
'Fb = ApnWFb(A)
'AcsOpn Acs, Fb
'AcsVis Acs
'End Sub
'Private Sub FbEns(A$)
'If FfnIsExist(A) Then Exit Sub
'FbCrt A
'End Sub
'Private Sub FbCrt(A$)
'DBEngine.CreateDatabase A, dbLangGeneral
'End Sub
'Private Sub RfhWcStr(A, Fb$)
'WbRfhCnStr(FxWb(A), Fb).Close True
'End Sub
'Private Function WbRfhCnStr(A As Workbook, Fb$) As Workbook
'ItrDoXP A.Connections, "RfhWcCnStr", FbWbCnStr(Fb$)
'Set WbRfhCnStr = A
'End Function
'Private Sub OpnFb(A)
'Acs.OpenCurrentDatabase A
'AcsVis Acs
'End Sub
'Private Function FbDb(A) As Database
'Set FbDb = DBEngine.OpenDatabase(A)
'End Function
'
'Private Function ApnWFb$(A$)
'ApnWFb = ApnWPth(A) & "Wrk.accdb"
'End Function
'
'Private Function ApnWPth$(A$)
'Dim P$
'P = TmpPthHom & A & "\"
'EnsPth P
'ApnWPth = P
'End Function
'Private Function DbIsOk(A As Database) As Boolean
'On Error GoTo X
'DbIsOk = IsStr(A.Name)
'Exit Function
'X:
'End Function
'
'Private Function ApnWDb(A$) As Database
'Static X As Boolean, Y As Database
'If Not X Then
'    X = True
'    FbEns ApnWFb(A)
'    Set Y = FbDb(ApnWFb(A))
'End If
'If Not DbIsOk(Y) Then Set Y = FbDb(ApnWFb(A))
'Set ApnWDb = Y
'End Function
'Private Function DbqAny(A as Database, Sql) As Boolean
'DbqAny = RsAny(DbqRs(A, Sql))
'End Function
'Private Function DbHasTbl(A as Database, T$) As Boolean
'DbHasTbl = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type in (1,6)", T))
'End Function
'Private Function WdtzSy%(A)
'Dim O%, J&
'For J = 0 To UB(A)
'    O = Max(O, Len(A(J)))
'Next
'WdtzSy = O
'End Function
'Private Function TblStru(Tny0) As String()
'TblStru = DbtStru(CurrentDb, Tny0)
'End Function
'
'Private Function TblSql$(T, Optional WhBExpr$)
'TblSql = FmtQQ("Select * from [?]?", T, SqpWhere(WhBExpr))
'End Function
'Private Function FbtFny(A, T$) As String()
'FbtFny = RsFny(DbqRs(FbDb(A), TblSql(T)))
'End Function
'Private Function Max(A, B)
'If A > B Then
'    Max = A
'Else
'    Max = B
'End If
'End Function
'Private Function Min(A, B)
'If A > B Then
'    Min = B
'Else
'    Min = A
'End If
'End Function
'
'Private Function DftNy(Ny0) As String()
'Select Case True
'Case IsMissing(Ny0)
'Case IsStr(Ny0): DftNy = SslSy(Ny0)
'Case IsSy(Ny0): DftNy = Ny0
'Case IsArray(Ny0): DftNy = AySy(Ny0)
'Case Else: Stop
'End Select
'End Function
'Private Function AySy(A) As String()
'If Si(A) = 0 Then Exit Function
'AySy = ItrInto(A, EmpSy)
'End Function
'Private Function EmpSy() As String()
'End Function
'Private Function EmpAy() As Variant()
'End Function
'
'Private Function ItrInto(A, OInto)
'Dim O, I
'O = OInto
'Erase O
'For Each I In A
'    Push O, I
'Next
'ItrInto = O
'End Function
'Private Function OupPth$()
'OupPth = PmnmVal("OupPth")
'End Function
'Private Function YYYYMMDD_IsVdt(A) As Boolean
'On Error Resume Next
'YYYYMMDD_IsVdt = Format(CDate(A), "YYYY-MM-DD") = A
'End Function
'Private Function TpPth$()
'TpPth = PthzCurDb & "PgmObj\Template\"
'End Function
'Private Function FfnPth$(A)
'Dim P%: P = InStrRev(A, "\")
'If P = 0 Then Exit Function
'FfnPth = Left(A, P)
'End Function
'Private Function ErzFws__2(Fx$, WsNm$, ColNy$()) As String()
'
'End Function
'Private Function ErzFws__3(Fx$, WsNm$, ColNy$(), DtaTyAy() As Dao.DataTypeEnum) As String()
'
'End Function
'Private Sub ZZ_ErAyzFxWsMissingCol()
''" [Material]             As Sku," & _
''" [Plant]                As Whs," & _
''" [Storage Location]     As Loc," & _
''" [Batch]                As BchNo," & _
''" [Unrestricted]         As OH " & _
'
'End Sub
'Private Function TblF_Ty(T, F) As Dao.DataTypeEnum
'
'End Function
'Private Function TblErAyzCol(A as Database, T, ColNy$(), DtaTyAy() As Dao.DataTypeEnum, Optional AddTblLinMsg As Boolean) As String()
'Dim Fny$(), F, Fny1$(), Fny2$()
'Fny = FnyzT(A, T)
'For Each F In ColNy
'    If AyHas(Fny, F) Then
'        Push F, Fny1
'    Else
'        Push F, Fny2
'    End If
'Next
'Dim O$()
'If Si(Fny2) > 0 Then
'    Dim J%
'    For J = 0 To UB(ColNy)
'        If AyHas(Fny2, ColNy(J)) Then
'            If TblF_Ty(T, ColNy(J)) <> DtaTyAy(J) Then
'                Push O, "Column [?] has unexpected DataType[?].  It is expected to be [?]"
'            End If
'        End If
'    Next
'End If
'If AddTblLinMsg Then
'    Push O, ""
'
'End If
'End Function
'Private Function ErzFfnNotExist(A) As String()
'Dim O$(), M$
'If Not FfnIsExist(A) Then
'    Push O, A
'    M = "Above file not exist"
'    Push O, M
'    Push O, String(Len(M), "-")
'End If
'ErzFfnNotExist = O
'End Function
'Private Function ErzThen(ParamArray ErFunNmAp()) As String()
'Dim Av(), O$(), I
'Av = ErFunNmAp
'For Each I In Av
'    O = Run(I)
'    If Si(O) > 0 Then
'        ErzThen = O
'    End If
'Next
'End Function
'Private Function UnderLin$(A)
'UnderLin = String(Len(A), "-")
'End Function
'Private Function UnderLinDbl$(A)
'UnderLinDbl = String(Len(A), "=")
'End Function
'Private Function ErzFxWs(A, WsNm$) As String()
''ErThen "ErzFfnNotExist ErzFxHasNoWs"
'Dim O$()
'O = ErzFfnNotExist(A)
'If Si(O) > 0 Then
'    ErzFxWs = O
'    Exit Function
'End If
'
''B = ErzFxWs__1(A, WsNm)
'If Si(A) > 0 Then
''    ErAyzFxWs = A
'    Exit Function
'End If
'
'
'If Not FfnIsExist(A) Then
'    Push O, A
'    Push O, "Above Excel file not found"
'    Push O, "--------------------------"
'    'ErAyzFxWsLnk = O
'    Exit Function
'End If
'Dim B$
''B = FxWs_LnkErMsg(A, WsNm)
'If B <> "" Then
'    Push O, "Excel File: " & A
'    Push O, "Worksheet : " & WsNm
'    Push O, "System Msg: " & B
'    Push O, "Above Excel file & Worksheet cannot be linked to Access"
'    Push O, "-------------------------------------------------------"
'    'ErAyzFxWsLnk = O
'    Exit Function
'End If
'On Error GoTo X
'TblLnkFx "#", CStr(A), WsNm
'DrpT CurrentDb, "#"
'Exit Function
'X:
''FxWs_LnkErMsg = Err.Description
'
''A = ErAyzFxWsMissingCol(
'End Function
'Private Function PthzCurDb$()
'PthzCurDb = FfnPth(CurrentDb.Name)
'End Function
'Private Property Get PmnmVal$(Pmnm$)
'PmnmVal = CurrentDb.TableDefs("Prm").OpenRecordset.Fields(Pmnm).Value
'End Property
'Private Property Let PmnmVal(Pmnm$, V$)
'Stop
''Should not use
'With CurrentDb.TableDefs("Prm").OpenRecordset
'    .Edit
'    .Fields(Pmnm).Value = V
'    .Update
'End With
'End Property
'
'Private Function FldsFny(A As Dao.Fields) As String()
'FldsFny = ItrNy(A)
'End Function
'Private Sub PthBrw(A)
'Shell FmtQQ("Explorer ""?""", A), vbMaximizedFocus
'End Sub
'Private Function EnsPthSfx$(A)
'If Right(A, 1) <> "\" Then
'    EnsPthSfx = A & "\"
'Else
'    EnsPthSfx = A
'End If
'End Function
'Private Function ItrNy(A) As String()
'Dim O$(), I
'For Each I In A
'    Push O, I.Name
'Next
'ItrNy = O
'End Function
'Private Sub Push(O, M)
'Dim N&
'N = Si(O)
'ReDim Preserve O(N)
'If IsObject(M) Then
'    Set O(N) = M
'Else
'    O(N) = M
'End If
'End Sub
'Private Sub ZZ_PthFxAy()
'Dim A$()
''A = PthFxAy(PermitImpPth)
'D A
'End Sub
'
'Private Function DteIsVdt(A$) As Boolean
'On Error Resume Next
'DteIsVdt = Format(CDate(A), "YYYY-MM-DD") = A
'End Function
'Private Sub ZZ_Fny()
'Dim Db As database
'D Fny(A, ">KE24")
'End Sub
'Private Function RsSy(A As Dao.Recordset) As String()
'Dim O$()
'With A
'    While Not .EOF
'        Push O, .Fields(0).Value
'        .MoveNext
'    Wend
'End With
'RsSy = O
'End Function
'Private Sub ZZ_SqlFny()
'Const S$ = "SELECT qSku.*" & _
'" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
'D SqlFny(S)
'End Sub
'Private Function SqlFny(A) As String()
'SqlFny = RsFny(SqlRs(A))
'End Function
'Private Sub ZZ_SqlRs()
'Const S$ = "SELECT qSku.*" & _
'" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
'AyBrw RsCsvLy(SqlRs(S))
'End Sub
'
'Private Function SqlRs(A) As Dao.Recordset
'Set SqlRs = CurrentA.OpenRecordset(A)
'End Function
'Private Sub ZZ_SqlSy()
'D SqlSy("Select Distinct UOR from [>Imp]")
'End Sub
'Private Function SqpzInBExpr$(Ay, FldNm$, Optional WithQuote As Boolean)
'Const C$ = "[?] in (?)"
'Dim B$
'    If WithQuote Then
'        B = JnComma(SyQuoteSng(Ay))
'    Else
'        B = JnComma(Ay)
'    End If
'SqpzInBExpr = FmtQQ(C, FldNm, B)
'End Function
'Private Function SqlSy(A) As String()
'SqlSy = DbqSy(CurrentDb, A)
'End Function
'Private Function AyAdd(A, B)
'Dim O
'O = A
'PushAy O, B
'AyAdd = O
'End Function
'Private Sub TblBrw(T)
'DoCmd.OpenTable T
'End Sub
'
'Private Function DbtFny(A as Database, T$) As String()
'DbtFny = RsFny(RszT(A, T))
'End Function
'Private Function SplitSpc(A) As String()
'SplitSpc = Split(A, " ")
'End Function
'Private Function SqlAny(A) As Boolean
'SqlAny = DbqAny(CurrentDb, A)
'End Function
'Private Function RsAny(A As Dao.Recordset) As Boolean
'RsAny = Not A.EOF
'End Function
'Private Function TblIsExist(T) As Boolean
'TblIsExist = DbHasTbl(CurrentDb, T)
'End Function
'Private Sub TblOpn(TblSsl$)
'AyDo SslSy(TblSsl), "TblOpn_1"
'End Sub
'Private Sub AyDo(A, FunNm$)
'If Si(A) = 0 Then Exit Sub
'Dim I
'For Each I In A
'    Run FunNm, I
'Next
'End Sub
'Private Sub TblOpn_1(T)
'DoCmd.OpenTable T
'End Sub
'Private Function RplDblSpc$(A)
'Dim P%, O$, J%
'O = A
'While InStr(O, "  ") > 0
'    J = J + 1
'    If J > 50000 Then Stop
'    O = Replace(O, "  ", " ")
'Wend
'RplDblSpc = O
'End Function
'
'Private Function SslSy(A) As String()
'SslSy = SplitSpc(RplDblSpc(Trim(A)))
'End Function
'Private Sub ItrNmDo(A, DoFun$)
'Dim I
'For Each I In A
'    Run DoFun, I.Name
'Next
'End Sub
'Private Sub AcsClsTbl(A As Access.Application)
'Dim T
'For Each T In A.CodeData.AllTables
'    A.DoCmd.Close acTable, T.Name
'Next
'End Sub
'
'Private Sub AcsTbl_Cls(A As Access.Application, Tny0)
''AyNmDoPX A.CodeData.AllTables, "AcsTbl_Cls"
'End Sub
'
'
'Private Function DbHasQry(A as Database, Q) As Boolean
'DbHasQry = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type=5", Q))
'End Function
'
'Private Sub DbDrpQry(A as Database, Q)
'If DbHasQry(A, Q) Then A.QueryDefs.Delete Q
'End Sub
'
'Private Sub DbCrtQry(A as Database, Q, Sql$)
'Dim QQ As New QueryDef
'DbDrpQry A, Q
'QQ.Sql = Sql
'QQ.Name = Q
'A.QueryDefs.Append QQ
'End Sub
'
'Private Function LinShiftTerm$(O$)
'Dim A$, P%
'A = LTrim(O)
'P = InStr(A, " ")
'If P = 0 Then
'    LinShiftTerm = A
'    O = ""
'    Exit Function
'End If
'LinShiftTerm = Left(A, P - 1)
'O = LTrim(Mid(A, P + 1))
'End Function
'
'Private Sub LinTTRstAsg(A, OT1$, OT2$, ORst$)
'Dim Ay$()
'Ay = LinTTRst(A)
'OT1 = Ay(0)
'OT2 = Ay(1)
'ORst = RTrim(Ay(2))
'End Sub
'
'Private Function LinTTRst(A) As String()
'Dim O$(2), L$
'L = A
'O(0) = LinShiftTerm(L)
'O(1) = LinShiftTerm(L)
'O(2) = L
'LinTTRst = O
'End Function
'Private Function AyMinus(A, B)
'If Si(B) = 0 Then AyMinus = A: Exit Function
'If Si(A) = 0 Then AyMinus = A: Exit Function
'Dim O, I
'O = A
'Erase O
'For Each I In A
'    If Not AyHas(B, I) Then Push O, I
'Next
'AyMinus = O
'End Function
'
'Private Sub DbtRen(A as Database, Fm$, ToTbl$)
'A.TableDefs(Fm).Name = ToTbl
'End Sub
'
'Private Function DbtChkCol(A as Database, T, LnkColStr$) As String()
'Dim Ay() As LnkCol, O$(), Fny$(), J%, Ty As Dao.DataTypeEnum, F$
'Ay = LnkColStr_LnkColAy(LnkColStr)
'Fny = LnkColAy_ExtNy(Ay)
'O = DbtChkFny(A, T, Fny)
'If Si(O) > 0 Then DbtChkCol = O: Exit Function
'For J = 0 To UB(Ay)
'    F = Ay(J).Extnm
'    Ty = Ay(J).Ty
'    PushNonEmpty O, DbtChkFldType(A, T, F, Ty)
'Next
'If Si(0) > 0 Then
'    PushMsgUnderLin O, "Some field has unexpected type"
'    DbtChkCol = O
'End If
'End Function
'Private Function TakAft$(A, S)
'Dim P%
'P = InStr(A, S)
'If P = 0 Then Exit Function
'TakAft = Mid(A, P + Len(S))
'End Function
'Private Function TakBefOrAll$(A, S)
'Dim O$
'O = TakBef(A, S)
'If O = "" Then
'    TakBefOrAll = A
'Else
'    TakBefOrAll = O
'End If
'End Function
'Private Function TakAftOrAll$(A, S)
'Dim O$
'O = TakAft(A, S)
'If O = "" Then
'    TakAftOrAll = A
'Else
'    TakAftOrAll = O
'End If
'End Function
'
'
'Private Function TakBef$(A, S)
'Dim P%
'P = InStr(A, S)
'If P = 0 Then Exit Function
'TakBef = Left(A, P - 1)
'End Function
'
'Private Function FxwzDbt(A as Database, T$) As Fxw
'Dim Cn$
'Cn = CnStrzT(A, T)
'If Not IsPfx(Cn, "Excel") Then Exit Function
'With FxwzDbt
'    .Fx = TakBefOrAll(TakAft(Cn, "DATABASE="), ";")
'    .Wsn = A.TableDefs(T).SourceTableName
'    If LasChr(.Wsn) <> "$" Then Stop
'    .Wsn = RmvLasChr(.Wsn)
'End With
'End Function
'
'Private Function ISpecINm$(A$)
'ISpecINm = LinT1(A)
'End Function
'Private Sub LnkSpec_Dmp(A)
'Debug.Print RplVBar(A)
'End Sub
'Private Function LnkSpec_Ly(A) As String()
'Const L2Spec$ = ">GLAnp |" & _
'    "Whs    Txt Plant |" & _
'    "Loc    Txt [Storage Location]|" & _
'    "Sku    Txt Material |" & _
'    "PstDte Txt [Posting Date] |" & _
'    "MovTy  Txt [Movement Type]|" & _
'    "Qty    Txt Quantity|" & _
'    "BchNo  Txt Batch |" & _
'    "Where Plant='8601' and [Storage Location]='0002' and [Movement Type] like '6*'"
'End Function
'Private Function HasPfx(A, Pfx$) As Boolean
'HasPfx = Left(A, Len(Pfx)) = Pfx
'End Function
'Private Sub LnkSpec_Asg(A, Optional OTblNm$, Optional OLnkColStr$, Optional OWhBExpr$)
'Dim Ay$()
'Ay = AyTrim(SplitVBar(A))
'OTblNm = AyShift(Ay)
'If LinT1(AyLasEle(Ay)) = "Where" Then
'    OWhBExpr = LinRmvTerm(Pop(Ay))
'Else
'    OWhBExpr = ""
'End If
'OLnkColStr = JnVBar(Ay)
'End Sub
'Private Function Pop(A)
'Pop = AyLasEle(A)
'SyRmvLasEle A
'End Function
'Private Sub SyRmvLasEle(A)
'If Si(A) = 1 Then
'    Erase A
'    Exit Sub
'End If
'ReDim Preserve A(UB(A) - 1)
'End Sub
'Private Function JnVBar$(A)
'JnVBar = Join(A, "|")
'End Function
'Private Sub LnkSpec_Ay_Asg(A$(), OTny$(), OLnkColStrAy$(), OWhBExprAy$())
'Dim U%, J%
'U = UB(A)
'ReDim OTny(U)
'ReDim OLnkColStrAy(U)
'ReDim OWhBExprAy(U)
'For J = 0 To U
'    LnkSpec_Asg A(J), OTny(J), OLnkColStrAy(J), OWhBExprAy(J)
'Next
'End Sub
'
'Private Function DbImp(A as Database, LnkSpec_$()) As String()
'Dim O$(), J%, T$(), L$(), W$(), U%
'LnkSpec_Ay_Asg LnkSpec_, T, L, W
'U = UB(LnkSpec_)
'For J = 0 To U
'    PushAy O, DbtChkCol(A, T(J), L(J))
'Next
'If Si(O) > 0 Then DbImp = O: Exit Function
'For J = 0 To U
'    DbtImpMap A, T(J), L(J), W(J)
'Next
'DbImp = O
'End Function
'
'Private Function DbtMissFny_Er(A as Database, T, MissFny$(), ExistingFny$()) As String()
'Dim X As Fxw, O$(), I
'If Si(MissFny) = 0 Then Exit Function
'X = FxwzDbt(A, T)
'If X.Fx <> "" Then
'    Push O, "Excel File       : " & X.Fx
'    Push O, "Worksheet        : " & X.Wsn
'    PushUnderLin O
'    For Each I In ExistingFny
'        Push O, "Worksheet Column : " & QuoteSqBkt(I)
'    Next
'    PushUnderLin O
'    For Each I In MissFny
'        Push O, "Missing Column   : " & QuoteSqBkt(I)
'    Next
'    PushMsgUnderLinDbl O, "Columns are missing"
'Else
'    Push O, "Database : " & A.Name
'    Push O, "Table    : " & T
'    For Each I In MissFny
'        Push O, "Field    : " & QuoteSqBkt(I)
'    Next
'    PushMsgUnderLinDbl O, "Above Fields are missing"
'End If
'DbtMissFny_Er = O
'End Function
'
'Private Function DbtChkFny(A as Database, T, ExpFny$()) As String()
'Dim Miss$(), TFny$(), O$(), I
'TFny = DbtFny(A, T)
'Miss = AyMinus(ExpFny, TFny)
'DbtChkFny = DbtMissFny_Er(A, T, Miss, TFny)
'End Function
'Private Function QuoteSqBkt$(A)
'QuoteSqBkt = "[" & A & "]"
'End Function
'Private Function PushMsgUnderLin(O$(), M$)
'Push O, M
'Push O, UnderLin(M)
'End Function
'Private Function PushUnderLin(O$())
'Push O, UnderLin(AyLasEle(O))
'End Function
'Private Function PushUnderLinDbl(O$())
'Push O, UnderLinDbl(AyLasEle(O))
'End Function
'Private Function PushMsgUnderLinDbl(O$(), M$)
'Push O, M
'Push O, UnderLinDbl(M)
'End Function
'Private Function DaoTy_ShtTy$(A As Dao.DataTypeEnum)
'Dim O$
'Select Case A
'Case Dao.DataTypeEnum.dbByte: O = "Byt"
'Case Dao.DataTypeEnum.dbLong: O = "Lng"
'Case Dao.DataTypeEnum.dbInteger: O = "Int"
'Case Dao.DataTypeEnum.dbDate: O = "Dte"
'Case Dao.DataTypeEnum.dbText: O = "Txt"
'Case Dao.DataTypeEnum.dbBoolean: O = "Yes"
'Case Dao.DataTypeEnum.dbDouble: O = "Dbl"
'Case Else: Stop
'End Select
'DaoTy_ShtTy = O
'End Function
'Private Function DaoShtTy_Ty(A$) As Dao.DataTypeEnum
'Dim O As Dao.DataTypeEnum
'Select Case A
'Case "Byt": O = Dao.DataTypeEnum.dbByte
'Case "Lng": O = Dao.DataTypeEnum.dbLong
'Case "Int": O = Dao.DataTypeEnum.dbInteger
'Case "Dte": O = Dao.DataTypeEnum.dbDate
'Case "Txt": O = Dao.DataTypeEnum.dbText
'Case "Yes": O = Dao.DataTypeEnum.dbBoolean
'Case "Dbl": O = Dao.DataTypeEnum.dbDouble
'Case Else: Stop
'End Select
'DaoShtTy_Ty = O
'End Function
'Private Function DftFfnAy(FfnAy0) As String()
'Select Case True
'Case IsStr(FfnAy0): DftFfnAy = ApSy(FfnAy0)
'Case IsSy(FfnAy0): DftFfnAy = FfnAy0
'Case IsArray(FfnAy0): DftFfnAy = AySy(FfnAy0)
'End Select
'End Function
'Private Property Get FfnCpyToPthIfDif(FfnAy0, Pth$) As String()
'Const M_Sam$ = "File is same the one in Path."
'Const M_Copied$ = "File is copied to Path."
'Const M_NotFnd$ = "File not found, cannot copy to Path."
'Dim B$, Ay$(), I, O$(), M$(), Msg$
'Ay = DftFfnAy(FfnAy0): If Si(Ay) = 0 Then Exit Property
'For Each I In Ay
'    Select Case True
'    Case FfnIsExist(I)
'        B = Pth & Fn(I)
'        Select Case True
'        Case FfnIsSam(B, CStr(I))
'            Msg = M_Sam: GoSub Prt
'        Case Else
'            Fso.CopyFile I, B, True
'            Msg = M_Copied: GoSub Prt
'        End Select
'    Case Else
'        Msg = M_NotFnd: GoSub Prt
'        Push O, "File : " & I
'    End Select
'Next
'If Si(O) > 0 Then
'    PushMsgUnderLinDbl O, "Above files not found"
'    FfnCpyToPthIfDif = O
'End If
'Exit Property
'Prt:
'    Debug.Print FmtQQ("FfnCpyToPthIfDif: ? Path=[?] File=[?]", Msg, Pth, I)
'    Return
'End Property
'Private Function FfnIsSamMsg(A$, B$, Si&, Tim$, Optional Msg$) As String()
'Dim O$()
'Push O, "File 1   : " & A
'Push O, "File 2   : " & B
'Push O, "File Size: " & Si
'Push O, "File Time: " & Tim
'Push O, "File 1 and 2 have same size and time"
'If Msg <> "" Then Push O, Msg
'FfnIsSamMsg = O
'End Function
'Private Function FfnIsSam(A$, B$) As Boolean
'If FfnTim(A) <> FfnTim(B) Then Exit Function
'If SizFfn(A) <> SizFfn(B) Then Exit Function
'FfnIsSam = True
'End Function
'Private Function SizFfn&(A$)
'If FfnIsExist(A) Then
'    SizFfn = FileLen(A)
'Else
'    SizFfn = -1
'End If
'End Function
'Private Function FfnTim(A$) As Date
'If FfnIsExist(A) Then FfnTim = FileDateTime(A)
'End Function
'Private Function AyTrim(A) As String()
'If Si(A) = 0 Then Exit Function
'Dim O$(), J&, U&
'U = UB(A)
'ReDim O(U)
'For J = 0 To U
'    O(J) = Trim(A(J))
'Next
'AyTrim = O
'End Function
'Private Function DbtChkFldType$(A as Database, T, F, Ty As Dao.DataTypeEnum)
'Dim ActTy As Dao.DataTypeEnum
'ActTy = A.TableDefs(T).Fields(F).Type
'If ActTy <> Ty Then
'    DbtChkFldType = FmtQQ("Table[?] field[?] should have type[?], but now it has type[?]", T, F, DaoTy_ShtTy(Ty), DaoTy_ShtTy(ActTy))
'End If
'End Function
'Private Function OyPrpSy(A, PrpNm$)
'OyPrpSy = OyPrpInto(A, PrpNm, EmpSy)
'End Function
'Private Function OyPrpInto(A, PrpNm$, OInto)
'Dim O, J&
'O = OInto
'Erase O
'For J = 0 To UB(A)
'    Push O, Prp(A(J), PrpNm)
'Next
'OyPrpInto = O
'End Function
'Private Function LnkColAy_ExtNy(A() As LnkCol) As String()
'LnkColAy_ExtNy = OyPrpSy(A, "Extnm")
'End Function
'Private Function LnkColAy_Ny(A() As LnkCol) As String()
'LnkColAy_Ny = OyPrpSy(A, "Nm")
'End Function
'Private Sub WbVdtOupNy(A As Workbook, OupNy$())
'Dim O$(), N$, B$(), WsCdNy$()
'WsCdNy = WbWsCdNy(A)
'O = AyMinus(SyAddPfx(OupNy, "WsO"), WsCdNy)
'If Si(O) > 0 Then
'    N = "OupNy":  B = OupNy:  GoSub Dmp
'    N = "WbCdNy": B = WsCdNy: GoSub Dmp
'    N = "Mssing": B = O:      GoSub Dmp
'    Stop
'    Exit Sub
'End If
'Exit Sub
'Dmp:
'Debug.Print UnderLin(N)
'Debug.Print N
'Debug.Print UnderLin(N)
'D B
'Return
'End Sub
'Private Function RsDrs(A As Dao.Recordset) As Drs
'Dim Fny$(), Dry()
'Fny = RsFny(A)
'Dry = RsDry(A)
'Set RsDrs = Drs(Fny, Dry)
'End Function
'Private Function RsDr(A As Dao.Recordset) As Variant()
'RsDr = FldsDr(A.Fields)
'End Function
'Private Function RsDry(A As Dao.Recordset) As Variant()
'Dim O()
'With A
'    While Not .EOF
'        Push O, RsDr(A)
'        .MoveNext
'    Wend
'End With
'RsDry = O
'End Function
'Private Function LoHasFny(A As ListObject, Fny$()) As Boolean
'Dim Miss$(), FnyzLo$()
'FnyzLo = LoFny(A)
'Miss = AyMinus(Fny, FnyzLo)
'If Si(Miss) > 0 Then Exit Function
'LoHasFny = True
'End Function
'Private Function WsFstLo(A As Worksheet) As ListObject
'Set WsFstLo = ItrFstItm(A.ListObjects)
'End Function
'Private Function ItrFstItm(A)
'Dim I
'For Each I In A
'    Asg I, ItrFstItm
'Next
'End Function
'Private Function DrsNRow&(A As Drs)
'DrsNRow = Si(A.Dry)
'End Function
'Private Function SqAddSngQuote(A)
'Dim NC%, C%, R&, O
'O = A
'NC = UBound(A, 2)
'For R = 1 To UBound(A, 1)
'    For C = 1 To NC
'        If IsStr(O(R, C)) Then
'            O(R, C) = "'" & O(R, C)
'        End If
'    Next
'Next
'SqAddSngQuote = O
'End Function
'Private Function RsSq(A As Dao.Recordset) As Variant()
'Dim O(), R&, NR&, NC&, C&
'NR = A.RecordCount
'NC = A.Fields.Count
'If NR = 0 Then Stop
'ReDim O(1 To NR, 1 To NC)
'For R = 1 To NR
'    For C = 1 To NC
'        O(R, C) = A.Fields(C - 1).Value
'    Next
'    A.MoveNext
'Next
'RsSq = O
'End Function
'Private Sub DbtPutLo(A as Database, T, Lo As ListObject)
'Dim Sq(), Drs As Drs, Rs As Dao.Recordset
'Set Rs = RszT(A, T)
'If Not AyIsEq(RsFny(Rs), LoFny(Lo)) Then
'    Debug.Print "--"
'    Debug.Print "Rs"
'    Debug.Print "--"
'    D RsFny(Rs)
'    Debug.Print "--"
'    Debug.Print "Lo"
'    Debug.Print "--"
'    D LoFny(Lo)
'    Stop
'End If
'Sq = SqAddSngQuote(RsSq(Rs))
'LoMin Lo
'SqPutAt Sq, Lo.DataBodyRange
'End Sub
'Private Sub LoEnsNRow(A As ListObject, NRow&)
'LoMin A
'Exit Sub
'If NRow > 1 Then
'    Debug.Print A.InsertRowRange.Address
'    Stop
'End If
'End Sub
'Private Function DrsCol(A As Drs, F) As Variant()
'DrsCol = DrsColInto(A, F, EmpAy)
'End Function
'Private Function AyIx&(A, M)
'Dim J&
'For J = 0 To UB(A)
'    If A(J) = M Then AyIx = J: Exit Function
'Next
'AyIx = -1
'End Function
'Private Function LoSy(A As ListObject, ColNm$) As String()
'Dim Sq()
'Sq = A.ListColumns(ColNm).DataBodyRange.Value
'LoSy = SqColSy(Sq(), 1)
'End Function
'Private Function LoFny(A As ListObject) As String()
'LoFny = ItrNy(A.ListColumns)
'End Function
'Private Sub AyPutLoCol(A, Lo As ListObject, ColNm$)
'Dim At As Range, C As ListColumn, R As Range
''D LoFny(Lo)
''Stop
'Set C = Lo.ListColumns(ColNm)
'Set R = C.DataBodyRange
'Set At = R.Cells(1, 1)
'AyPutCol A, At
'End Sub
'Private Function AySqH(A) As Variant()
'Dim O(), N&, J&
'N = Si(A)
'If N = 0 Then Exit Function
'ReDim Sq(1 To 1, 1 To N)
'For J = 1 To N
'    O(1, J) = A(J - 1)
'Next
'AySqH = O
'End Function
'Private Function AySqV(A) As Variant()
'Dim O(), N&, J&
'N = Si(A)
'If N = 0 Then Exit Function
'ReDim O(1 To N, 1 To 1)
'For J = 1 To N
'    O(J, 1) = A(J - 1)
'Next
'AySqV = O
'End Function
'Private Sub AyPutCol(A, At As Range)
'Dim Sq()
'Sq = AySqV(A)
'RgReSz(At, Sq).Value = Sq
'End Sub
'Private Sub AyPutRow(A, At As Range)
'Dim Sq()
'Sq = AySqH(A)
'RgReSz(At, Sq).Value = Sq
'End Sub
'Private Function DrsColInto(A As Drs, F, OInto)
'Dim O, Ix%, Dry(), Dr
'Ix = AyIx(A.Fny, F): If Ix = -1 Then Stop
'O = OInto
'Erase O
'Dry = A.Dry
'If Si(Dry) = 0 Then DrsColInto = O: Exit Function
'For Each Dr In Dry
'    Push O, Dr(Ix)
'Next
'DrsColInto = O
'End Function
'Private Function DrsColSy(A As Drs, F) As String()
'DrsColSy = DrsColInto(A, F, EmpSy)
'End Function
'Private Sub DbOupWb(A as Database, Wb As Workbook, OupNmSsl$)
''OupNm is used for Table-Name-@*, WsCdNm-Ws*, LoNm-Tbl*
'Dim Ay$(), OupNm
'Ay = SslSy(OupNmSsl)
'WbVdtOupNy Wb, Ay
'Dim T$
'For Each OupNm In Ay
'    T = "@" & OupNm
'    DbtOupWb A, T, Wb, OupNm
'Next
'End Sub
'
'Private Sub DbtOupWb(A as Database, T, Wb As Workbook, OupNm)
''OupNm is used for WsCdNm-Ws*, LoNm-Tbl*
'Dim Ws As Worksheet
'Set Ws = WbWsCd(Wb, "WsO" & OupNm)
'DbtPutWs A, T, Ws
'End Sub
'
'Private Sub DbtDrp(A as Database, Tny0)
'Dim Tny$(), T
'Tny = DftNy(Tny0)
'If Si(Tny) = 0 Then Exit Sub
'For Each T In Tny
'    If DbHasTbl(A, T) Then A.Execute FmtQQ("Drop Table [?]", T)
'Next
'End Sub
'
'Private Function DbtLnk(A as Database, T$, S$, Cn$) As String()
'On Error GoTo X
'Dim TT As New Dao.TableDef
'DbDrpTbl A, T
'With TT
'    .Connect = Cn
'    .Name = T
'    .SourceTableName = S
'    A.TableDefs.Append TT
'End With
'Exit Function
'X:
'Debug.Print Err.Description
'Dim O$(), M$
'M = "Cannot create Table in Database from Source by Cn with Er from system"
'Push O, "Program  : DbtLnk"
'Push O, "Database : " & A.Name
'Push O, "Table    : " & T
'Push O, "Source   : " & S
'Push O, "Cn       : " & Cn
'Push O, "Er       : " & Err.Description
'PushMsgUnderLin O, M
'DbtLnk = O
'End Function
'Private Function TblLnk(T$, S$, Cn$) As String()
'TblLnk = DbtLnk(CurrentDb, T, S, Cn)
'End Function
'Private Function WbWsCd(A As Workbook, WsCdNm$) As Worksheet
'Set WbWsCd = ItrFstPrpEq(A.Sheets, "CodeName", WsCdNm)
'End Function
'Private Function WbLasWs(A As Workbook) As Worksheet
'Set WbLasWs = A.Sheets(A.Sheets.Count)
'End Function
'Private Function WbWs(A As Workbook, WsNm$) As Worksheet
'Set WbWs = A.Sheets(WsNm)
'End Function
'Private Function FxWb(A) As Workbook
'Set FxWb = Xls.Workbooks.Open(A)
'End Function
'Private Function WszLo(A As Worksheet, LoNm$) As ListObject
'Set WszLo = A.ListObjects(LoNm)
'End Function
'Private Function TblPutAt(A, At As Range) As Range
'Set TblPutAt = DbtPutAt(CurrentDb, A, At)
'End Function
'Private Function DbtPutAt(A as Database, T, At As Range) As Range
'Set DbtPutAt = SqPutAt(DbtSq(A, T), At)
'End Function
'Private Function AyAddAp(ParamArray Ap())
'Dim Av(), O, J%
'O = Ap(0)
'Av = Ap
'For J = 1 To UB(Av)
'    PushAy O, Av(J)
'Next
'AyAddAp = O
'End Function
'Private Function AlignL$(A, W%)
'AlignL = A & Space(W - Len(A))
'End Function
'
'Private Function AyMapXPSy(A, MapXPFunNm$, P) As String()
'AyMapXPSy = AyMapXPInto(A, MapXPFunNm, P, EmpSy)
'End Function
'
'Private Function AyMapXPInto(A, MapXPFunNm$, P, OInto)
'Dim O, J&
'O = OInto
'Erase O
'If Si(A) = 0 Then AyMapXPInto = O: Exit Function
'ReDim O(UB(A))
'For J = 0 To UB(A)
'    Asg Run(MapXPFunNm, A(J), P), O(J)
'Next
'AyMapXPInto = O
'End Function
'
'Private Function AyAlignL(A) As String()
'AyAlignL = AyMapXPSy(A, "AlignL", WdtzSy(A))
'End Function
'Private Function LnkSpec_LnkColStr$(A)
'Dim L$
'LnkSpec_Asg A, , L
'LnkSpec_LnkColStr = L
'End Function
'Private Function LnkColAy_ImpSql$(A() As LnkCol, T, Optional WhBExpr$)
'If FstChr(T) <> ">" Then
'    Debug.Print "T must have first char = '>'"
'    Stop
'End If
'Dim Ny$(), ExtNy$(), J%, O$(), S$, N$(), E$()
'Ny = LnkColAy_Ny(A)
'ExtNy = LnkColAy_ExtNy(A)
'N = AyAlignL(Ny)
'E = AyAlignL(SyQuoteSqBkt(ExtNy))
'Erase O
'For J = 0 To UB(Ny)
'    If ExtNy(J) = Ny(J) Then
'        Push O, FmtQQ("     ?    ?", Space(Len(E(J))), N(J))
'    Else
'        Push O, FmtQQ("     ? As ?", E(J), N(J))
'    End If
'Next
'S = Join(O, "," & vbCrLf)
'LnkColAy_ImpSql = FmtQQ("Select |?| Into [#I?]| From [?] |?", S, RmvFstChr(T), T, SqpWhere(WhBExpr))
'End Function
'Private Sub WbMinLo(A As Workbook)
'ItrDo A.Sheets, "WsMinLo"
'End Sub
'Private Sub WsMinLo(A As Worksheet)
'If A.CodeName = "WsIdx" Then Exit Sub
'ItrDo A.ListObjects, "LoMin"
'End Sub
'Private Function ZZLo() As ListObject
'Set ZZLo = SqLo(TblSq("MB52"))
'End Function
'Private Sub ZZ_LoMin()
'LoMin ZZLo
'End Sub
'Private Sub LoMin(A As ListObject)
'Dim R1 As Range, R2 As Range
'Set R1 = A.DataBodyRange
'If R1.Rows.Count >= 2 Then
'    Set R2 = RgRR(R1, 2, R1.Rows.Count)
'    R2.Delete
'End If
'End Sub
'Private Function RgRR(A As Range, R1, R2) As Range
'Set RgRR = RgCRR(A, 1, R1, R2).EntireRow
'End Function
'Private Sub FxMinLo(A)
'Dim Wb As Workbook
'Set Wb = FxWb(A)
'WbMinLo Wb
'Wb.Save
'Wb.Close
'End Sub
'Private Sub PcRfh(A As PivotCache)
'A.MissingItemsLimit = xlMissingItemsNone
'A.Refresh
'End Sub
'Private Sub ItrDo(A, DoNm$)
'Dim I
'For Each I In A
'    Run DoNm, I
'Next
'End Sub
'Private Sub ItrDoXP(A, DoXPNm$, P)
'Dim I
'For Each I In A
'    Run DoXPNm, I, P
'Next
'End Sub
'Private Sub WbVis(A As Workbook)
'VisXls A.Application
'End Sub
'
'Private Sub WbRfh(A As Workbook, Fb$)
'ItrDoXP A.Connections, "RfhWc", Fb
'ItrDo A.PivotCaches, "PcRfh"
'ItrDo A.Sheets, "RfhWs"
''ItrDo A.Connections, "DltWc"
'End Sub
'
'Private Sub ZZ_RplBet()
'Dim A$, Exp$, By$, s1$, s2$
's1 = "Data Source="
's2 = ";"
'A = "aa;Data Source=???;klsdf"
'By = "xx"
'Exp = "aa;Data Source=xx;klsdf"
'GoSub Tst
'Exit Sub
'Tst:
'Dim Act$
'Act = RplBet(A, By, s1, s2)
'Debug.Assert Exp = Act
'Return
'End Sub
'Private Function RplBet$(A, By$, s1$, s2$)
'Dim P1%, P2%, B$, C$
'
'P1 = InStr(A, s1)
'If P1 = 0 Then Stop
'P2 = InStr(P1 + Len(s1), CStr(A), s2)
'If P2 = 0 Then Stop
'B = Left(A, P1 + Len(s1) - 1)
'C = Mid(A, P2 + Len(s2) - 1)
'RplBet = B & By & C
'End Function
'
'Private Function FbWbCnStr$(A)
'FbWbCnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
'End Function
'Private Sub RfhWcCnStr(A As WorkbookConnection, Fb)
'If IsNothing(A.OLEDBConnection) Then Exit Sub
'Dim Cn$
'Const Ver$ = "0.0.1"
'Select Case Ver
'Case "0.0.1"
'    Dim S$
'    S = A.OLEDBConnection.Connection
'    Cn = RplBet(S, CStr(Fb$), "Data Source=", ";")
'Case "0.0.2"
'    Cn = FbWbCnStr(Fb$)
'End Select
'A.OLEDBConnection.Connection = Cn
'End Sub
'Private Sub RfhWc(A As WorkbookConnection, Fb)
'If IsNothing(A.OLEDBConnection) Then Exit Sub
'RfhWcCnStr A, Fb
'A.OLEDBConnection.BackgroundQuery = False
'A.OLEDBConnection.Refresh
'End Sub
'Private Sub DltWc(A As WorkbookConnection)
'A.Delete
'End Sub
'
'Private Sub RfhWs(A As Worksheet)
'ItrDo A.QueryTables, "RfhQt"
'ItrDo A.PivotTables, "RfhPt"
'End Sub
'
'Private Sub RfhQt(A As Excel.QueryTable)
'A.BackgroundQuery = False
'A.Refresh
'End Sub
'Private Sub RfhPt(A As Excel.PivotTable)
'A.Update
'End Sub
'Private Function LoVis(A As ListObject) As ListObject
'VisXls A.Application
'Set LoVis = A
'End Function
'Private Function WsVis(A As Worksheet)
'VisXls A.Application
'Set WsVis = A
'End Function
'Private Sub VisXls(A As Excel.Application)
'If Not A.Visible Then A.Visible = True
'End Sub
'Private Function SqPutAt(A, At As Range) As Range
'Dim O As Range
'Set O = RgReSz(At, A)
'O.Value = A
'Set SqPutAt = O
'End Function
'Private Function RgWs(A As Range) As Worksheet
'Set RgWs = A.Parent
'End Function
'Private Function RgRC(A As Range, R, C) As Range
'Set RgRC = A.Cells(R, C)
'End Function
'Private Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
'Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
'End Function
'Private Function RgReSz(A As Range, Sq) As Range
'Set RgReSz = RgRCRC(A, 1, 1, UBound(Sq(), 1), UBound(Sq(), 2))
'End Function
'Private Sub ZZ_TblSq()
'Dim A()
'A = TblSq("@Oup")
'Stop
'End Sub
'Private Function NewWb(Optional WsNm$ = "Sheet1") As Workbook
'Dim O As Workbook, Ws As Worksheet
'Set O = NewXls.Workbooks.Add
'Set Ws = WbFstWs(O)
'If Ws.Name <> WsNm Then Ws.Name = WsNm
'Set NewWb = O
'End Function
'Private Function WbFstWs(A As Workbook) As Worksheet
'Set WbFstWs = A.Sheets(1)
'End Function
'Private Function NewWs(Optional WsNm$ = "Sheet") As Worksheet
'Set NewWs = WbFstWs(NewWb(WsNm))
'End Function
'Private Function NewA1(Optional WsNm$ = "Sheet1") As Range
'Set NewA1 = WsA1(NewWs(WsNm))
'End Function
'Private Function SqA1(A, Optional WsNm$ = "Data") As Range
'Dim A1 As Range
'Set A1 = NewA1(WsNm)
'Set SqA1 = SqPutAt(A, A1)
'End Function
'Private Function WsRC(A As Worksheet, R, C) As Range
'Set WsRC = A.Cells(R, C)
'End Function
'Private Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
'Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
'End Function
'Private Function RgA1LasCell(A As Range) As Range
'Dim L As Range, R, C
'Set L = A.SpecialCells(xlCellTypeLastCell)
'R = L.Row
'C = L.Column
'Set RgA1LasCell = WsRCRC(RgWs(A), A.Row, A.Column, R, C)
'End Function
'Private Function RgLo(A As Range, Optional LoNm$) As ListObject
'Dim O As ListObject
'Set O = RgWs(A).ListObjects.Add(xlSrcRange, A, , XlYesNoGuess.xlYes)
'If LoNm <> "" Then O.Name = LoNm
'Set RgLo = O
'End Function
'Private Function SqLo(A, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As ListObject
'Dim R As Range
'Set R = SqA1(A, WsNm)
'Set SqLo = RgLo(R, LoNm)
'End Function
'Private Sub RgVis(A As Range)
'VisXls A.Application
'End Sub
'Private Function DbtPutFx(A as Database, T, Fx$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As Workbook
'Dim O As Workbook, Ws As Worksheet
'Set O = FxWb(Fx$)
'Set Ws = WbWs(O, WsNm)
'WsClrLo Ws
'Stop ' LoNm need handle?
'DbtPutWs A, T, WbWs(O, WsNm)
'Set DbtPutFx = O
'End Function
'Private Sub WsClrLo(A As Worksheet)
'Dim Ay() As ListObject, J%
'Ay = ItrInto(A.ListObjects, Ay)
'For J = 0 To UB(Ay)
'    Ay(J).Delete
'Next
'End Sub
'Private Function TblPutFx(T, Fx$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As Workbook
'Set TblPutFx = DbtPutFx(CurrentDb, T, Fx, WsNm, LoNm)
'End Function
'Private Function AddWs(A As Workbook, Optional WsNm$, Optional BefWsNm$, Optional AftWsNm$) As Worksheet
'Dim O As Worksheet, Bef As Worksheet, Aft As Worksheet
'WbDltWs A, WsNm
'Select Case True
'Case BefWsNm <> ""
'    Set Bef = A.Sheets(BefWsNm)
'    Set O = A.Sheets.Add(Bef)
'Case AftWsNm <> ""
'    Set Aft = A.Sheets(AftWsNm)
'    Set O = A.Sheets.Add(, Aft)
'Case Else
'    Set O = A.Sheets.Add
'End Select
'O.Name = WsNm
'Set AddWs = O
'End Function
'Private Sub WbDltWs(A As Workbook, WsNm$)
'If WbHasWs(A, WsNm) Then
'    A.Application.DisplayAlerts = False
'    WbWs(A, WsNm).Delete
'    A.Application.DisplayAlerts = True
'End If
'End Sub
'Private Function ItrHasNm(A, Nm$) As Boolean
'Dim I
'For Each I In A
'    If I.Name = Nm Then ItrHasNm = True: Exit Function
'Next
'End Function
'Private Function WbHasWs(A As Workbook, WsNm$) As Boolean
'WbHasWs = ItrHasNm(A.Sheets, WsNm)
'End Function
'Private Sub FfnCpy(A, ToFfn$, Optional OvrWrt As Boolean)
'If OvrWrt Then FfnDlt ToFfn
'FileSystem.FileCopy A, ToFfn
'End Sub
'Private Sub FfnDlt(A)
'If FfnIsExist(A) Then Kill A
'End Sub
'Private Function PthIsExist(A) As Boolean
'On Error Resume Next
'PthIsExist = Dir(A, vbDirectory) <> ""
'End Function
'Private Function FfnIsExist(A) As Boolean
'On Error Resume Next
'FfnIsExist = Dir(A) <> ""
'End Function
'Private Sub TblPutWs(T, Ws As Worksheet, Optional LoNm$)
'RgLo TblPutAt(T, WsA1(Ws)), LoNm
'End Sub
'Private Function DbqSy(A as Database, Sql) As String()
'DbqSy = RsSy(A.OpenRecordset(Sql))
'End Function
'Private Function DbStru(A as Database, Optional Tny0) As String()
'DbStru = DbtStru(A, Tny0)
'End Function
'Private Function DbTny(A As Database) As String()
'DbTny = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'")
'Exit Function
'Dim T As TableDef, O$()
'Dim X As Dao.TableDefAttributeEnum
'X = Dao.TableDefAttributeEnum.dbHiddenObject Or Dao.TableDefAttributeEnum.dbSystemObject
'For Each T In A.TableDefs
'    Select Case True
'    Case T.Attributes And X
'    Case Else
'        Push O, T.Name
'    End Select
'Next
'DbTny = O
'End Function
'Private Function IsPfx(A$, Pfx$) As Boolean
'IsPfx = Left(A, Len(Pfx)) = Pfx
'End Function
'Private Function DbtNRec&(A as Database, T)
'DbtNRec = DbqV(A, FmtQQ("Select Count(*) from [?]", T))
'End Function
'Private Function DbtCsv(A as Database, T$) As String()
'DbtCsv = RsCsvLy(RszT(A, T))
'End Function
'Private Sub DbtPutWs(A as Database, T, Ws As Worksheet)
''Assume the WsCdNm is WsXXX and there will only 1 Lo with Name TblXXX
''Else stop
'Dim Lo As ListObject
'Set Lo = WsFstLo(Ws)
'
'If Not IsPfx(Ws.CodeName, "WsO") Then Stop
'If Ws.ListObjects.Count <> 1 Then Stop
'If Mid(Lo.Name, 4) <> Mid(Ws.CodeName, 4) Then Stop
'DbtPutLo A, T, Lo
'End Sub
'Private Function DSpecNm$(A)
'DSpecNm = TakAftDotOrAll(LinT1(A))
'End Function
'Private Function TakAftDotOrAll$(A)
'TakAftDotOrAll = TakAftOrAll(A, ".")
'End Function
'Private Function TblWs(T, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As Worksheet
'Set TblWs = WszLo(SqLo(TblSq(T), WsNm, LoNm))
'End Function
'Private Function WszLo(A As ListObject) As Worksheet
'Set WszLo = A.Parent
'End Function
'Private Function TblRs(T) As Dao.Recordset
'Set TblRs = RszT(CurrentDb, T)
'End Function
'Private Sub TimFn(FnNm$)
'Dim A!, B!
'A = Timer
'Run FnNm
'B = Timer
'Debug.Print FnNm, B - A
'End Sub
'Private Function RsCsvLy(A As Dao.Recordset) As String()
'Dim O$(), J&, I%, UFld%, Dr(), F As Dao.Field
'UFld = A.Fields.Count - 1
'While Not A.EOF
'    J = J + 1
'    If J Mod 5000 = 0 Then Debug.Print "RsCsvLy: " & J
'    If J > 100000 Then Stop
'    ReDim Dr(UFld)
'    I = 0
'    For Each F In A.Fields
'        Dr(I) = VarCsv(F.Value)
'        I = I + 1
'    Next
'    Push O, Join(Dr, ",")
'    A.MoveNext
'Wend
'RsCsvLy = O
'End Function
'
'Private Function TblNRow&(T, Optional WhBExpr$)
'TblNRow = DbtNRow(CurrentDb, T, WhBExpr)
'End Function
'Private Function SqpWhere$(WhBExpr$)
'If WhBExpr = "" Then Exit Function
'SqpWhere = " Where " & WhBExpr
'End Function
'Private Function DbtNRow&(A as Database, T, Optional WhBExpr$)
'Dim S$
'S = FmtQQ("Select Count(*) from [?]?", T, SqpWhere(WhBExpr))
'DbtNRow = DbqLng(A, S)
'End Function
'Private Function TblNCol&(T)
'TblNCol = DbtNCol(CurrentDb, T)
'End Function
'Private Function DbtNCol&(A as Database, T)
'DbtNCol = A.OpenRecordset(T).Fields.Count
'End Function
'Private Function TblSq(A) As Variant()
'TblSq = DbtSq(CurrentDb, A)
'End Function
'Private Function DbtSq(A as Database, T$) As Variant()
'Dim NR&, NC&, Rs As Dao.Recordset
'Dim O(), J&
'NR = DbtNRow(A, T)
'NC = DbtNCol(A, T)
'Set Rs = RszT(A, T)
'ReDim O(1 To NR + 1, 1 To NC)
'With Rs
'    DrPutSq ItrNy(.Fields), O
'    J = 2
'    While Not .EOF
'        RsPutSq Rs, O, J
'        J = J + 1
'        .MoveNext
'    Wend
'    .Close
'End With
'DbtSq = O
'End Function
'Private Function FxWs(A, Optional WsNm$ = "Data") As Worksheet
'Set FxWs = WbWs(FxWb(A), WsNm)
'End Function
'Private Sub FldsPutSq(A As Dao.Fields, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
'Dim F As Dao.Field, J%
'If NoTxtSngQ Then
'    For Each F In A
'        J = J + 1
'        Sq(R, J) = F.Value
'    Next
'    Exit Sub
'End If
'For Each F In A
'    J = J + 1
'    If F.Type = Dao.DataTypeEnum.dbText Then
'        Sq(R, J) = "'" & F.Value
'    Else
'        Sq(R, J) = F.Value
'    End If
'Next
'End Sub
'Private Sub DrPutSq(A, Sq, Optional R& = 1)
'Dim J%, I
'For Each I In A
'    J = J + 1
'    Sq(R, J) = I
'Next
'End Sub
'Private Sub RsPutSq(A As Dao.Recordset, Sq, R&, Optional NoTxtSngQ As Boolean)
'FldsPutSq A.Fields, Sq, R, NoTxtSngQ
'End Sub
'Private Function WsRCC(A As Worksheet, R, C1, C2) As Range
'Set WsRCC = WsRCRC(A, R, C1, R, C2)
'End Function
'Private Function WsCC(A As Worksheet, C1, C2) As Range
'Set WsCC = WsRCC(A, 1, C1, C2).EntireColumn
'End Function
'Private Function WsRR(A As Worksheet, R1&, R2&) As Range
'Set WsRR = A.Rows(R1 & ":" & R2)
'End Function
'Private Function WsA1(A As Worksheet) As Range
'Set WsA1 = A.Cells(1, 1)
'End Function
'Private Function FxLo(A$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As ListObject
'Set FxLo = WszLo(WbWs(FxWb(A), WsNm), LoNm)
'End Function
'Private Function TblCnStr$(T)
'TblCnStr = CurrentDb.TableDefs(T).Connect
'End Function
'Private Function DbqLng&(A as Database, Sql)
'DbqLng = DbqV(A, Sql)
'End Function
'Private Function SqlLng&(A)
'SqlLng = DbqLng(CurrentDb, A)
'End Function
'Private Function SqlV(A)
'SqlV = DbqV(CurrentDb, A)
'End Function
'Private Function DbqV(A as Database, Sql)
'DbqV = A.OpenRecordset(Sql).Fields(0).Value
'End Function
'Private Function TblNRec&(A)
'TblNRec = SqlLng(FmtQQ("Select Count(*) from [?]", A))
'End Function
'Private Function ErzFileNotFound(FfnAy0) As String()
'Dim Ay$(), I, O$()
'Ay = DftFfnAy(FfnAy0)
'If Si(Ay) = 0 Then Exit Function
'For Each I In Ay
'    If Not FfnIsExist(I) Then
'        Push O, "File: " & I
'        PushMsgUnderLin O, "Above file not found"
'    End If
'Next
'ErzFileNotFound = O
'End Function
'Private Function DbtLnkFx(A as Database, T$, Fx$, Optional WsNm$ = "Sheet1") As String()
'Dim O$()
'O = ErzFileNotFound(Fx$)
'If Si(O) > 0 Then
'    DbtLnkFx = O
'    Exit Function
'End If
'Dim Cn$: Cn = FxDaoCnStr(Fx$)
'Dim Src$: Src = WsNm & "$"
'DbtLnkFx = DbtLnk(A, T, Src, Cn)
'End Function
'Private Function TblLnkFb(Tny0, Fb$, Optional FbTny0) As String()
'TblLnkFb = DbtLnkFb(CurrentDb, Tny0, Fb, FbTny0)
'End Function
'Private Function DbtLnkFb(A as Database, Tny0, Fb$, Optional FbTny0) As String()
'Dim Tny$(), FbTny$()
'Tny = DftNy(Tny0)
'FbTny = DftNy(FbTny0)
'    Select Case True
'    Case Si(FbTny) = Si(Tny)
'    Case Si(FbTny) = 0
'        FbTny = Tny
'    Case Else
'        Stop
'    End Select
'Dim Cn$: Cn = FbCnStr(Fb$)
'Dim J%, O$()
'For J = 0 To UB(Tny)
'    O = AyAdd(O, DbtLnk(A, Tny(J), FbTny(J), Cn))
'Next
'DbtLnkFb = O
'End Function
'Private Function TblLnkFx(T$, Fx$, Optional WsNm$ = "Sheet1") As String()
'TblLnkFx = DbtLnkFx(CurrentDb, T, Fx, WsNm)
'End Function
'Private Function FbCnStr$(A)
'FbCnStr = ";DATABASE=" & A & ";"
'End Function
'
'Private Function AyHas(A, M) As Boolean
'Dim I
'If Si(A) = 0 Then Exit Function
'For Each I In A
'    If I = M Then
'        AyHas = True
'        Exit Function
'    End If
'Next
'End Function
'
'Private Function SyQuoteSqBkt(A) As String()
'SyQuoteSqBkt = SyQuote(A, "[]")
'End Function
'Private Function DbtPk(A as Database, T$) As String()
'
'End Function
'Private Function SyQuoteSng(A) As String()
'SyQuoteSng = SyQuote(A, "'")
'End Function
'Private Function DbtStru(A as Database, Tny0) As String()
'Dim Tny$()
'Tny = DftNy(Tny0)
'Select Case Si(Tny)
'Case 0:
'    DbtStru = DbtStru(A, DbTny(A))
'Case 1:
'    DbtStru = ApSy(Tny(0) & ": " & JnSpc(DbtFny(A, Tny(0))))
'Case Else
'    Dim O$(), T
'    For Each T In Tny
'        PushAy O, DbtStru(A, T)
'    Next
'    DbtStru = O
'    Exit Function
'End Select
'End Function
'Private Sub DbtfChgDteToTxt(A as Database, T, F)
'A.Execute FmtQQ("Alter Table [?] add column [###] text(12)", T)
'A.Execute FmtQQ("Update [?] set [###] = Format([?],'YYYY-MM-DD')", T, F)
'A.Execute FmtQQ("Alter Table [?] Drop Column [?]", T, F)
'A.Execute FmtQQ("Alter Table [?] Add Column [?] text(12)", T, F)
'A.Execute FmtQQ("Update [?] set [?] = [###]", T, F)
'A.Execute FmtQQ("Alter Table [?] Drop Column [###]", T)
'End Sub
'Private Function JnComma$(A)
'JnComma = Join(A, ",")
'End Function
'Private Function JnSpc$(A)
'JnSpc = Join(A, " ")
'End Function
'Private Function UB&(A)
'UB = Si(A) - 1
'End Function
'
'Private Sub PushNonEmpty(O, A)
'If A = "" Then Exit Sub
'Push O, A
'End Sub
'Private Function DaoTy_Str$(T As Dao.DataTypeEnum)
'Dim O$
'Select Case T
'Case Dao.DataTypeEnum.dbBoolean: O = "Boolean"
'Case Dao.DataTypeEnum.dbDouble: O = "Double"
'Case Dao.DataTypeEnum.dbText: O = "Text"
'Case Dao.DataTypeEnum.dbDate: O = "Date"
'Case Dao.DataTypeEnum.dbByte: O = "Byte"
'Case Dao.DataTypeEnum.dbInteger: O = "Int"
'Case Dao.DataTypeEnum.dbLong: O = "Long"
'Case Dao.DataTypeEnum.dbDouble: O = "Doubld"
'Case Dao.DataTypeEnum.dbDate: O = "Date"
'Case Dao.DataTypeEnum.dbDecimal: O = "Decimal"
'Case Dao.DataTypeEnum.dbCurrency: O = "Currency"
'Case Dao.DataTypeEnum.dbSingle: O = "Single"
'Case Else: Stop
'End Select
'DaoTy_Str = O
'End Function
'Private Function DbqryRs(A as Database, Q) As Dao.Recordset
'Set DbqryRs = A.QueryDefs(Q).OpenRecordset
'End Function
'Private Function RplVBar$(A)
'RplVBar = Replace(A, "|", vbCrLf)
'End Function
'
'Private Function AyBrwEr(A) As Boolean
'If Si(A) = 0 Then Exit Function
'AyBrwEr = True
'AyBrw A
'End Function
'Private Sub AyBrw(A)
'StrBrw Join(A, vbCrLf)
'End Sub
'Private Function TblFld_Ty(T, F) As Dao.DataTypeEnum
'TblFld_Ty = CurrentDb.TableDefs(T).Fields(F).Type
'End Function
'
'Private Sub StrWrt(A, Ft$, Optional IsNotOvrWrt As Boolean)
'Fso.CreateTextFile(Ft$, Overwrite:=Not IsNotOvrWrt).Write A
'End Sub
'Private Sub FtBrw(A)
''Shell "code.cmd """ & A & """", vbHide
'Shell "notepad.exe """ & A & """", vbMaximizedFocus
'End Sub
'Private Function JnCrLf$(A)
'JnCrLf = Join(A, vbCrLf)
'End Function
'Private Sub AyWrt(A, Ft$)
'StrWrt JnCrLf(A), Ft
'End Sub
'
'Private Sub StrBrw(A)
'Dim T$
'T = TmpFt
'StrWrt A, T
'FtBrw T
'End Sub
'
'Private Function TmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
'Dim Fnn$
'If Fnn0 = "" Then
'    Fnn = TmpNm
'Else
'    Fnn = Fnn0
'End If
'TmpFfn = TmpPth(Fdr) & Fnn & Ext
'End Function
'
'Private Function TmpFt$(Optional Fdr$, Optional Fnn$)
'TmpFt = TmpFfn(".txt", Fdr, Fnn)
'End Function
'Private Function TmpFx$(Optional Fdr$, Optional Fnn$)
'TmpFx = TmpFfn(".xlsx", Fdr, Fnn)
'End Function
'
'Private Function TmpNm$()
'Static X&
'TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
'X = X + 1
'End Function
'
'Private Function TmpPth$(Optional Fdr$)
'Dim X$
'   If Fdr <> "" Then
'       X = Fdr & "\"
'   End If
'Dim O$
'   O = TmpPthHom & X:   EnsPth O
'   O = O & TmpNm & "\": EnsPth O
'   EnsPth O
'TmpPth = O
'End Function
'
'Private Function UpdEndDte__1(A as Database, T, KeyFld$, FmDteFld$) As Date()
'Dim K$(), FmDte() As Date, ToDte() As Date, J&, CurKey$, NxtKey$, NxtFmDte As Date
'With RszT(A, T)
'    While Not .EOF
'        Push FmDte, .Fields(FmDteFld).Value
'        Push K, .Fields(KeyFld).Value
'        .MoveNext
'    Wend
'End With
'Dim U&
'U = UB(K)
'ReDim ToDte(U)
'For J = 0 To U - 1
'    CurKey = K(J)
'    NxtKey = K(J + 1)
'    NxtFmDte = FmDte(J + 1)
'    If CurKey = NxtKey Then
'        ToDte(J) = DateAdd("D", -1, NxtFmDte)
'    Else
'        ToDte(J) = DateSerial(2099, 12, 31)
'    End If
'Next
'ToDte(U) = DateSerial(2099, 12, 31)
'UpdEndDte__1 = ToDte
'End Function
'Private Sub ZZ_UpdEndDte()
'DoCmd.RunSQL "Select * into [#A] from ZZ_UpdEndDte order by Sku,PermitDate"
'UpdEndDte CurrentDb, "#A", "PermitDateEnd", "Sku", "PermitDate"
'Stop
'DrpT CurrentDb, "#A"
'End Sub
'
'Private Sub UpdEndDte(A as Database, T, ToDteFld$, KeyFld$, FmDteFld$)
'Dim ToDte() As Date, J&
'ToDte = UpdEndDte__1(A, T, KeyFld, FmDteFld)
'With RszT(A, T)
'    While Not .EOF
'        .Edit
'        .Fields(ToDteFld).Value = ToDte(J): J = J + 1
'        .Update
'        .MoveNext
'    Wend
'    .Close
'End With
'End Sub
'
'Private Function LinT1$(A)
'LinT1 = LinShiftTerm(CStr(A))
'End Function
'
'Private Property Get TblImpSpec(T$, LnkSpec$, Optional WhBExpr$) As TblImpSpec
'Dim O As New TblImpSpec
'Set TblImpSpec = O.Init(T, LnkSpec$, WhBExpr)
'End Property
'
'Private Function TmpPthHom$()
'Static X$
'If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
'TmpPthHom = X
'End Function
'
'Private Function FmtQQ$(QQVbl$, ParamArray Ap())
'Dim Av(): Av = Ap
'FmtQQ = FmtQQAv(QQVbl, Av)
'End Function
'
'Private Function SqlDry(A) As Variant()
'Dim O(), Rs As Dao.Recordset
'Set Rs = CurrentA.OpenRecordset(A)
'With Rs
'    While Not .EOF
'        Push O, FldsDr(Rs.Fields)
'        .MoveNext
'    Wend
'    .Close
'End With
'SqlDry = O
'End Function
'Private Function Xls(Optional Vis As Boolean) As Excel.Application
'Static X As Boolean, Y As Excel.Application
'Dim J%
'Beg:
'    J = J + 1
'    If J > 10 Then Stop
'If Not X Then
'    X = True
'    Set Y = New Excel.Application
'End If
'On Error GoTo xx
'Dim A$
'A = Y.Name
'Set Xls = Y
'If Vis Then VisXls Y
'Exit Function
'xx:
'    X = True
'    GoTo Beg
'End Function
'Private Function DbtPutAtByCn(A as Database, T, At As Range, Optional LoNm0$) As ListObject
'If FstChr(T) <> "@" Then Stop
'Dim LoNm$, Lo As ListObject
'If LoNm0 = "" Then
'    LoNm = "Tbl" & RmvFstChr(T)
'Else
'    LoNm = LoNm0
'End If
'Dim AtA1 As Range, CnStr, Ws As Worksheet
'Set AtA1 = RgRC(At, 1, 1)
'Set Ws = RgWs(At)
'With Ws.ListObjects.Add(SourceType:=0, Source:=Array( _
'        FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share D", A.Name) _
'        , _
'        "eny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Databa" _
'        , _
'        "se Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Je" _
'        , _
'        "t OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Com" _
'        , _
'        "pact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=" _
'        , _
'        "False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
'        ), Destination:=AtA1).QueryTable '<---- At
'        .CommandType = xlCmdTable
'        .CommandText = Array(T) '<-----  T
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
'        .ListObject.DisplayName = LoNm '<------------ LoNm
'        .Refresh BackgroundQuery:=False
'    End With
'
'End Function
'Private Function NewXls(Optional Vis As Boolean) As Excel.Application
'Dim O As New Excel.Application
'If Vis Then O.Visible = True
'Set NewXls = O
'End Function
'Private Function SqlStrCol(A) As String()
'SqlStrCol = RsStrCol(CurrentA.OpenRecordset(A))
'End Function
'Private Sub DicDmp(A As Dictionary)
'Dim K
'For Each K In A
'    Debug.Print K, A(K)
'Next
'End Sub
'
'Private Sub SqlAy_Run(SqlAy$())
'Dim I
'For Each I In SqlAy
'    DoCmd.RunSQL I
'Next
'End Sub
'
'Private Function RsStrCol(A As Dao.Recordset) As String()
'Dim O$()
'With A
'    While Not .EOF
'        Push O, .Fields(0).Value
'        .MoveNext
'    Wend
'End With
'RsStrCol = O
'End Function
'Private Function SqColInto(A, C%, OInto) As String()
'Dim O
'O = OInto
'Erase O
'Dim NR&, J&
'NR = UBound(A, 1)
'ReDim O(NR - 1)
'For J = 1 To NR
'    O(J - 1) = A(J, C%)
'Next
'SqColInto = O
'End Function
'Private Function SqColSy(A, C%) As String()
'SqColSy = SqColInto(A, C, EmpSy)
'End Function
'Private Function AtVBar(A As Range) As Range
'If IsEmpty(A.Value) Then Stop
'If IsEmpty(RgRC(A, 2, 1).Value) Then
'    Set AtVBar = RgRC(A, 1, 1)
'    Exit Function
'End If
'Set AtVBar = RgCRR(A, 1, 1, A.End(xlDown).Row - A.Row + 1)
'End Function
'Private Function RgCRR(A As Range, C, R1, R2) As Range
'Set RgCRR = RgRCRC(A, R1, C, R2, C)
'End Function
'Private Function SqSyV(A) As String()
'SqSyV = SqColSy(A, 1)
'End Function
'Private Sub RgFillCol(A As Range)
'Dim Rg As Range
'Dim Sq()
'Sq = SqzVBar(A.Rows.Count)
'RgReSz(A, Sq).Value = Sq
'End Sub
'Private Sub RgFillRow(A As Range)
'Dim Rg As Range
'Dim Sq()
'Sq = SqzHBar(A.Rows.Count)
'RgReSz(A, Sq).Value = Sq
'End Sub
'Private Function SqzVBar(N%) As Variant()
'Dim O(), J%
'ReDim O(1 To N, 1 To 1)
'For J = 1 To N
'    O(J, 1) = J
'Next
'SqzVBar = O
'End Function
'Private Function SqzHBar(N%) As Variant()
'Dim O(), J%
'ReDim O(1 To 1, 1 To N)
'For J = 1 To N
'    O(1, J) = J
'Next
'SqzHBar = O
'End Function
'Private Sub OpnFx(A)
'If Not FfnIsExist(A) Then
'    MsgBox "File not found: " & vbCrLf & vbCrLf & A
'    Exit Sub
'End If
'Dim C$
'C = FmtQQ("Excel ""?""", A)
'Debug.Print C
'Shell C, vbMaximizedFocus
''Xls(Vis:=True).Workbooks.Open A
'End Sub
'Private Function SyQuote(A, Q$) As String()
'If Si(A) = 0 Then Exit Function
'Dim Q1$, Q2$
'Select Case True
'Case Len(Q) = 1: Q1 = Q: Q2 = Q
'Case Len(Q) = 2: Q1 = Left(Q, 1): Q2 = Right(Q, 1)
'Case Else: Stop
'End Select
'
'Dim I, O$()
'For Each I In A
'    Push O, Q1 & I & Q2
'Next
'SyQuote = O
'End Function
'Private Function CvFld(A) As Dao.Field
'Set CvFld = A
'End Function
'Private Function FldsDr(A As Dao.Fields) As Variant()
'Dim O(), F
'For Each F In A
'    Push O, CvFld(F).Value
'Next
'FldsDr = O
'End Function
'Private Function SubStrCnt%(A, SubStr$)
'Dim J&, O%, P%, L%
'L = Len(SubStr)
'P = InStr(A, SubStr)
'While P > 0
'    O = O + 1
'    J = J + 1: If J > 100000 Then Stop
'    P = InStr(P + L, A, SubStr)
'Wend
'SubStrCnt = O
'End Function
'Private Function RgCC(A As Range, C1, C2) As Range
'Set RgCC = RgRCRC(A, 1, C1, A.Rows.Count, C2)
'End Function
'
'Private Sub ZZ_FmtQQAv()
'Debug.Print FmtQQ("klsdf?sdf?dsklf", 2, 1)
'End Sub
'Private Function FmtQQAv$(QQVbl, Av)
'Dim O$, I, Cnt
'O = Replace(QQVbl, "|", vbCrLf)
'Cnt = SubStrCnt(QQVbl, "?")
'If Cnt <> Si(Av) Then Stop
'For Each I In Av
'    O = Replace(O, "?", I, Count:=1)
'Next
'FmtQQAv = O
'End Function
'Private Sub PushAy(O, A)
'If Si(A) = 0 Then Exit Sub
'Dim I
'For Each I In A
'    Push O, I
'Next
'End Sub
'
'Private Function AyIsEmpty(A) As Boolean
'AyIsEmpty = Si(A) = 0
'End Function
'Private Function FfnNxt$(A)
'If Not FfnIsExist(A) Then FfnNxt = A: Exit Function
'Dim J%, O$
'For J = 1 To 99
'    O = FfnNxtN(A, J)
'    If Not FfnIsExist(O) Then FfnNxt = O: Exit Function
'Next
'Stop
'End Function
'
'Private Function FfnAddFnSfx$(A, Sfx$)
'FfnAddFnSfx = FfnPth(A) & Fnn(A) & Sfx & FfnExt(A)
'End Function
'
'Private Function FfnNxtN$(A, N%)
'If 1 > N Or N > 99 Then Stop
'Dim Sfx$
'Sfx = "(" & Format(N, "00") & ")"
'FfnNxtN = FfnAddFnSfx(A, Sfx)
'End Function
'
'Private Function PthSel$(A, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
'With FileDialog(msoFileDialogFolderPicker)
'    .AllowMultiSelect = False
'    .InitialFileName = Nz(A, "")
'    .Show
'    If .SelectedItems.Count = 1 Then
'        PthSel = EnsPthSfx(.SelectedItems(1))
'    End If
'End With
'End Function
'Private Sub ZZ_PthSel()
'MsgBox FfnSel("C:\")
'End Sub
'Private Function FfnSel$(A, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
'With FileDialog(msoFileDialogFilePicker)
'    .Filters.Clear
'    .Title = Tit
'    .AllowMultiSelect = False
'    .Filters.Add "", FSpec
'    .InitialFileName = A
'    .ButtonName = BtnNm
'    .Show
'    If .SelectedItems.Count = 1 Then
'        FfnSel = .SelectedItems(1)
'    End If
'End With
'End Function
'Private Sub TxtbSelPth(A As Access.TextBox)
'Dim R$
'R = PthSel(A.Value)
'If R = "" Then Exit Sub
'A.Value = R
'End Sub
'Private Function Fn$(A)
'Dim P%: P = InStrRev(A, "\")
'If P = 0 Then Fn = A: Exit Function
'Fn = Mid(A, P + 1)
'End Function
'
'Private Function Fnn$(A)
'Fnn = FfnCutExt(Fn(A))
'End Function
'Private Function FfnCutExt$(A)
'Dim B$, C$, P%
'B = Fn(A)
'P = InStrRev(B, ".")
'If P = 0 Then
'    C = B
'Else
'    C = Left(B, P - 1)
'End If
'FfnCutExt = FfnPth(A) & C
'End Function
'Private Sub EnsPth(A)
'If Dir(A, VbFileAttribute.vbDirectory) = "" Then MkDir A
'End Sub
'
'Private Function PthFfnAy(A, Spec$) As String()
'Dim O$(), B$, P$
'P = EnsPthSfx(A)
'B = Dir(A & Spec)
'Dim J%
'While B <> ""
'    J = J + 1
'    If J > 1000 Then Stop
'    Push O, P & B
'    B = Dir
'Wend
'PthFfnAy = O
'End Function
'
'Private Function FfnExt$(Ffn$)
'Dim P%: P = InStrRev(Ffn$, ".")
'If P = 0 Then Exit Function
'FfnExt = Mid(Ffn$, P)
'End Function
'
'Private Function PthFxAy(A) As String()
'Dim O$(), B$
'If Right(A, 1) <> "\" Then Stop
'B = Dir(A & "*.xls")
'Dim J%
'While B <> ""
'    J = J + 1
'    If J > 1000 Then Stop
'    If FfnExt(B) = ".xls" Then
'        Push O, A & B
'    End If
'    B = Dir
'Wend
'PthFxAy = O
'End Function
'
'Private Function RmvLasChr$(A)
'RmvLasChr = Left(A, Len(A) - 1)
'End Function
'Private Function RmvFstChr$(A)
'RmvFstChr = Mid(A, 2)
'End Function
'
'Private Function AyIsEq(A, B) As Boolean
'Dim U&, J&
'U = UB(A)
'If UB(B) <> U Then Exit Function
'For J = 0 To U
'    If A(J) <> B(J) Then Exit Function
'Next
'AyIsEq = True
'End Function
'Private Function RsIsBrk(A As Dao.Recordset, GpKy$(), LasVy()) As Boolean
'RsIsBrk = Not AyIsEq(RsVy(A, GpKy), LasVy)
'End Function
'Private Function RsVy(A As Dao.Recordset, Optional Ky0) As Variant()
'RsVy = FldsVy(A.Fields, Ky0)
'End Function
'Private Function FldsVyByKy(A As Dao.Fields, Ky$()) As Variant()
'Dim O(), J%, K
'If Si(Ky) = 0 Then
'    FldsVyByKy = ItrVy(A)
'    Exit Function
'End If
'ReDim O(UB(Ky))
'For Each K In Ky
'    O(J) = A(K).Value
'    J = J + 1
'Next
'FldsVyByKy = O
'End Function
'Private Sub ZZ_FldsVy()
'Dim Rs As Dao.Recordset, Vy()
'Set Rs = CurrentA.OpenRecordset("Select * from SkuB")
'With Rs
'    While Not .EOF
'        Vy = RsVy(Rs)
'        Debug.Print JnComma(Vy)
'        .MoveNext
'    Wend
'    .Close
'End With
'End Sub
'Private Function ItrPrpAy(A, PrpNm$) As Variant()
'Dim O(), I
'For Each I In A
'    Push O, CallByName(I, PrpNm, VbGet)
'Next
'ItrPrpAy = O
'End Function
'Private Function ItrVy(A) As Variant()
'ItrVy = ItrPrpAy(A, "Value")
'End Function
'Private Function IsDte(A) As Boolean
'IsDte = VarType(A) = vbDate
'End Function
'Private Function IsStr(A) As Boolean
'IsStr = VarType(A) = vbString
'End Function
'Private Function IsSy(A) As Boolean
'IsSy = VarType(A) = vbString + vbArray
'End Function
'Private Function CvSy(A) As String()
'CvSy = A
'End Function
'Private Function FldsVy(A As Dao.Fields, Optional Ky0) As Variant()
'Select Case True
'Case IsMissing(Ky0)
'    FldsVy = ItrVy(A)
'Case IsStr(Ky0)
'    FldsVy = FldsVyByKy(A, SslSy(Ky0))
'Case IsSy(Ky0)
'    FldsVy = FldsVyByKy(A, CvSy(Ky0))
'Case Else
'    Stop
'End Select
'End Function
'Private Sub ZZ_SslSqBktCsv()
'Debug.Print SslSqBktCsv("a b c")
'End Sub
'Private Function SslSqBktCsv$(A)
'Dim B$(), C$()
'B = SslSy(A)
'C = SyQuoteSqBkt(B)
'SslSqBktCsv = JnComma(C)
'End Function
'Private Function Ny0SqBktCsv$(A)
'Dim B$(), C$()
'B = DftNy(A)
'C = SyQuoteSqBkt(B)
'Ny0SqBktCsv = JnComma(C)
'End Function
'Private Function RsFny(A As Dao.Recordset) As String()
'RsFny = FldsFny(A.Fields)
'End Function
'
'Private Function AyHasAy(A, Ay) As Boolean
'Dim I
'For Each I In Ay
'    If Not AyHas(A, I) Then Exit Function
'Next
'AyHasAy = True
'End Function
'
'Private Function SqlQQStr_Sy(Sql$, QQStr$) As String()
'Dim Dry: Dry = SqlDry(Sql)
'If AyIsEmpty(Dry) Then Exit Function
'Dim O$()
'Dim Dr
'For Each Dr In Dry
'    Push O, FmtQQAv(QQStr, Dr)
'Next
'SqlQQStr_Sy = O
'End Function
'
'
'Private Function FldsCsv$(A As Dao.Fields)
'FldsCsv = AyCsv(ItrVy(A))
'End Function
'Private Function VarCsv$(A)
'Select Case True
'Case IsStr(A): VarCsv = """" & A & """"
'Case IsDte(A): VarCsv = Format(A, "YYYY-MM-DD HH:MM:SS")
'Case Else: VarCsv = Nz(A, "")
'End Select
'End Function
'Private Function AyMapInto(A, MapFunNm$, OInto)
'Dim J&, O, I, U&
'O = OInto
'Erase O
'U = UB(A)
'If U = -1 Then
'    AyMapInto = O
'    Exit Function
'End If
'ReDim O(U)
'For Each I In A
'    Asg Run(MapFunNm, I), O(J)
'    J = J + 1
'Next
'AyMapInto = O
'End Function
'Private Sub Asg(Fm, OTo)
'If IsObject(Fm) Then
'    Set OTo = Fm
'Else
'    OTo = Nz(Fm, "")
'End If
'End Sub
'Private Function AyMapSy(A, MapFunNm$) As String()
'AyMapSy = AyMapInto(A, MapFunNm, EmpSy)
'End Function
'Private Function AyCsv$(A)
'AyCsv = Join(A, ",")
'Exit Function
'Dim J%
'For J = 0 To UB(A)
'    A(J) = VarCsv(A(J))
'Next
'AyCsv = Join(A, ",")
'End Function
'Private Sub ZZ_DbtUpdSeq()
'DoCmd.SetWarnings False
'DoCmd.RunSQL "Select * into [#A] from ZZ_DbtUpdSeq order by Sku,PermitDate"
'DoCmd.RunSQL "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
'DbtUpdSeq CurrentDb, "#A", "BchRateSeq", "Sku", "Sku Rate"
'TblOpn "#A"
'Stop
'DoCmd.RunSQL "Drop Table [#A]"
'End Sub
'Private Sub DbtUpdSeq(A as Database, T$, SeqFldNm$, Optional RestFny0, Optional IncFny0)
''Assume T is sorted
''
''Update A->T->SeqFldNm using RestFny0,IncFny0, assume the table has been sorted
''Update A->T->SeqFldNm using OrdFny0, RestFny0,IncFny0
'Dim RestFny$(), IncFny$(), Sql$
'Dim LasRestVy(), LasIncVy(), Seq&, OrdS$, Rs As Dao.Recordset
''OrdFny RestAy IncAy Sql
'RestFny = DftNy(RestFny0)
'IncFny = DftNy(IncFny0)
'If Si(RestFny) = 0 And Si(IncFny) = 0 Then
'    With A.OpenRecordset(T)
'        Seq = 1
'        While Not .EOF
'            .Edit
'            .Fields(SeqFldNm) = Seq
'            Seq = Seq + 1
'            .Update
'            .MoveNext
'        Wend
'        .Close
'    End With
'    Exit Sub
'End If
''--
'Set Rs = A.OpenRecordset(T) ', RecordOpenOptionsEnum.dbOpenForwardOnly, dbForwardOnly)
'With Rs
'    While Not .EOF
'        If RsIsBrk(Rs, RestFny, LasRestVy) Then
'            Seq = 1
'            LasRestVy = RsVy(Rs, RestFny)
'            LasIncVy = RsVy(Rs, IncFny)
'        Else
'            If RsIsBrk(Rs, IncFny, LasIncVy) Then
'                Seq = Seq + 1
'                LasIncVy = RsVy(Rs, IncFny)
'            End If
'        End If
'        .Edit
'        .Fields(SeqFldNm).Value = Seq
'        .Update
'        .MoveNext
'    Wend
'End With
'End Sub
'
'Private Sub LoCol_SetFml(A As ListObject, ColNm$, Fml$)
'A.ListColumns(ColNm).DataBodyRange.Formula = Fml
'End Sub
'
'Private Function PmnmFfn$(A$)
'PmnmFfn = PmnmVal("OupPth") & PmnmFn(A)
'End Function
'Private Function PmnmFn$(A$)
'PmnmFn = PmnmVal(A & "Fn")
'End Function
'
'Private Function RsCsv$(A As Dao.Recordset)
'RsCsv = FldsCsv(A.Fields)
'End Function
'
'Private Function WsC(A As Worksheet, C) As Range
'Dim R As Range
'Set R = A.Columns(C)
'Set WsC = R.EntireColumn
'End Function
'
'Private Function SyQuoteSqBktCsv$(A)
'SyQuoteSqBktCsv = JnComma(SyQuoteSqBkt(A))
'End Function
'
'Private Function LinRmvTerm$(ByVal A$)
'LinShiftTerm A
'LinRmvTerm = A
'End Function
'
'Private Sub ZZ_DbtReSeqFld()
'DbtReSeqFld CurrentDb, "ZZ_DbtUpdSeq", "Permit PermitD"
'End Sub
'
'Private Sub DbtReSeqFld(A as Database, T, ReSeqSpec$)
'DbtReSeqFldByFny A, T, ReSeqSpec_Fny(ReSeqSpec)
'End Sub
'
'Private Function SslyDic(A$()) As Dictionary
'Dim I, L$, K$, O As New Dictionary
'If Si(A) > 0 Then
'    For Each I In A
'        L = I
'        K = LinShiftTerm(L)
'        O.Add K, L
'    Next
'End If
'Set SslyDic = O
'End Function
'
'Private Sub ZZ_ReSeqSpec_Fny()
'AyBrw ReSeqSpec_Fny("*Flg RecTy Amt *Key *Uom MovTy Qty BchRateUX RateTy *Bch *Las *GL |" & _
'" Flg IsAlert IsWithSku |" & _
'" Key Sku PstMth PstDte |" & _
'" Bch BchNo BchPermitDate BchPermit |" & _
'" Las LasBchNo LasPermitDate LasPermit |" & _
'" GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
'" Uom Des StkUom Ac_U")
'End Sub
'
'Private Function ReSeqSpec_Fny(A$) As String()
'Dim Ay$(), D As Dictionary, O$(), L1$, L
'Ay = SplitVBar(A)
'L1 = AyShift(Ay)
'Set D = SslyDic(Ay)
'For Each L In SslSy(L1)
'    If D.Exists(L) Then
'        Push O, D(L)
'    Else
'        Push O, L
'    End If
'Next
'ReSeqSpec_Fny = SslSy(JnSpc(O))
'End Function
'Private Sub DbReOpn(A As Database)
'Dim Nm$
'Nm = A.Name
'A.Close
'Set A = Dao.DBEngine.OpenDatabase(Nm)
'End Sub
'Private Sub DbtReSeqFldByFny(A as Database, T, Fny$())
'Dim TFny$(), F$(), J%, FF
'TFny = DbtFny(A, T)
'If Si(TFny) = Si(Fny) Then
'    F = Fny
'Else
'    F = AyAdd(Fny, AyMinus(TFny, Fny))
'End If
'For Each FF In F
'    J = J + 1
'    A.TableDefs(T).Fields(FF).OrdinalPosition = J
'Next
'End Sub
'Private Function OyDRs(A, PrpNy0) As Drs
'Dim Fny$(), Dry()
'Fny = DftNy(PrpNy0)
'Dry = OyDry(A, Fny)
'Set OyDrs = Drs(Fny, Dry)
'End Function
'Private Function ObjDr(A, PrpNy0) As Variant()
'Dim PrpNy$(), U%, O(), J%
'PrpNy = DftNy(PrpNy0)
'U = UB(PrpNy)
'ReDim O(U)
'For J = 0 To U
'    Asg Prp(A, PrpNy(J)), O(J)
'Next
'ObjDr = O
'End Function
'Private Function OyDry(A, PrpNy0) As Variant()
'Dim O(), U%, I
'Dim PrpNy$()
'PrpNy = DftNy(PrpNy0)
'For Each I In A
'    Push O, ObjDr(I, PrpNy)
'Next
'OyDry = O
'End Function
'Private Sub ZZ_OyDrs()
'WsVis DrsWs(OyDrs(CurrentDb.TableDefs("ZZ_DbtUpdSeq").Fields, "Name Type OrdinalPosition"))
'End Sub
'Private Function DrsWs(A As Drs) As Worksheet
'DrsWs = SqWs(DrsSq(A))
'End Function
'Private Function DryWs(A) As Worksheet
'Set DryWs = SqWs(DrySq(A))
'End Function
'Private Function DryNCol%(A)
'Dim O%, Dr
'For Each Dr In A
'    O = Max(O, Si(Dr))
'Next
'DryNCol = O
'End Function
'Private Function DrySq(A) As Variant()
'Dim O(), C%, R&, Dr()
'Dim NC%, NR&
'NC = DryNCol(A)
'NR = UB(A)
'ReDim O(1 To NR, 1 To NC)
'For R = 1 To NR
'    Dr = A(R - 1)
'    For C = 1 To Min(Si(Dr), NC)
'        O(R, C) = Dr(C - 1)
'    Next
'Next
'DrySq = O
'End Function
'Private Function DbPth$(A As Database)
'DbPth = FfnPth(A.Name)
'End Function
'Private Function CurDbPth$()
'CurDbPth = DbPth(CurrentDb)
'End Function
'Private Function DrsNCol%(A As Drs)
'DrsNCol = Max(Si(A.Fny), DryNCol(A.Dry))
'End Function
'Private Function DrsSq(A As Drs) As Variant()
'Dim O(), C%, R&, Dr(), Dry()
'Dim Fny$(), NC%, NR&
'Dry = A.Dry
'Fny = A.Fny
'
'NR = Si(Dry)
'NC = DrsNCol(A)
'If Si(Fny) <> NC Then Stop
'ReDim O(1 To NR + 1, 1 To NC)
'For C = 1 To NC
'    O(1, C) = Fny(C - 1)
'Next
'For R = 1 To NR
'    Dr = Dry(R - 1)
'    For C = 1 To Min(Si(Dr), NC)
'        O(R + 1, C) = Dr(C - 1)
'    Next
'Next
'DrsSq = O
'End Function
'Private Function SqWs(A) As Worksheet
'Set SqWs = WszLo(SqLo(A))
'End Function
'
'Private Sub WQuit()
'WCls
'Set W = Nothing
'Quit
'End Sub
'Private Sub WReOpn()
'WCls
'WOpn
'End Sub
'Private Sub WClsTbl()
'AcsClsTbl WAcs
'End Sub
'
'
'Private Sub WBrw()
'AcsVis WAcs
'End Sub
'Private Sub WRenTbl(Fm$, ToTbl$)
'DbtRen W, Fm, ToTbl
'End Sub
'
'Private Function WTny() As String()
'WTny = DbTny(W)
'End Function
'
'Private Function WStru(Optional Tny0) As String()
'WStru = DbStru(W, Tny0)
'End Function
'
'Private Sub WClr()
'Dim T, Tny$()
'Tny = WTny: If Si(Tny) = 0 Then Exit Sub
'For Each T In Tny
'    WDrp T
'Next
'End Sub
'
'Private Function WAcs() As Access.Application
'Set WAcs = ApnAcs(Apn)
'End Function
'
'Private Function WFb$()
'WFb = ApnWFb(Apn)
'End Function
'
'Private Sub WKill()
'WCls
'Kill WFb
'End Sub
'
'Private Function WDStru1() As String()
'WDStru1 = WStru(DNm1)
'End Function
'Private Function WDb() As Database
'Set WDb = ApnWDb(Apn)
'End Function
'
'
'Private Function WLnkFb(T$, Fb$) As String()
'WLnkFb = DbtLnkFb(W, T, Fb)
'End Function
'
'Private Function WLnkFx(T$, Fx$) As String()
'WLnkFx = DbtLnkFx(W, T, Fx)
'End Function
'
'Private Function WWbCnStr$()
'WWbCnStr = FbWbCnStr(WFb)
'End Function
'
'Private Sub ZZ_WLnkFx()
'WOpn
'D WLnkFx(">MB51", IFx_MB51)
'End Sub
'
'Private Function WFny(T) As String()
'WFny = Fny(W, T)
'End Function
'
