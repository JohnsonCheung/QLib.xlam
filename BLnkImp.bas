Attribute VB_Name = "BLnkImp"
Option Explicit
Private Type A
    WDb As Database
    InpFilSrc() As String
    LnkImpSrc() As String
End Type
Private Type BLnk
    Tny() As String
    SrcNy() As String
    CnSy() As String
End Type
Private Type B
    ImpSqy() As String
    Lnk As BLnk
End Type
Private A As A
Private B As B


Sub LnkImp(InpFilSrc$(), LnkImpSrc$(), WDb As Database)
A.InpFilSrc = InpFilSrc
A.LnkImpSrc = LnkImpSrc
Set A.WDb = WDb
R0_B
R1_Lnk
R2_Imp
End Sub

Private Sub R1_Lnk()
With B.Lnk
    Dim J%, T$, S$, Cn$
    For J = 0 To UB(B.Lnk.Tny)
        T = .Tny(J)
        S = .SrcNy(J)
        Cn = .CnSy(J)
        LnkTbl A.WDb, T, S, Cn
    Next
End With
End Sub
Private Sub R2_Imp()
RunSqy A.WDb, B.ImpSqy
End Sub
Private Sub R0_B()
R01_ThwIfErzInpFilSrc
R02_ThwIfErzLnkImpSrc
R03_ThwIfErzLnkImpSrcPlusInp
R04_SetB_BLnk
R05_SetB_ImpSqy
End Sub

Private Sub R05_SetB_ImpSqy()
R051_SetB_ImpSqy_Fx
R052_SetB_ImpSqy_Fb
End Sub

Private Sub R04_SetB_BLnk()
Erase B.Lnk.CnSy
Erase B.Lnk.Tny
Erase B.Lnk.SrcNy
R041_SetB_BLnk_Fx
R042_SetB_BLnk_Fb
End Sub

Private Sub R0421_Asg_FbSy_Tny(OFbSy$(), OTny$())
Stop
End Sub
Private Sub R041_SetB_BLnk_Fx()
Dim FxSy$(), Tny$(), Wsny$()
    R0411_Asg_FxSy_Tny_Wsny FxSy, Tny, Wsny
Dim J%, Fx$, Wsn$, T$, M As LnkTblPm

Dim SrcNy$(), CnSy$()
    For J = 0 To UB(Tny)
        PushI CnSy, CnStrzFxAdo(FxSy(J))
        PushI SrcNy, Wsny(J) & "$"
    Next
With B.Lnk
    PushIAy .Tny, Tny
    PushIAy .CnSy, CnSy
    PushIAy .SrcNy, SrcNy
End With

End Sub
Function R0411_Asg_FxSy_Tny_Wsny(OFxSy$(), OTny$(), OWsny$())

End Function
Private Sub R042_SetB_BLnk_Fb()
Dim FbSy$(), Tny$()
    R0421_Asg_FbSy_Tny FbSy, Tny

Dim SrcNy$(), CnSy$()
    Dim J%
    For J = 0 To UB(Tny)
        PushI CnSy, CnStrzFbAdo(FbSy(J))
        PushI SrcNy, ">" & Tny(J)
    Next
With B.Lnk
    PushIAy .Tny, Tny
    PushIAy .CnSy, CnSy
    PushIAy .SrcNy, SrcNy
End With
End Sub
Private Sub R02_ThwIfErzLnkImpSrc()

End Sub
Private Sub R01_ThwIfErzInpFilSrc()

End Sub
Private Sub R03_ThwIfErzLnkImpSrcPlusInp()

End Sub


Private Property Get TaxAlertLnkPmLy() As String()
Erase XX
X "@Tbl  GLBal Uom CurRate MB52 Permit PermitD SkuRepackMulti SkuTaxBy3rdParty SkuNoLongerTax"
X "@Stru GLBal Uom CurRate MB52 Permit PermitD SkuRepackMulti SkuTaxBy3rdParty SkuNoLongerTax"
X "TblWhere.MB52 Plant='8601' and [Storage Location] in ('0002','')"
X "TblWhere.Uom  Plant='8601'"
X "Stru.ZHT0"
X " Sku       Txt Material   "
X " CurRateAc Dbl Amount     "
X " VdtFm     Txt Valid From "
X " VdtTo     Txt Valid to   "
X " HKD       Txt Unit       "
X " Per       Txt per        "
X " CA_Uom    Txt Uom        "
X "Stru.MB52"
X " Sku    Txt Material          "
X " Whs    Txt Plant             "
X " Loc    Txt Storage Location  "
X " BchNo  Txt Batch             "
X " QInsp  Dbl In Quality Insp#  "
X " QUnRes Dbl UnRestricted      "
X " QBlk   Dbl Blocked           "
X " VInsp  Dbl Value in QualInsp#"
X " VUnRes Dbl Value Unrestricted"
X " VBlk   Dbl Value BlockedStock"
X "Stru.Uom"
X " Sku     Txt Material            "
X " Des     Txt Material Description"
X " AC_U    Txt Unit per case]      "
X " SkuUom  Txt Base Unit of Measure"
X " BusArea Txt Business Area       "
X "Stru.GLBar"
X " BusArea Txt Business Area Code"
X " GLBal   Dbl                   "
X "Stru.PermitD"
X " Permit           GLBal   Dbl                     "
X " PermitD          GLBal   Dbl                     "
X "Stru.SkuRepackMulti"
X " SkuRepackMulti   GLBal   Dbl                     "
X "Stru.SkuTaxBy3rdParty"
X " SkuTaxBy3rdParty GLBal   Dbl                     "
X "Stru.SkuNoLongerTax"
X " SkuNoLongerTax"
End Property
Private Property Get SampLnkImpSrc() As String()
Erase XX
X "#     T    Fx   Ws   S"
X "FxTbl Z86  ZHT1.8601 ZHT1"
X "FxTbl Z87  ZHT1.8701 ZHT1"
X "FxTbl Uom  ZHT1      Uom"
X "FxTbl Uom  ZHT1      Uom"
X "Stru.ZHT1"
X " ZHT1   D Brand  "
X " RateSc M Amount "
X " VdtFm  M Valid From  "
X " VdtTo  M Valid to"
X "Stru.Uom"
X " Sku     M Material"
X " Des     M Material Description"
X " Sc_U    M SC "
X " StkUom  M Base Unit of Measure"
X " Topaz   M Topaz Code "
X " ProdH   M Product hierarchy"
X "Stru.MB52"
X " Sku    M Material "
X " Whs    M Plant    "
X " QInsp  D In Quality Insp#"
X " QUnRes D Unrestricted"
X " QBlk   D Blocked"
SampLnkImpSrc = XX
Erase XX
End Property
Private Function SampInpFilSrc() As String()

End Function

Private Sub Z_LnkImp()
Dim Db As Database: Set Db = TmpDb
LnkImp SampInpFilSrc, SampLnkImpSrc, Db
BrwDb Db
Stop
End Sub


Private Function R0521_SetB_ImpSqy_Fb_OneTbl(T$, Fny$(), Bexpr$) As String()
Dim Fm$: Fm = ">" & T
Dim Into$: Into = "#I" & T
PushI B.ImpSqy, SqlSel_Fny_Into_T(Fny, Into, Fm, Bexpr)
End Function

Function SampSrczLnkImp() As String()
Erase XX
X "Nm ShpCst"
X "WszT ZHT1 8701   ZHT18701"
X "WszT ZHT1 8601   ZHT18601"
X "WszT MB52 Sheet1 MB52"
X "WszT UOM  Sheet1 UOM"
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
SampSrczLnkImp = XX
Erase XX
End Function

Private Sub R052_SetB_ImpSqy_Fb()
Dim J%, Tny$(), T$, Fny$(), Bexpr$
For J = 0 To UB(Tny)
    T = Tny(J)
    R0521_SetB_ImpSqy_Fb_OneTbl T, Fny, Bexpr
Next
End Sub

Private Sub R051_SetB_ImpSqy_Fx()

End Sub
