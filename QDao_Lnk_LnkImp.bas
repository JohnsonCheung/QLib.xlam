Attribute VB_Name = "QDao_Lnk_LnkImp"
Option Compare Text
Option Explicit
Private Const CMod$ = "BLnkImp."
Sub ZZ_LnkImp()
Dim LnkImpSrc$(), Db As Database
GoSub T0
Exit Sub
T0:
    LnkImpSrc = Y_LnkImpSrc
    Set Db = TmpDb
    GoTo Tst
Tst:
    LnkImp LnkImpSrc, Db
    Return
End Sub

Sub LnkImp(LnkImpSrc$(), Db As Database)
'ThwIf_Er ErzLnk(InpFilSrc, LnkImpSrc), CSub
Dim Ip          As DLTDH:                   Ip = DLTDH(LnkImpSrc)
Dim FbTblLy$():                        FbTblLy = IndentedLy(LnkImpSrc, "FbTbl")
Dim Dic_Fbt_Fbn As Dictionary: Set Dic_Fbt_Fbn = DiczVkkLy(FbTblLy)
Dim FbTny$():                            FbTny = SyzDicKey(Dic_Fbt_Fbn)

Dim FxTblLy$(): FxTblLy = IndentedLy(LnkImpSrc, "FxTbl")
Dim DFx As Drs:     DFx = WDFx(FxTblLy)                  ' T Fxn Wsn Stru
Dim FxTny$():     FxTny = StrColzDrs(DFx, "T")

Dim Dic_Fn_Ffn As Dictionary: Set Dic_Fn_Ffn = Dic(IndentedLy(LnkImpSrc, "Inp"))

'== Lnk=================================================================================================================
Dim D1   As Drs:   D1 = WLnkFx(DFx, Dic_Fn_Ffn)         ' T S Cn
Dim D2   As Drs:   D2 = WLnkFb(Dic_Fbt_Fbn, Dic_Fn_Ffn)
Dim D    As Drs:    D = AddDrs(D1, D2)
Dim OLnk As Unt: OLnk = LnkTblzDrs(Db, D)               ' <======
            
'== Imp=================================================================================================================
Dim Wh         As Dictionary:         Set Wh = Dic(IndentedLy(LnkImpSrc, "Tbl.Where"))
Dim Dic_T_Stru As Dictionary: Set Dic_T_Stru = WDic_T_Stru(FbTny, DFx)

Dim DStru    As Drs:    DStru = WDStru(Ip)                     ' Stru F Ty E
Dim ImpSqy$():         ImpSqy = WImpSqy(Dic_T_Stru, DStru, Wh)
Dim OImp     As Unt:     OImp = RunSqy(Db, ImpSqy)             ' <==========
Dim ODmpNRec As Unt: ODmpNRec = DmpNRec(Db)
End Sub

Private Function WDStru(Ip As DLTDH) As Drs
'Fm Ip : L T1 Dta IsHdr}
'Ret WDStru: Stru F Ty E
Dim A As Drs, Dr, Dry(), B As Drs, T1$, Dta$
A = DrswColEqSel(Ip.D, "IsHdr", False, "T1 Dta")
B = DrswColPfx(A, "T1", "Stru.") 'T1 Dta
For Each Dr In Itr(B.Dry)
    T1 = Dr(0)
    Dta = Dr(1)
    PushI Dry, XDrOfStru(T1, Dta)
Next
WDStru = DrszFF("Stru F Ty E", Dry)
End Function

Private Function XDrOfStru(T1$, Dta$) As Variant()
Dim F$, Ty$, E$, Stru$
Stru = RmvPfx(T1, "Stru.")
F = ShfT1(Dta)
Ty = ShfT1(Dta)
E = RmvSqBkt(Dta)
XDrOfStru = Array(Stru, F, Ty, E)
End Function

Private Function WDic_T_Stru(FbTny$(), DFx As Drs) As Dictionary
Dim Dr, IxT%, IxStru%, T
Set WDic_T_Stru = New Dictionary
For Each T In Itr(FbTny)
    WDic_T_Stru.Add T, T
Next
AsgIx DFx, "T Stru", IxT, IxStru
For Each Dr In Itr(DFx.Dry)
    WDic_T_Stru.Add Dr(IxT), Dr(IxStru)
Next
End Function

Private Function WImpSqy(Dic_T_Stru As Dictionary, DStru As Drs, Dic_T_Bexpr As Dictionary) As String()
Dim I, Fny$(), Ix As Dictionary, Ey$(), T$, Into$, LnkColLy$(), Bexpr$, A As Drs, Stru$
For Each I In Dic_T_Stru.Keys
    Stru = Dic_T_Stru(I)
       T = ">" & I
    Into = "#I" & I
       A = DrswColEqSel(DStru, "Stru", Stru, "F Ty E")
     Fny = StrColzDrs(A, "F")
      Ey = RmvSqBktzSy(StrColzDrs(A, "E"))
   Bexpr = ValzDicIf(Dic_T_Bexpr, I)
    PushI WImpSqy, SqlSel_Fny_Extny_Into_T_OB(Fny, Ey, Into, T, Bexpr)
Next
End Function

Private Function WDFx(FxTblLy$()) As Drs
'Ret DFx : T Fxn Ws Stru
Dim Lin, L$, A$, T$, Fxn$, Ws$, Stru$, Dry()
For Each Lin In Itr(FxTblLy)
    L = Lin
    T = ShfT1(L)
    A = ShfT1(L)
    Fxn = BefDotOrAll(A)
    Ws = AftDot(A)
    If Fxn = "" Then Fxn = T
    If Ws = "" Then Ws = "Sheet1"
    Stru = StrDft(L, T)
    PushI Dry, Array(T, Fxn, Ws, Stru)
Next
WDFx = DrszFF("T Fxn Ws Stru", Dry)
End Function

Private Function WLnkFb(Dic_Fbt_Fbn As Dictionary, Dic_Fbn_Fb As Dictionary) As Drs
'Ret: *LnkFb::Drs{T S Cn)
Dim Fbn$, A$, S$, Fbt, T$, Cn$, Fb$, Dry()
For Each Fbt In Dic_Fbt_Fbn.Keys
    Fbn = Dic_Fbt_Fbn(Fbt)
    If Not Dic_Fbn_Fb.Exists(Fbn) Then
        Thw CSub, "Dic_Fbn_Fb does not contains Fbn", "Fbn Dic_Fbn_Fb", Fbn, Dic_Fbn_Fb
    End If
    Fb = Dic_Fbn_Fb(Fbn)
    Cn = CnStrzFbDao(Fb)
    T = ">" & Fbt
    S = Fbt
    PushI Dry, Array(T, S, Cn)
Next
WLnkFb = DrszFF("T S Cn", Dry)
End Function

Private Function WLnkFx(DFx As Drs, Dic_Fxn_Fx As Dictionary) As Drs
'Fm : @DFx :: Drs{T Fxn Ws Stru}
'Ret: *LnkFx::Drs{T S Cn}
Dim Dry(), Dr, S$, Fx$, Ws$, Cn$, T$, Fxn$, IxT%, IxWs%, IxFxn%
AsgIx DFx, "T Ws Fxn", IxT, IxWs, IxFxn
For Each Dr In Itr(DFx.Dry)
    T = Dr(IxT)
    Ws = Dr(IxWs)
    Fxn = Dr(IxFxn)
    If Not Dic_Fxn_Fx.Exists(Fxn) Then Thw CSub, "Dic_Fxn_Fx does not have Key", "Fxn-Key Dic_Fxn_Fx", T, Dic_Fxn_Fx
    Fx = Dic_Fxn_Fx(Fxn)
    If IsNeedQuote(Ws) Then
        S = "'" & Ws & "$'"
    Else
        S = Ws & "$"
    End If
    Cn = CnStrzFxDao(Fx)
    T = ">" & T
    PushI Dry, Array(T, S, Cn)
Next
WLnkFx = DrszFF("T S Cn", Dry)
End Function

Private Property Get Y_LnkImpSrc() As String()
Erase XX
X "Inp"
X " DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X " ZHT0  C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X " MB52  C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X " Uom   C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X " GLBal C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
X "FbTbl"
X "--  Fbn TblNm.."
X " DutyPay Permit PermitD"
X "FxTbl T  FxNm.Wsn  Stru"
X " ZHT086  ZHT0.8600 ZHT0"
X " ZHT087  ZHT0.8700 ZHT0"
X " MB52                  "
X " Uom                   "
X " GLBal"
X "Tbl.Where"
X " MB52 Plant='8601' and [Storage Location] in ('0002','')"
X " Uom  Plant='8601'"
X "Stru.Permit"
X " Permit"
X " PermitNo"
X " PermitDate"
X " PostDate"
X " NSku"
X " Qty"
X " Tot"
X " GLAc"
X " GLAcName"
X " BankCode"
X " ByUsr"
X " DteCrt"
X " DteUpd"
X "Stru.PermitD"
X " PermitD"
X " Permit"
X " Sku"
X " SeqNo"
X " Qty"
X " BchNo"
X " Rate"
X " Amt"
X " DteCrt"
X " DteUpd"
X "Stru.ZHT0"
X " Sku       Txt Material    "
X " CurRateAc Dbl [     Amount]"
X " VdtFm     Txt Valid From  "
X " VdtTo     Txt Valid to    "
X " HKD       Txt Unit        "
X " Per       Txt per         "
X " CA_Uom    Txt Uom         "
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
X " Sc_U    Txt SC "
X " Topaz   Txt Topaz Code "
X " ProdH   Txt Product hierarchy"
X " Sku     Txt Material            "
X " Des     Txt Material Description"
X " AC_U    Txt Unit per case       "
X " SkuUom  Txt Base Unit of Measure"
X " BusArea Txt Business Area       "
X "Stru.GLBal"
X " BusArea Txt Business Area Code"
X " GLBal   Dbl                   "
X "Stru.SkuRepackMulti"
X " SkuRepackMulti   GLBal   Dbl                     "
X "Stru.SkuTaxBy3rdParty"
X " SkuTaxBy3rdParty GLBal   Dbl                     "
X "Stru.SkuNoLongerTax"
X " SkuNoLongerTax"
Y_LnkImpSrc = XX
Erase XX
End Property

Private Sub ZZZ()
QDao_Lnk_LnkImp:
End Sub




