Attribute VB_Name = "QDao_Lnk_LnkImp"
Option Explicit
Private Const CMod$ = "BLnkImp."
Private Type FxtRec
    T As String
    Fxn As String
    Wsn As String
    Stru As String
End Type
Private Type FxtRecs: N As Byte: Ay() As FxtRec: End Type
Private Sub ZZ_LnkImp()
Dim InpFilSrc$(), LnkImpSrc$(), Db As Database
GoSub T0
Exit Sub
T0:
    InpFilSrc = Y_InpFilSrc
    LnkImpSrc = Y_LnkImpSrc
    Set Db = TmpDb
    GoTo Tst
Tst:
    LnkImp InpFilSrc, LnkImpSrc, Db
    Return
End Sub

Sub LnkImp(InpFilSrc$(), LnkImpSrc$(), Db As Database)
'ThwIf_LnkImpPmEr InpFilSrc, LnkImpSrc
Dim a___FbTbl__fm_TblLy
    Dim FbTblLy$():                      FbTblLy = IndentedLy(LnkImpSrc, "FbTbl")
    Dim b___FbTbl$
    Dim FbTny$():                          FbTny = B_TnyFb(FbTblLy)
    Dim Dic_Fbt_Fn As Dictionary: Set Dic_Fbt_Fn = DiczVkkLy(FbTblLy)

Dim a___FxtRecs__fm_FxTblLy$
    Dim FxTblLy$():         FxTblLy = IndentedLy(LnkImpSrc, "FxTbl")
    Dim b___FxtRecs$
    Dim FxtRecs As FxtRecs: FxtRecs = B_Fnd_FxtRecs(FxTblLy)   'T* :: TnyOf*
    
Dim b___Tny
    Dim FxTny$():             FxTny = B_TnyFx(FxtRecs)
    Dim Tny$():                 Tny = AddSy(FbTny, FxTny)

Dim a___Stru$
    Dim StruSy() As String:  StruSy = B_Fnd_StruSy(LnkImpSrc)

Dim a___Dic_T_Stru__$
    Dim Dic_Fn_Ffn As Dictionary:         Set Dic_Fn_Ffn = Dic(InpFilSrc)
    Dim Dic_Fxt_Fn As Dictionary:         Set Dic_Fxt_Fn = B_Dic_Fxt_Fn(FxtRecs)
    Dim Dic_T_Fn As Dictionary:             Set Dic_T_Fn = AddDic(Dic_Fxt_Fn, Dic_Fbt_Fn)
    Dim b___Dic_T_Stru$
    Dim T_zDi_Ffn As Dictionary:           Set T_zDi_Ffn = AzDiC(Dic_T_Fn, Dic_Fn_Ffn)
    Dim Dic_T_Stru As Dictionary:         Set Dic_T_Stru = B_Dic_T_Stru(FbTny, FxtRecs)

Dim a___ImpSqy$
    Dim Dic_T_LnkColLy As Dictionary: Set Dic_T_LnkColLy = B_Dic_T_LnkColLy(LnkImpSrc, Dic_T_Stru, StruSy)
    Dim LTblWh$():                                LTblWh = IndentedLy(LnkImpSrc, "Tbl.Where")
    Dim Dic_T_Wh As Dictionary:             Set Dic_T_Wh = Dic(LTblWh)
    Dim b____ImpSqy$
    Dim ImpSqy$():                                ImpSqy = B__ImpSqy(Tny, Dic_T_LnkColLy, Dic_T_Wh)

Dim a___Lnk__AllInpTbl_LnkTblPms$
    Dim LnkFb As LnkTblPms, LnkFx As LnkTblPms
    LnkFx = B__LnkTblPms_Fx(FxtRecs, T_zDi_Ffn)
    LnkFb = B__LnkTblPms_Fb(FbTny, T_zDi_Ffn)
    Dim b___Lnk$
    Dim Lnk As LnkTblPms: Lnk = AddLnkTblPms(LnkFb, LnkFx)
    
Dim a___doing_Lnk_and_Imp$
LnkImpzLnkpmSqyDb Lnk, ImpSqy, Db
End Sub
Private Function B_Fnd_StruSy(LnkImpSrc$()) As String()
Static X As Boolean, Y
If Not X Then
    Y = Sy()
    X = True
    Dim I, L$
    For Each I In LnkImpSrc
        L = I
        If HasPfx(L, "Stru.") Then
            PushI Y, BefSpcOrAll(RmvPfx(L, "Stru."))
        End If
    Next
End If
B_Fnd_StruSy = Y
End Function


Sub LnkImpzLnkpmSqyDb(L As LnkTblPms, ImpSqy$(), Db As Database)
LnkTblzPms Db, L '<==============
RunSqy Db, ImpSqy    '<==========
DmpNRec Db
End Sub
Private Function B_Dic_T_Stru(Fbt$(), A As FxtRecs) As Dictionary
Set B_Dic_T_Stru = New Dictionary
Dim T, J%
For Each T In Itr(Fbt)
    B_Dic_T_Stru.Add T, T
Next
For J = 0 To A.N - 1
    With A.Ay(J)
        B_Dic_T_Stru.Add .T, .Stru
    End With
Next
End Function

Private Function B_Dic_Fxt_Ffn(A As FxtRecs, T_Ffn As Dictionary) As Dictionary
Dim J%
Set B_Dic_Fxt_Ffn = New Dictionary
For J = 0 To A.N - 1
    With A.Ay(J)
        B_Dic_Fxt_Ffn.Add .T, T_Ffn(.Fxn)
    End With
Next
End Function
Private Function B_Dic_Fxt_Fn(A As FxtRecs) As Dictionary
Set B_Dic_Fxt_Fn = New Dictionary
Dim J%
For J = 0 To A.N - 1
    With A.Ay(J)
    B_Dic_Fxt_Fn.Add .T, .Fxn
    End With
Next
End Function
Private Function B_TnyFx(A As FxtRecs) As String()
Dim J%
For J = 0 To A.N - 1
    PushI B_TnyFx, A.Ay(J).T
Next
End Function
Private Function B_TnyFb(LFbTbl$()) As String()
Dim J%
For J = 0 To UB(LFbTbl)
    PushIAy B_TnyFb, SyzSS(RmvT1(LFbTbl(J)))
Next
End Function

Private Function B__ImpSqy(Tny$(), TzDiLnkColLy As Dictionary, TzDiBexpr As Dictionary) As String()
Dim I, Fny$(), Ey$(), T$, Into$, LnkColLy$(), Bexpr$
For Each I In Itr(Tny)
       T = ">" & I
    Into = "#I" & I
LnkColLy = ValzDicK(TzDiLnkColLy, I, Dicn:="TzDiLnkColLy", Kn:="TblNm", Fun:=CSub)
     Fny = T1Ay(LnkColLy)
      Ey = RmvSqBktzSy(RmvTTzAy(LnkColLy))
   Bexpr = ValzDicIf(TzDiBexpr, I)
    PushI B__ImpSqy, SqlSel_Fny_Extny_Into_T_OB(Fny, Ey, Into, T, Bexpr)
Next
End Function

Private Function B_Fnd_FxtRecs(FxTblLy$()) As FxtRecs
Dim OAy() As FxtRec, J%, L$, A$
For J = 0 To UB(FxTblLy)
    L = FxTblLy(J)
    ReDim Preserve OAy(J)
    With OAy(J)
        .T = ShfT1(L)
        A = ShfT1(L)
        .Fxn = B_Fnd_Fxn(A, .T)
        .Wsn = B_Fnd_Wsn(A)
        .Stru = StrDft(L, .T)
    End With
Next
B_Fnd_FxtRecs.N = Si(FxTblLy)
B_Fnd_FxtRecs.Ay = OAy
End Function
Private Function B_Fnd_Wsn(Fxn_dot_Wsn)
Dim A$: A = Fxn_dot_Wsn
If A = "" Then B_Fnd_Wsn = "Sheet1": Exit Function
If Not HasDot(A) Then B_Fnd_Wsn = "Sheet1": Exit Function
B_Fnd_Wsn = AftDot(A)
End Function
Private Function B_Fnd_Fxn(Fxn_dot_Wsn, T)
Dim A$: A = Fxn_dot_Wsn
If A = "" Then B_Fnd_Fxn = T: Exit Function
If HasDot(A) Then B_Fnd_Fxn = BefDot(A): Exit Function
B_Fnd_Fxn = Fxn_dot_Wsn
End Function
Private Function B__LnkTblPms_Fb(TFb$(), TzDiFbFx As Dictionary) As LnkTblPms
Dim Fbn$, A$, T, Cn$
For Each T In Itr(TFb)
    If Not TzDiFbFx.Exists(T) Then
        Thw CSub, "TzDiFbFx does not contains T", "T TzDiFbFx TFb", Fbn, TzDiFbFx, T, TFb
    End If
    Cn = CnStrzFbDao(TzDiFbFx(T))
    PushLnkTblPm B__LnkTblPms_Fb, LnkTblPm(">" & T, T, Cn)
Next
End Function

Private Function B__LnkTblPms_Fx(A As FxtRecs, TzDiFbFx As Dictionary) As LnkTblPms
Dim J%, S$, Fx$, Cn$
For J = 0 To A.N - 1
    With A.Ay(J)
    Fx = TzDiFbFx(.T)
    If Fx = "" Then Thw CSub, "TzDiFbFx does not have Key", "Tbl-Key TblNmToTzDiFbFx", .T, TzDiFbFx
    If IsNeedQuote(.Wsn) Then
        S = "'" & .Wsn & "$'"
    Else
        S = .Wsn & "$"
    End If
    Cn = CnStrzFxDao(Fx)
    PushLnkTblPm B__LnkTblPms_Fx, LnkTblPm(">" & .T, S, Cn)
    End With
Next
End Function


Private Function B_Dic_T_LnkColLy(LnkImpSrc$(), TzDiStru As Dictionary, StruSy$()) As Dictionary
Dim T, Stru$, LnkColLy$()
Set B_Dic_T_LnkColLy = New Dictionary
For Each T In TzDiStru.Keys
    Stru = TzDiStru(T)
    LnkColLy = IndentedLy(LnkImpSrc, "Stru." & Stru)
    B_Dic_T_LnkColLy.Add T, LnkColLy
Next
End Function



Private Property Get Y_InpFilSrc() As String()
Erase XX
X "DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X "ZHT0  C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X "MB52  C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X "Uom   C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X "GLBal C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
Y_InpFilSrc = XX
Erase XX
SampSrczInpFil
End Property


Private Property Get Y_LnkImpSrc() As String()
Erase XX
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
X "Stru.PermitD"
X " Permit           GLBal   Dbl                     "
X " PermitD          GLBal   Dbl                     "
X "Stru.SkuRepackMulti"
X " SkuRepackMulti   GLBal   Dbl                     "
X "Stru.SkuTaxBy3rdParty"
X " SkuTaxBy3rdParty GLBal   Dbl                     "
X "Stru.SkuNoLongerTax"
X " SkuNoLongerTax"
Y_LnkImpSrc = XX
Erase XX
End Property
Private Sub ZZ()
ZZ_LnkImp
End Sub
Private Sub ZZZ()
QDao_Lnk_LnkImp.ZZ_LnkImp
End Sub

Sub Z1()
ZZ_LnkImp
End Sub
