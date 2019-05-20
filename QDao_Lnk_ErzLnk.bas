Attribute VB_Name = "QDao_Lnk_ErzLnk"
Option Compare Text
Option Explicit
'==================================================
'I Inp
'T Tmp
'O Oup
'OU for user action
'OS for Stru
'OX for FxTbl
'OB for FbTbl
'OW for TblWh
Const I1$ = "L Lin IsHdr IsRmk T1 Rst" 'ILnkImp
Const I2$ = "L Lin Fn Ffn" 'IInpFil
Const J1$ = "L Stru" 'JStru
Const J2$ = "L Stru F Ty Extn" 'JStruF
Const J3$ = "L T Fxn Wsn Stru" 'JFxTbl
Const J4$ = "L Fbn Tss" 'JFbTbl
Const J5$ = "L T Wh" 'JTblWh
Const T1$ = "T Stru Ffn" 'TTbl
Const T2$ = "T Fxn Wsn F Ty Stru L" 'TEFxTbl
Const T3$ = "T Fbn Wsn F Ty" 'TEFbTbl
Const T4$ = "Fxn Wsn F Ty" 'TAFxTblF
Const T5$ = "Fbn T F Ty" 'TAFbTblF
Const O1$ = "ErTy Ern L Fn Ffn Fbn Fxn Fb Fx Wsn Lss Stru T F Fss Ty ActTy ActFss"
Const OU1$ = "L Fn Ffn" 'OU_FfnNFnd
Const OU2$ = "L Fbn Fb Tss" 'OU_FbtNFnd
Const OU3$ = "L Fxn Fx Wsn" 'OU_WsnNFnd
Const OU4$ = "Fn Ffn T Fss ActFss" 'OU_FxExtnNFnd
Const OU5$ = "Fn Ffn T F Ty ActTy" 'OU_FxFldTyNMch
Const OS1$ = "Lss S F" 'OS_FldDup
Const OS2$ = "Lss S" 'OS_StruDup
Const OS3$ = "L S F Ty" 'OS_FldTyEr
Const OS4$ = "L S" 'StruExcess
Const OS5$ = "Lss S" 'OS_StruInUse
Const OS6$ = "Fxn Fx T Fss" 'OS_FxFldTyMis
Const OS7$ = "L S" 'OS_FldMis
Const OX1$ = "L S" 'OX_StruNDef
Const OX2$ = "Lss T" 'OX_FxtDup
Const OX3$ = "Lss Fxn" 'OX_FxnNDef
Const OB1$ = "L T Fbn" 'OB_StruNDef
Const OB2$ = "L Fbn" 'OB_FbnNDef
Const OB3$ = "L Fbn Tss" 'OB_FbtMis
Const OW1$ = "Lss T" 'OW_TblDup
Const OW2$ = "L T" 'OW_TblNDef
Const OW3$ = "L T" 'OW_TblExcess

Const ILnkImp_$ = I1
Const IInpFil_$ = I2
Const JStru_$ = J1
Const JStruF_$ = J2
Const JFxTbl_$ = J3
Const JFbTbl_$ = J4
Const JTblWh_$ = J5
Const TTbl_$ = T1
Const TEFxTbl_$ = T2
Const TEFbTbl_$ = T3
Const TAFxTblF_$ = T4
Const TAFbTblF_$ = T5
Const OEr_$ = O1
Const OU_FfnNFnd_$ = OU1
Const OU_FbtNFnd_$ = OU2
Const OU_WsnNFnd_$ = OU3
Const OU_FxExtnNFnd_$ = OU4
Const OU_FxFldTyNMch_$ = OU5
Const OS_FldDup_$ = OS1
Const OS_StruDup_$ = OS2
Const OS_FldTyEr_$ = OS3
Const OS_StruExcess_$ = OS4
Const OS_StruInUse_$ = OS5
Const OS_FxFldTyMis_$ = OS6
Const OS_FldMis_$ = OS7
Const OX_StruNDef_$ = OX1
Const OX_FxtDup_$ = OX1
Const OX_FxnNDef_$ = OX2
Const OB_StruNDef_$ = OB1
Const OB_FbnNDef_$ = OB2
Const OB_FbtMis_$ = OB3
Const OW_TblDup_$ = OW1
Const OW_TblNDef_$ = OW2
Const OW_TblExcess_$ = OW3
Sub ResiDrs(O As Drs, NCol%)
Dim U%: U = NCol - 1
Dim Dr, J&
For J = 0 To UB(O.Dry)
    Dr = O.Dry(J)
    ReDim Preserve Dr(U)
    O.Dry(J) = Dr
Next
End Sub
Private Function B_OU_FfnNFnd(IInpFil As Drs) As Drs
IInpFil_: 'L Lin Fn Ffn
Dim A As Drs: A = SelDrs(IInpFil, "L Fn Ffn")
PushI A.Fny, "HasFfn"
ResiDrs A, 4
Dim Dr, Ffn$, Has As Boolean, J%
For Each Dr In A.Dry
    Ffn = Dr(2)
    Has = HasFfn(Ffn)
    Dr(3) = Has
    A.Dry(J) = Dr
    J = J + 1
Next
Dim B As Drs: B = DrswColEq(A, "HasFfn", False)
Dim C():      C = KeepFstNCol(B.Dry, 3)
B_OU_FfnNFnd = DrszFF(OU_FfnNFnd_, C) ' L Fn Ffn
End Function
Private Function B_OU_FbtNFnd() As Drs

End Function
Private Function B_OU_WsnNFnd() As Drs

End Function
Private Function B_OU_FxExtnNFnd() As Drs

End Function
Private Function B_OU_FxFldTyNMch() As Drs

End Function
Private Function B_OS_FldDup() As Drs

End Function
Private Function B_OS_StruDup() As Drs

End Function
Private Function B_OS_FldTyEr() As Drs

End Function
Private Function B_OS_StruExcess() As Drs

End Function
Private Function B_OS_StruInUse() As Drs

End Function
Private Function B_OS_FxFldTyMis() As Drs

End Function
Private Function B_OS_FldMis() As Drs

End Function
Private Function B_OX_StruNDef() As Drs

End Function
Private Function B_OX_FxtDup() As Drs

End Function
Private Function B_OX_FxnNDef() As Drs

End Function
Private Function B_OB_StruNDef() As Drs

End Function
Private Function B_OB_FbnNFnd() As Drs

End Function

Private Function B_OB_FbtMis() As Drs

End Function

Private Function B_OB_FbnMis() As Drs

End Function

Private Function B_OW_TblDup() As Drs

End Function

Private Function B_OW_TblNDef() As Drs

End Function

Private Function B_OW_TblExcess() As Drs
Dim Dry()
B_OW_TblExcess = DrszFF(OW_TblExcess_, Dry)
End Function
Private Function B_JStru(ILnkImp As Drs) As Drs
Dim A(): A = SelDrs(ILnkImp, "L T1 IsHdr").Dry
Dim B(): B = DrywColPfx(A, 1, "Stru.")
Dim C(): C = DrywColPfx(B, 2, True)
Dim D(): D = KeepFstNCol(C, 2)
Dim J&
For J = 0 To UB(D)
    D(J)(1) = Mid(D(J)(1), 6)
Next
B_JStru = DrszFF(JStru_, D) 'L Stru
End Function

Private Function B_JStruF(ILnkImp As Drs) As Drs
'L Lin IsHdr IsRmk T1 Rst  ILnkImp
'L Stru F Ty Extn          JStruF
Dim A As Drs: A = DrswColEq(ILnkImp, "IsHdr", False)
Dim B As Drs: B = DrswColEq(A, "IsRmk", False)
Dim C As Drs: C = DrswColPfx(B, "T1", "Stru.")
Dim D As Drs: D = SelDrsAlwEmp(C, "L Stru F Ty Extn T1 Rst")
Dim E As Drs: E = D
    Dim Dry(): Dry = E.Dry
    Dim J%
    Const IStru% = 1
    Const IFld% = 2
    Const ITy% = 3
    Const IExtn$ = 4
    Const IT1% = 5
    Const IRst% = 6
    Dim Stru$, F$, Ty$, Extn$, T1$, Rst$, Dr(), S$
    For J = 0 To UB(Dry)
        Dr = Dry(J)
        T1 = Dr(IT1)
        Rst = Dr(IRst)
        '
        Stru = Mid(T1, 6)
        F = T1zS(Rst):  S = RmvT1(Rst)
        Ty = T1zS(S):   S = RmvT1(S)
        Extn = S
        '
        Dr(IStru) = Stru
        Dr(IFld) = F
        Dr(ITy) = Ty
        Dr(IExtn) = RmvSqBkt(Extn)
        Dry(J) = Dr
    Next
    E.Dry = Dry

B_JStruF = DrszFF(JStruF_, KeepFstNColzDrs(E, 5).Dry)
End Function

Private Sub ZZ_ErzLnk(): B ErzLnk(Y_InpFilSrc, Y_LnkImpSrc): End Sub

Private Function B_JFxTbl(ILnkImp As Drs) As Drs
'L Lin IsHdr IsRmk T1 Rst  ILnkImp
'L T Fxn Wsn Stru          JFxTbl
Dim A As Drs: A = DrswColEq(ILnkImp, "T1", "FxTbl")
Dim B As Drs: B = DrswColEq(A, "IsHdr", False)
Dim C As Drs: C = SelDrsAlwEmp(B, "L T Fxn Wsn Stru Rst")
Dim D As Drs: D = C
    Dim Dry(): Dry = D.Dry
    Dim J%
    Const IT% = 1
    Const IFxn% = 2
    Const IWsn% = 3
    Const IStru$ = 4
    Const IRst% = 5
    Dim T$, Fxn$, Wsn$, Stru$
    Dim Rst$
    Dim Dr(), S$, S1$
    For J = 0 To UB(Dry)
        Dr = Dry(J)
        Rst = Dr(IRst)
        '
        T = T1zS(Rst):          S = RmvT1(S)
        S1 = T1zS(S)
        Fxn = BefDotOrAll(S1)
        Wsn = AftDot(S1):       S = RmvT1(S)
        Stru = S
        '
        If Fxn = "" Then Fxn = T
        If Wsn = "" Then Wsn = "Sheet1"
        If Stru = "" Then Stru = T
        '
        Dr(IStru) = Stru
        Dr(IT) = T
        Dr(IFxn) = Fxn
        Dr(IWsn) = Wsn
        Dr(IStru) = Stru
        Dry(J) = Dr
    Next
    D.Dry = Dry
    BrwDrs D
    Stop
B_JFxTbl = DrszFF(JFxTbl_, KeepFstNColzDrs(D, 5).Dry)
End Function
Private Function B_JFbTbl(ILnkImp As Drs) As Drs
Dim Dry()
B_JFbTbl = DrszFF(JFbTbl_, Dry)
End Function
Private Function B_JTblWh(ILnkImp As Drs) As Drs
Dim Dry()
B_JTblWh = DrszFF(JTblWh_, Dry)
End Function
Private Function B_TTbl(JFxTbl As Drs, JFbTbl As Drs) As Drs

End Function
Private Function B_TEFxTbl() As Drs

End Function
Private Function B_TEFbTbl() As Drs

End Function
Private Function B_TAFxTbl() As Drs

End Function
Private Function B_TAFbTbl() As Drs

End Function
Function ErzLnk(InpFilSrc$(), LnkImpSrc$()) As String()
ThwIf_KFsEr KFs(InpFilSrc), CSub

Dim ILnkImp  As Drs: ILnkImp = B_ILnkImp(LnkImpSrc)
Dim IInpFil  As Drs: IInpFil = B_IInpFil(InpFilSrc)
Dim JStru       As Drs:      JStru = B_JStru(ILnkImp)
Dim JStruF      As Drs:     JStruF = B_JStruF(ILnkImp)
Dim JFxTbl      As Drs:     JFxTbl = B_JFxTbl(ILnkImp)
Dim JFbTbl      As Drs:     JFbTbl = B_JFbTbl(ILnkImp)
Dim JTblWh      As Drs:     JTblWh = B_JTblWh(ILnkImp)
Dim TTbl        As Drs:       TTbl = B_TTbl(JFxTbl, JFbTbl)
Dim TEFxTbl     As Drs:    TEFxTbl = B_TEFxTbl()
Dim TEFbTbl     As Drs:     JFbTbl = B_TEFbTbl()
Dim TAFxTbl     As Drs:    TAFxTbl = B_TAFxTbl()
Dim TAFbTbl     As Drs:    TAFbTbl = B_TAFbTbl()
Dim OEr         As Drs:        OEr = DrszF(OEr_)

ApdDrsSub OEr, B_OU_FfnNFnd(IInpFil)
ApdDrsSub OEr, B_OU_FbtNFnd
ApdDrsSub OEr, B_OU_WsnNFnd
ApdDrsSub OEr, B_OU_FxExtnNFnd
ApdDrsSub OEr, B_OU_FxFldTyNMch

If HasReczDrs(OEr) Then ErzLnk = ErLy(OEr, InpFilSrc, LnkImpSrc): Exit Function
ApdDrsSub OEr, B_OS_FldDup
ApdDrsSub OEr, B_OS_StruDup
ApdDrsSub OEr, B_OS_FldTyEr
ApdDrsSub OEr, B_OS_StruExcess
ApdDrsSub OEr, B_OS_StruInUse
ApdDrsSub OEr, B_OS_FxFldTyMis
ApdDrsSub OEr, B_OS_FldMis
ApdDrsSub OEr, B_OX_StruNDef
ApdDrsSub OEr, B_OX_FxtDup
ApdDrsSub OEr, B_OX_FxnNDef
ApdDrsSub OEr, B_OB_StruNDef
ApdDrsSub OEr, B_OB_FbnNFnd
ApdDrsSub OEr, B_OB_FbtMis
ApdDrsSub OEr, B_OB_FbtMis
ApdDrsSub OEr, B_OB_FbnMis
ApdDrsSub OEr, B_OW_TblDup
ApdDrsSub OEr, B_OW_TblNDef
ApdDrsSub OEr, B_OW_TblExcess
ErzLnk = ErLy(OEr, InpFilSrc, LnkImpSrc)
End Function
Private Function ErLy(OEr As Drs, InpFilSrc$(), LnkImpSrc$()) As String()
Dim O$()
Dim IsEr As Boolean: IsEr = HasReczDrs(OEr)
PushIAy O, IfNmLy(IsEr, "InpFilSrc", InpFilSrc, EiBeg1)
PushIAy O, IfNmLy(IsEr, "LnkImpSrc", LnkImpSrc, EiBeg1)
ErLy = O
End Function

Sub Z6()
ZZ_ErzLnk
End Sub
Private Function B_ILnkImp(LnkImpSrc$()) As Drs
Dim Dry(), L&, Lin, TLin$, T1$, Rst$, IsHdr As Boolean, IsRmk As Boolean, LasT1$
For Each Lin In Itr(LnkImpSrc)
    IsRmk = FstTwoChr(LTrim(Lin)) = "--"
    IsHdr = FstChr(Lin) <> " "
    L = L + 1
    
    Select Case True
    Case IsRmk: T1 = "": Rst = ""
    Case IsHdr: T1 = T1zS(Lin): Rst = "": LasT1 = T1
    Case Else:  T1 = LasT1:     Rst = Trim(Lin)
    End Select
    PushI Dry, Av(L, Lin, IsHdr, IsRmk, T1, Rst)
Next
B_ILnkImp = DrszFF(ILnkImp_, Dry) 'L Lin IsHdr IsRmk T1 Rst"
End Function
Private Function B_IInpFil(InpFilSrc$()) As Drs
Dim Dry(), L&, Lin
For Each Lin In Itr(InpFilSrc)
    L = L + 1
    PushI Dry, Av(L, Lin, T1zS(Lin), RmvT1(Lin))
Next
B_IInpFil = DrszFF(IInpFil_, Dry) 'L Lin Fn Ffn
End Function

Private Property Get Y_InpFilSrc() As String()
Erase XX
X "DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X "ZHT0    C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X "MB52    C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X "Uom     C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X "GLBal   C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
Y_InpFilSrc = XX
Erase XX
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

