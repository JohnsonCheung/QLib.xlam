Attribute VB_Name = "QDao_Lnk_ThwIf_ErOfLnkImpSrc"
Private Type MisExtn
    Stru As String
    MisFny() As String
End Type
Private Type MisExtns: N As Long: Ay() As MisExtn: End Type
Private Type DupFld
    Stru As String
    Fld As String
    Lnoss As String
End Type
Private Type DupFlds: N As Long: Ay() As DupFld: End Type
Private Type StruLnkCol
    Stru As String
    LnkCol As Lnxs
End Type
Private Type StruLnkCols: N As Byte: Ay() As StruLnkCol: End Type

Private Const M_Inp_DupFfn$ = "Dup Fbxn[?] at Lno[?]"
Private Const M_Inp_DupNm$ = "Dup Fbxn[?] at Lno[?]"
Private Const M_Stru_DupFld$ = ""
Private Const M_Stru_DupStru$ = ""
Private Const M_Stru_ErFldTy$ = ""
Private Const M_Stru_ExcessStru$ = ""
Private Const M_Stru_MisExtNm$ = ""
Private Const M_Stru_MisFldTy$ = ""
Private Const M_Stru_NoFld$ = "Lno[?] Stru[?] has no field"
Private Const M_FxTbl_MisWsn$ = ""
Private Const M_FxTbl_MisStru$ = ""
Private Const M_FxTbl_MisFxn$ = ""
Private Const M_FxTbl_DupFxt$ = ""
Private Const M_FxTbl_DupFxn$ = ""
Private Const M_FbTbl_MisStru$ = ""
Private Const M_FbTbl_MisFbn$ = ""
Private Const M_FbTbl_DupFbt$ = ""
Private Const M_FbTbl_DupFbn$ = ""
Private Const M_TblWh_DupTbl$ = ""
Private Const M_TblWh_MisTn$ = ""
Private Const M_TblWh_ExcessTn$ = ""



Private Function B_TblWh_MisTn(MisTny$(), Lno&()) As String()
Dim J%
For J = 0 To UB(MisTn)
    With MisTn.Ay(J)
    PushI B_TblWh_MisTn, FmtQQ(M_TblWh_MisTn, MisTny(J), Lno(J))
    End With
Next
End Function

Private Property Get C_CMSrc() As String()
Erase XX
X "Stru_NoFld"
X "Inp_DupFbxn"

End Property
Sub ThwIf_LnkImpPmEr(InpFilSrc$(), LnkImpSrc$())
Dim Inp As KFs: Inp = KFs(InpFilSrc)
                         ThwIf_MisKFs Inp, CSub

Dim a___EI_ErMsgFor_Inp$
    Dim InpLnxs As Lnxs: InpLnxs = Lnxs(InpFilSrc)
    Dim DupNm$(), NmLnoss$()
    Dim DupFfn$(), FfnLnoss$()
    Dim I1$(), I2$()
    I1 = B_Inp_Dup(DupNm, NmLnoss, M_Inp_DupNm)
    I2 = B_Inp_Dup(DupFfn, FfnLnoss, M_Inp_DupFfn)
    Dim a___EI$
    Dim EI$(): EI = Sy(I1, I2) '<== EI

Dim a__LISpec__LnkImpSpec_WhichIsKLys$
    Dim LISpec As KLxs
Dim a__StruSy$
    Dim StruSy$(): 'StruSy = RmvPfxzAy(KAy_FmKLys_WhKPfx(LISpec, "Stru."))
    
Dim Tny$()

Dim a___ES_ErMsg_ForStru
    Dim IsNoStru As Boolean: IsNoStru = Si(StruSy) = 0
    Dim LnkCol As Lnxses:    LnkCol = B_LnkCol(StruSy, LnkImpSrc)
    Dim DupFlds As DupFlds: 'DupFlds = B_DupFlds(StruSy, LnkCol)
    Dim MisExtn$()
    Dim NoFldStru As Lnxs
    
    Dim S1$(), S2$(), S3$(), S4$(), S5$(), S6$(), S7$, S8$()
    S1 = B_Stru_DupStru(LnkImpSrc, StruSy)
    S2 = B_Stru_DupFld(DupFlds)
    S3 = B_Stru_ErFldTy
    S4 = B_Stru_ExcessStru
    S5 = B_Stru_MisExtNm(MisExtn)
    S6 = B_Stru_MisFldTy
    S7 = B_Stru_NoStru(IsNoStru)
    S8 = B_Stru_NoFld(NoFldStru)
    Dim a___ES$
    Dim ES$(): ES = Sy(S1, S2, S3, S4, S5, S6, S7, S8) '<= ES

Dim a___EB_ErMsgFor_FbTbl
    Dim B1$(), B2$()
    B1 = B_FbTbl_DupFbt
    B2 = B_FbTbl_DupFbn
    Dim a___EB$
    Dim EB$(): EB = Sy(B1, B2) '<== EB
    
Dim a___EX_ErMsgFor_FxTbl
    Dim X1$(), X2$(), X3$(), X4$()
    X1 = B_FxTbl_DupFxt
    X2 = B_FxTbl_MisFxn
    X3 = B_FxTbl_MisStru
    X4 = B_FxTbl_MisWsn
    Dim a___EX$
    Dim EX$(): EX = Sy(X1, X2, X3, X4)  '<== EX

Dim a___EW_Pre

Dim a___EW_ErMsgFor_TblWh
    Dim WhTny$():
    Dim MisWhTny$():             MisWhTny = MinusAy(WhTny, Tny)
    Dim MisWhTblLnossSy$: ' MisWhTblLnossSy = C_xMisWhTblLnossSy(MisWhTny)
    Dim DupWhTny$():             DupWhTny = AywDup(WhTny)
    Dim DupWhTblLnossSy$: 'DupWhTblLnossSy = C_xMisWhTblLnossSy(DupWhTny)
    Dim MisTny$(), Lno&()
    Dim W1$():      W1 = B_TblWh_MisTn(MisTny, Lno)
    Dim W2$():      W2 = B_TblWh_DupTbl
    Dim a___EW$
    Dim EW$():      EW = Sy(W1, W2) '<== EW
    
Dim a___Er_ErMsgFor_All
    Dim E$():   E = Sy(EI, ES, EB, EX, EW)
    Dim a___Er$
    Dim Er$(): Er = E   '<==
                    B_PushSrc_IfEr Er, InpFilSrc, "InpFil"
                    B_PushSrc_IfEr Er, LnkImpSrc, "LnkImp"
ThwIf_Er Er, CSub
End Sub

Private Function B_FxTbl_MisWsn() As String()

End Function

Private Function B_FxTbl_MisStru() As String()

End Function

Private Function B_FbTbl_DupFbn() As String()

End Function

Private Function B_FbTbl_DupFbt() As String()

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
Private Function B_DupFlds(A As StruLnkCols) As DupFlds
Dim J%
For J = 0 To A.N - 1
'    Stru = I
Next
End Function

Private Function B_Inp_Dup(Dup$(), LnossSy$(), MsgQQ$) As String()
Dim J%
For J = 0 To UB(Dup)
    PushI B_Inp_Dup, FmtQQ(MsgQQ, Dup(J), LnossSy(J))
Next
End Function

Private Function B_StruSyzNoFld(StruSy$()) As String()
Dim I
For Each I In StruSy
'    If Si(B_Ly("Stru." & I)) = 0 Then PushI B_StruSyzNoFld, I
Next
End Function
Private Function B_Stru_NoFld(NoFldStru As Lnxs) As String()
Dim I
'For Each I In Itr(StruSyzNoFld)
'    PushI B_Stru_NoFld, FmtQQ(Msg_Stru_NoFld, JnSpc(B_LnoAyzStru(LnkImpSrc, CStr(I))), I)
'Next
End Function
Private Function B_Stru_NoStru$(IsNoStru As Boolean)
If IsNoStru Then B_Stru_NoStru = M_Stru_NoStru
End Function
Private Function B_Stru_DupFld(A As DupFlds) As String()
Dim J%
For J = 0 To A.N - 1
    PushI B_Stru_DupFld, W_LinzDupFld(A.Ay(J))
Next
End Function
Private Function W_LinzDupFld$(A As DupFld)

End Function
Private Function B_TblWh_DupTbl() As String()

End Function
Private Function B_Stru_DupFld_PerStru$(Stru$, StruIx As Fei)
Dim A As Fei
'If Not B_Stru_DupFld_FeiHasDupFld(Stru) Then Exit Function

End Function

Private Function B_FxTbl_DupFxt() As String()
End Function
Private Function B_FxTbl_MisFxn() As String()
End Function

Private Function B_Stru_DupStru(LnkImpSrc$(), StruSy$()) As String()
Dim Dup$(): Dup = AywDup(StruSy)
Dim LnoAy&(), I
For Each I In Itr(Dup)
    LnoAy = B_LnoAyzStru(LnkImpSrc, CStr(I))
    PushI B_Stru_DupStru, FmtQQ("Dup Stru[?] at Lno#[?]", I, JnSpc(LnoAy))
Next
End Function


Private Function B_Stru_ErFldTy() As String()
End Function

Private Function B_Stru_ExcessStru() As String()

End Function


Private Function B_Stru_MisExtNm(MisExtny$()) As String()
End Function

Private Function B_Stru_MisFldTy() As String()
End Function


Private Function B_LnoAyzStru(LnkImpSrc$(), Stru$) As Long()
Dim J%, S$
For J = 0 To UB(LnkImpSrc)
    S = "Stru." & Stru
    If HasPfx(LnkImpSrc(J), S) Then
        PushI B_LnoAyzStru, J + 1
    End If
Next
End Function

Private Sub B_PushSrc_IfEr(OEr$(), Src$(), SrcKd$)
If Si(OEr) = 0 Then Exit Sub
PushI OEr, SrcKd
PushIAy OEr, TabSy(AddIxPfx(Src, 1))
End Sub


Sub Z()
ZZ_ThwIf_LnkImpPmEr
End Sub

Private Function B_LnkCol(StruSy$(), ImpLnkSrc$()) As Lnxses

End Function

Private Sub ZZ_ThwIf_LnkImpPmEr()
ThwIf_LnkImpPmEr Y_InpFilSrc, Y_LnkImpSrc
End Sub


Function ChkFxww(Fx, Wsnn$, Optional FxKd$ = "Excel file") As String()
Dim W$, I
'If Not HasFfn(Fx) Then ChkFxww = MsgzMisFfn(Fx, FxKd): Exit Function
For Each I In Ny(Wsnn)
    W = I
    PushIAy ChkFxww, ChkWs(Fx, W, FxKd)
Next
End Function
Function ChkWs(Fx, Wsn, FxKd$) As String()
If HasFxw(Fx, Wsn) Then Exit Function
Dim M$
M = FmtQQ("? does not have expected worksheet", FxKd)
ChkWs = LyzFunMsgNap(CSub, M, "Folder File Expected-Worksheet Worksheets-in-file", Pth(Fx), Fn(Fx), Wsn, Wny(Fx))
End Function
Function ChkFxw(Fx, Wsn, Optional FxKd$ = "Excel file") As String()
ChkFxw = ChkHasFfn(Fx, FxKd): If Si(ChkFxw) > 0 Then Exit Function
ChkFxw = ChkWs(Fx, Wsn, FxKd)
End Function
Function ChkLnkWs(A As Database, T, Fx, Wsn, Optional FxKd$ = "Excel file") As String()
Const CSub$ = CMod & "ChkLnkWs"
Dim O$()
    O = ChkFxw(Fx, Wsn, FxKd)
    If Si(O) > 0 Then
        ChkLnkWs = O
        Exit Function
    End If
On Error GoTo X
LnkFx A, T, Fx, Wsn
Exit Function
X: ChkLnkWs = _
    LyzMsgNap("Error in linking Xls file", "Er LnkFx LnkWs ToDb AsTbl", Err.Description, Fx, Wsn, Dbn(A), T)
End Function


