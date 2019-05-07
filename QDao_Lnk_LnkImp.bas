Attribute VB_Name = "QDao_Lnk_LnkImp"
Option Explicit
Private Const CMod$ = "BLnkImp."
Private Type DupFld
    Stru As String
    Fld As String
    Lnoss As String
End Type
Private Type DupFlds: N As Integer: Ay() As DupFld: End Type
Private Type FxtRec
    T As String
    Fxn As String
    Wsn As String
    Stru As String
End Type
Private Type FxtRecs: N As Byte: Ay() As FxtRec: End Type
Private Const Msg_Stru_NoFld$ = "Lno[?] Stru[?] has no field"

Sub Z_LnkImp()
Dim InpFilSrc$(), LnkImpSrc$(), Db As Database
GoSub T0
Exit Sub
T0:
    InpFilSrc = SampSrczInpFilzTaxAlert
    LnkImpSrc = SampSrczLnkImpzTaxAlert
    Set Db = TmpDb
    GoTo Tst
Tst:
    LnkImp InpFilSrc, LnkImpSrc, Db
    Return
End Sub

Private Function B_DupFlds(StruSy$(), LnkColLnxsAy() As Lnxs) As DupFlds

End Function

Private Function B_Inp_DupFbxn(DupFbxn$(), LnossSy$()) As String()
Const C$ = "Dup Fbxn[?] at Lno[?]"
Dim J%
For J = 0 To UB(DupFbxn)
    PushI B_Inp_DupFbxn, FmtQQ(C, DupFbxn(J), LnossSy(J))
Next
End Function

Private Function B_Inp_DupFbx(DupFbx$(), LnossSy$()) As String()
Const C$ = "Dup Fil[?] at Lno[?]"
Dim J%
For J = 0 To UB(DupFbx)
    PushI B_Inp_DupFbx, FmtQQ(C, DupFbx(J), LnossSy(J))
Next
End Function

Sub LnkImp(InpFilSrc$(), LnkImpSrc$(), Db As Database)
Dim Inp As FilKds: Inp = FilKdszLy(InpFilSrc)
                         ThwIfMisFilKds Inp, CSub

Dim TblWhLy$():         TblWhLy = IndentedLy(LnkImpSrc, "Tbl.Where")
Dim FxTblLy$():         FxTblLy = IndentedLy(LnkImpSrc, "FxTbl")
Dim FxtRecs As FxtRecs: FxtRecs = B_FxtRecs(FxTblLy)
Dim FbTblLy$():         FbTblLy = IndentedLy(LnkImpSrc, "FbTbl")
Dim Fbt$():                 Fbt = B_Fbt(FbTblLy)
Dim Fxt$():                 Fxt = B_Fxt(FxtRecs)
Dim Tny$():                 Tny = AddSy(Fbt, Fxt)
Dim StruSy() As String:  StruSy = B_StruSy(LnkImpSrc)
Dim FbxnSy$():           FbxnSy = KdSyzFilKds(Inp)
Dim FbxSy$():             FbxSy = FfnSyzFilKds(Inp)
'--
Dim DupFbxn$(): DupFbxn = AywDup(FbxnSy)
Dim DupFbx$():   DupFbx = AywDup(FbxSy)

Dim LnossSyzDupFbxn$(): LnossSyzDupFbxn = LnossSy(AySubAy(FbxnSy, DupFbxn))
Dim LnossSyzDupFbx$():   LnossSyzDupFbx = LnossSy(AySubAy(FbxSy, DupFbx))
'--
Dim StruSyzNoFld$()
Dim WhTny$():
Dim MisWhTny$():             MisWhTny = MinusAy(WhTny, Tny)
Dim MisWhTblLnossSy$: ' MisWhTblLnossSy = C_xMisWhTblLnossSy(MisWhTny)
Dim DupWhTny$():             DupWhTny = AywDup(WhTny)
Dim DupWhTblLnossSy$: 'DupWhTblLnossSy = C_xMisWhTblLnossSy(DupWhTny)
'--
Dim LnkColLnxsAy() As Lnxs: LnkColLnxsAy = B_LnkColLnxsAy(StruSy, LnkImpSrc)
Dim DupFlds As DupFlds: DupFlds = B_DupFlds(StruSy, LnkColLnxsAy)
'
Dim I1$(): I1 = B_Inp_DupFbxn(DupFbxn, LnossSyzDupFbxn)
Dim I2$(): I2 = B_Inp_DupFbx(DupFbx, LnossSyzDupFbx)
Dim S1$(): S1 = B_Stru_DupStru(LnkImpSrc, StruSy)
Dim S2$(): S2 = B_Stru_DupFld(DupFlds)
Dim S3$(): S3 = B_Stru_ErFldTy
Dim S4$(): S4 = B_Stru_ExcessStru
Dim S5$(): S5 = B_Stru_MisExtNm
Dim S6$(): S6 = B_Stru_MisFldTy
Dim S7$(): S7 = B_Stru_NoStru(StruSy)
Dim S8$(): S8 = B_Stru_NoFld(LnkImpSrc, StruSyzNoFld)
Dim B1$(): B1 = B_FbTbl_DupFbt
Dim B2$(): B2 = B_FbTbl_DupFbn
Dim X1$(): X1 = B_FxTbl_DupFxt
Dim X2$(): X2 = B_FxTbl_MisFxn
Dim X3$(): X3 = B_FxTbl_MisStru
Dim X4$(): X4 = B_FxTbl_MisWsn
Dim W1$(): W1 = B_TblWh_MisTn
Dim W2$(): W2 = B_TblWh_DupTbl
Dim EI$(): EI = AddSyAp(I1, I2)
Dim ES$(): ES = AddSyAp(S1, S2, S3, S4, S5, S6, S7, S8)
Dim EB$(): EB = AddSyAp(B1, B2)
Dim EX$(): EX = AddSyAp(X1, X2, X3, X4)
Dim EW$(): EW = AddSyAp(W1, W2)
Dim Er$(): Er = AddSyAp(EI, ES, EB, EX, EW)
           Er = B_AddSrcIf(Er, InpFilSrc, "InpFil")
           Er = B_AddSrcIf(Er, LnkImpSrc, "LnkImp")
ThwIfEr Er, CSub
'----
Dim FbxnzDiFbFx As Dictionary: Set FbxnzDiFbFx = Dic(InpFilSrc)
Dim FbtzDiFbn As Dictionary:     Set FbtzDiFbn = DiczVkkLy(FbTblLy)
Dim FxtzDiFxn As Dictionary:     Set FxtzDiFxn = B_FxtzDiFxn(FxtRecs)
Dim TzDiFbxn As Dictionary:       Set TzDiFbxn = AddDic(FxtzDiFxn, FbtzDiFbn)
Dim TzDiFbFx As Dictionary:       Set TzDiFbFx = AzDiC(TzDiFbxn, FbxnzDiFbFx)
'----
Dim TzDiBexpr As Dictionary:       Set TzDiBexpr = Dic(TblWhLy)
Dim TzDiStru As Dictionary:         Set TzDiStru = B_TzDiStru(Fbt, FxtRecs)
Dim TzDiLnkColLy As Dictionary: Set TzDiLnkColLy = B_TzDiLnkColLy(LnkImpSrc, TzDiStru, StruSy)
'-----------------
Dim ImpSqy$(): ImpSqy = B_ImpSqy(Tny, TzDiLnkColLy, TzDiBexpr)
Dim FxLnk As LnkTblPms: FxLnk = B_FxLnkPms(FxtRecs, TzDiFbFx)
Dim FbLnk As LnkTblPms: FbLnk = B_FbLnkPms(FbTblLy, TzDiFbFx)
'-----------------
LnkTblzPms Db, FxLnk '<==============
LnkTblzPms Db, FbLnk '<==============
RunSqy Db, ImpSqy    '<==============

Debug.Print "NRec"
Debug.Print UnderLin("NRec")
Dim T$, I
For Each I In Tny
    T = "#I" & I
    Debug.Print NReczT(Db, T), T
Next
End Sub
Private Function B_LnkColLnxsAy(StruSy$(), ImpLnkSrc$()) As Lnxs()

End Function
Private Function B_AddSrcIf(IfEr$(), Src$(), SrcKd$) As String()
If Si(IfEr) = 0 Then Exit Function
B_AddSrcIf = AddSyAp(IfEr, Sy(SrcKd), TabSy(AddIxPfx(Src, 1)))
End Function
Private Function B_TzDiStru(Fbt$(), A As FxtRecs) As Dictionary
Set B_TzDiStru = New Dictionary
Dim T, J%
For Each T In Itr(Fbt)
    B_TzDiStru.Add T, T
Next
For J = 0 To A.N - 1
    With A.Ay(J)
        B_TzDiStru.Add .T, .Stru
    End With
Next
End Function
Private Function B_FbTbl_DupFbt() As String()

End Function
Private Function B_FbTbl_DupFbn() As String()

End Function

Private Function B_FxtzDiFbFx(A As FxtRecs, TzDiFbFx As Dictionary) As Dictionary
Dim J%
Set B_TzDiFbFx = New Dictionary
For J = 0 To A.N - 1
    With A.Ay(J)
        B_FxtzDiFbFx.Add .T, TzDiFbFx(.Fxn)
    End With
Next
End Function
Function AzDiC(AzDiB As Dictionary, BzDiC As Dictionary) As Dictionary
Dim A, B, C
Set AzDiC = New Dictionary
For Each A In AzDiB.Keys
    B = AzDiB(A)
    If Not BzDiC.Exists(B) Then Thw CSub, "BzDiC does not contain B", "A B AzDiB BzDiC", A, B, AzDiB, BzDiC
    C = BzDiC(B)
    AzDiC.Add A, C
Next
End Function
Private Function B_FxtzDiFxn(A As FxtRecs) As Dictionary
Set B_FxtzDiFxn = New Dictionary
Dim J%
For J = 0 To A.N - 1
    With A.Ay(J)
    B_FxtzDiFxn.Add .T, .Fxn
    End With
Next
End Function

Private Function B_FxTnToFxDic(A As FxtRecs, FbxnToFbFxDic As Dictionary) As Dictionary

End Function
Private Function B_Fxt(A As FxtRecs) As String()
Dim J%
For J = 0 To A.N - 1
    PushI B_Fxt, A.Ay(J).T
Next
End Function
Private Function B_Fbt(FbTblLy$()) As String()
Dim J%
For J = 0 To UB(FbTblLy)
    PushIAy B_Fbt, SyzSsLin(RmvT1(FbTblLy(J)))
Next
End Function

Private Function B_ImpSqy(Tny$(), TzDiLnkColLy As Dictionary, TzDiBexpr As Dictionary) As String()
Dim J%, Fny$(), Ey$(), T$, Into$, LnkColLy$(), Bexpr$
For J = 0 To UB(Tny)
    T = ">" & Tny(J)
    Into = "#I" & Tny(J)
    LnkColLy = ValzDic(TzDiLnkColLy, Tny(J), "TzDiLnkColLy", "TblNm")
    Fny = T1Sy(LnkColLy)
    Ey = RmvSqBktzSy(RmvTTzSy(LnkColLy))
    Bexpr = ValzDicIf(TzDiBexpr, T)
    PushI B_ImpSqy, SqlSel_Fny_ExtNy_Into_T(Fny, Ey, Into, T, Bexpr)
Next
End Function

Private Function B_FxtRecs(FxTblLy$()) As FxtRecs
Dim OAy() As FxtRec, J%, L$, A$
For J = 0 To UB(FxTblLy)
    L = FxTblLy(J)
    ReDim Preserve OAy(J)
    With OAy(J)
        .T = ShfT1(L)
        A = ShfT1(L)
        .Fxn = B_Fxn(A, .T)
        .Wsn = B_Wsn(A)
        .Stru = StrDft(L, .T)
    End With
Next
B_FxtRecs.N = Si(FxTblLy)
B_FxtRecs.Ay = OAy
End Function
Private Function B_Wsn(Fxn_dot_Wsn$)
Dim A$: A = Fxn_dot_Wsn
If A = "" Then B_Wsn = "Sheet1": Exit Function
If Not HasDot(A) Then B_Wsn = "Sheet1": Exit Function
B_Wsn = AftDot(A)
End Function
Private Function B_Fxn(Fxn_dot_Wsn$, T$)
Dim A$: A = Fxn_dot_Wsn
If A = "" Then B_Fxn = T: Exit Function
If HasDot(A) Then B_Fxn = BefDot(A): Exit Function
B_Fxn = Fxn_dot_Wsn
End Function
Private Function B_FbLnkPms(FbTblLy$(), TzDiFbFx As Dictionary) As LnkTblPms
Dim J%, Fbn$, A$, I, T$, Cn$
For J = 0 To UB(FbTblLy)
    AsgBrk FbTblLy(J), " ", Fbn, A
    If Not TzDiFbFx.Exist(Fbn) Then Thw CSub, "TzDiFbFx does not contains Fbn", "Fbn TzDiFbFx FbTblLin", Fbn, TzDiFbFx, FbTblLy(J)
    Cn = CnStrzFbzAsDao(TzDiFbFx(Fbn))
    For Each I In SyzSsLin(A)
        T = I
        PushLnkTblPm B_FbLnkPms, LnkTblPm(T, T, Cn)
    Next
Next
End Function

Private Function B_FxLnkPms(A As FxtRecs, TzDiFbFx As Dictionary) As LnkTblPms
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
    Cn = CnStrzFxDAO(Fx)
    PushLnkTblPm B_FxLnkPms, LnkTblPm(">" & .T, S, Cn)
    End With
Next
End Function

Private Function B_StruSy(LnkImpSrc$()) As String()
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
B_StruSy = Y
End Function
Private Function B_SrcTblDic(Tny$(), WsDic As Dictionary) As Dictionary
Dim J%, O As New Dictionary, T$, S$
For J = 0 To UB(Tny)
    T = Tny(J)
    If WsDic.Exists(T) Then
        S = T & "$"
    Else
        S = T
    End If
    O.Add T, S
Next
End Function
Private Function B_CnStrDic(Tny$(), FilDic As Dictionary) As Dictionary
Dim J%, O As New Dictionary
For J = 0 To UB(Tny)
    T = Tny(J)
    O.Add T, CnStrzFxAdo(FilDic(T))
    PushI SrcNy, Wsny(J) & "$"
Next
Set B_CnStrDic = O
End Function
Private Function B_TzDiLnkColLy(LnkImpSrc$(), TzDiStru As Dictionary, StruSy$()) As Dictionary
Dim T, Stru$
Set B_TzDiLnkColLy = New Dictionary
For Each T In TzDiStru.Keys
    Stru = TzDiStru(T)
    B_TzDiLnkColLy.Add T, IndentedLy(LnkImpSrc, "Stru." & Stru)
Next
End Function
Private Function B_Stru_ErFldTy() As String()
End Function
Private Function B_Stru_MisExtNm() As String()
End Function
Private Function B_Stru_MisFldTy() As String()
End Function
Private Function B_Stru_ExcessStru() As String()

End Function
Private Function B_Stru_DupStru(LnkImpSrc$(), StruSy$()) As String()
Dim Dup$(): Dup = AywDup(StruSy)
Dim LnoAy&(), I
For Each I In Itr(Dup)
    LnoAy = B_LnoAyzStru(LnkImpSrc, CStr(I))
    PushI B_Stru_DupStru, FmtQQ("Dup Stru[?] at Lno#[?]", I, JnSpc(LnoAy))
Next
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
Private Function B_StruSyzNoFld(StruSy$()) As String()
Dim I
For Each I In StruSy
    If Si(B_Ly("Stru." & I)) = 0 Then PushI B_StruSyzNoFld, I
Next
End Function
Private Function B_Stru_NoFld(LnkImpSrc$(), StruSyzNoFld$()) As String()
Dim I
For Each I In Itr(StruSyzNoFld)
    PushI B_Stru_NoFld, FmtQQ(Msg_Stru_NoFld, JnSpc(B_LnoAyzStru(LnkImpSrc, CStr(I))), I)
Next
End Function
Private Function B_Stru_NoStru(StruSy$()) As String()
If Si(StruSy) > 0 Then Exit Function
B_Stru_NoStru = Sy("There is no Stru.XXX")
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
Private Function B_Stru_DupFld_PerStru$(Stru$, StruIx As FTIx)
Dim A As FTIx
If Not B_Stru_DupFld_FTIxHasDupFld(Stru) Then Exit Function

End Function

Private Function B_FxTbl_DupFxt() As String()
End Function
Private Function B_FxTbl_MisFxn() As String()
End Function
Private Function B_FxTbl_MisStru() As String()
End Function
Private Function B_FxTbl_MisWsn() As String()
End Function
Private Function B_TblWh_MisTn() As String()
End Function
Private Property Get SampSrczLnkImp() As String()
Erase XX
X "# T   FxNm.Ws   Stru"
X "Fxt"
X " Z86  ZHT0.8601 ZHT0"
X " Z87  ZHT0.8701 ZHT0"
X " Uom  ZHT0      Uom"
X "Stru.ZHT0"
X " ZHT0   D Brand  "
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
SampSrczLnkImp = XX
Erase XX
End Property
Private Property Get SampSrczLnkImp2() As String()
'    Sql: select
'[Material] As Sku,
'[Amount] As CurRateAc, [Valid From] As VdtFm, [Valid to] As VdtTo, [Unit] As HKD, [per] As Per, [Uom] As CA_Uom into [#I>ZHT086] from [>ZHT086] (String)
Erase XX
X "FxTbl T  FxNm.Wsn   Stru"
X " ZHT087 ZHT0.8701   ZHT0"
X " ZHT086 ZHT0.8601   ZHT0"
X " MB52                   "
X " UOM                    "
X "FbTbl FbNm TblNm...     "
X " Duty Permit PermitD"
X "Stru.Permit"
X "Stru.PermitD"
X "Stru.ZHT0"
X " ZHT0   M Brand"
X " RateSc D Amount"
X " VdtFm  M Valid From"
X " VdtTo  M Valid to"
X "Stru.Uom"
X " Sku     M Material"
X " Des     M Material Description"
X " Sc_U    M SC "
X " StkUom  M Base Unit of Measure"
X " Topaz   M Topaz Code"
X " ProdH   M Product hierarchy"
X "Stru.MB52"
X " Sku    M Material"
X " Whs    M Plant"
X " QInsp  D In Quality Insp#"
X " QUnRes D Unrestricted"
X " QBlk   D Blocked"
SampSrczLnkImp2 = XX
Erase XX
End Property
Private Property Get SampSrczInpFilzTaxAlert() As String()
Erase XX
X "ZHT0  C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X "MB52  C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X "Uom   C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X "GLBal C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
SampSrczInpFilzTaxAlert = XX
Erase XX
SampSrczInpFil
End Property


Private Property Get SampSrczLnkImpzTaxAlert() As String()
Erase XX
X "FxTbl T  FxNm.Wsn  Stru"
X " ZHT086  ZHT0.8600 ZHT0"
X " ZHT087  ZHT0.8700 ZHT0"
X " MB52                  "
X " Uom                   "
X " GLBal"
X "Tbl.Where"
X " MB52 Plant='8601' and [Storage Location] in ('0002','')"
X " Uom  Plant='8601'"
X "Stru.ZHT0"
X " Sku       Txt Material    "
X " CurRateAc Dbl [     Amount]"
X " VdtFm     Txt Valid From  "
X " VdtTo     Txt Valid to    "
X " HKD       Txt Unit        "
X " Per       Txt per         "
X " CA_Uom    Txt Uom         "
X "Stru.MB52"
X "#"
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
SampSrczLnkImpzTaxAlert = XX
Erase XX
End Property

