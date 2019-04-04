Attribute VB_Name = "A"
'QLib.Cls.ActLin.Friend Function Init(Act As eActLin, Lin$, Lno&) As ActLin
'QLib.Cls.ActLin.Property Get Ix&()
'QLib.Cls.ActLin.Property Get ActStr$()
'QLib.Cls.ActLin.Function ToStr$()
'QLib.Cls.ActMd.Friend Function Init(Md As CodeModule, ActLin() As ActLin) As ActMd
'QLib.Cls.ActMd.Function ActLinAy() As ActLin()
'QLib.Cls.ActMd.Function Hdr$()
'QLib.Cls.ActMd.Function ToFmt(Optional NoHdr As Boolean) As String()
'QLib.Cls.ActMd.Function ToLy() As String()
'QLib.Cls.Arg.Property Get ToStr$()
'QLib.Cls.Arg.Property Get ShtStr$()
'QLib.Cls.Aset.Property Get TermLin$()
'QLib.Cls.Aset.Property Get Cnt&()
'QLib.Cls.Aset.Sub Dmp()
'QLib.Cls.Aset.Sub Vc()
'QLib.Cls.Aset.Sub Brw(Optional Fnn$, Optional UseVc As Boolean)
'QLib.Cls.Aset.Function Srt() As Aset
'QLib.Cls.Aset.Function AddAset(A As Aset) As Aset
'QLib.Cls.Aset.Function RmvItm(Itm) As Aset
'QLib.Cls.Aset.Sub PushItm(Itm)
'QLib.Cls.Aset.Sub PushAy(A)
'QLib.Cls.Aset.Sub PushItr(Itr, Optional NoBlankStr As Boolean)
'QLib.Cls.Aset.Function Clone() As Aset
'QLib.Cls.Aset.Function Minus(B As Aset) As Aset
'QLib.Cls.Aset.Function Has(Itm) As Boolean
'QLib.Cls.Aset.Function IsEq(B As Aset) As Boolean
'QLib.Cls.Aset.Function IsEmp() As Boolean
'QLib.Cls.Aset.Function Av() As Variant()
'QLib.Cls.Aset.Function IsInOrdEq(B As Aset) As Boolean
'QLib.Cls.Aset.Function FstItm()
'QLib.Cls.Aset.Function AbcDic() As Dictionary 'AbcDic means the keys is comming from Aset the value is starting from A, B, C
'QLib.Cls.Aset.Function Itms()
'QLib.Cls.Aset.Function Lin$()
'QLib.Cls.Aset.Sub PushAset(A As Aset)
'QLib.Cls.Aset.Function Sy() As String()
'QLib.Cls.Aset.Private Sub ZZ()
'QLib.Cls.Aset.Private Sub Z()
'QLib.Std.AShpCst_Pm_LiPm.Property Get ShpCstLiPm() As LiPm
'QLib.Std.AShpCst_Pm_LiPm.Property Get ShpCstLtPm() As LtPm()
'QLib.Std.AShpCst_Pm_LiPm.Private Function LiFil(Itm$) As LiFil
'QLib.Std.AShpCst_Pm_LiPm.Private Property Get LiFilAy() As LiFil()
'QLib.Std.AShpCst_Pm_LiPm.Private Property Get LiFbAy() As LiFb()
'QLib.Std.AShpCst_Pm_LiPm.Private Property Get LiFxAy() As LiFx()
'QLib.Std.AShpCst_Rpt.Function RptFb$()
'QLib.Std.AShpCst_Rpt.Function RptDb() As Database
'QLib.Std.AShpCst_Rpt.Function RptAppDb() As Database
'QLib.Std.AShpCst_Rpt.Sub GenOupTbl(Apn$)
'QLib.Std.AShpCst_Rpt.Function OupFxzShpCst$()
'QLib.Std.AShpCst_Rpt.Sub BrwRptPm()
'QLib.Std.AShpCst_Rpt.Sub DocUOM (A, B)
'QLib.Std.AShpCst_Rpt.Private Sub GenOMain()
'QLib.Std.AShpCst_Rpt.Private Sub GenORate()
'QLib.Std.AShpCst_Rpt.Private Function ErzMB52MissingWhs8601Or8701(FxMB52$, Wsn$) As String()
'QLib.Std.AShpCst_Rpt.Function PnmStkDte(AppDb As Database) As Date
'QLib.Std.AShpCst_Rpt.Function PnmStkYYMD$(AppDb As Database)
'QLib.Std.AShpCst_Rpt.Sub ShpCstBrwLiAct()
'QLib.Std.AShpCst_Rpt.Property Get ShpCstLiAct() As LiAct
'QLib.Cls.AyAB.Property Get A()
'QLib.Cls.AyAB.Property Get B()
'QLib.Cls.AyABC.Friend Function Init(A, B, C) As AyABC
'QLib.Cls.AyABC.Property Get A()
'QLib.Cls.AyABC.Property Get B()
'QLib.Cls.AyABC.Property Get C()
'QLib.Cls.CmpCnt.Friend Function Init(NMod%, NCls%, NDoc%, NOth%) As CmpCnt
'QLib.Cls.CmpCnt.Property Get NCmp%()
'QLib.Cls.CmpCnt.Function Lin$(Optional Hdr As eHdr)
'QLib.Cls.Drs.Friend Function Init(Fny0, Dry()) As Drs
'QLib.Cls.Drs.Property Get Fny() As String()
'QLib.Cls.Drs.Property Get Dry() As Variant()
'QLib.Cls.Ds.Property Get DtAy() As Dt()
'QLib.Cls.Ds.Function Init(A() As Dt, Optional DsNm$ = "Ds") As Ds
'QLib.Cls.Ds.Sub Brw(Optional MaxColWdt% = 100, Optional DtBrkColDicVbl$, Optional NoIxCol As Boolean)
'QLib.Cls.Ds.Sub Dmp()
'QLib.Cls.Ds.Property Get UDt%()
'QLib.Cls.Ds.Function Dt(Ix%) As Dt
'QLib.Cls.Ds.Function Fmt(Optional MaxColWdt% = 100, Optional DtBrkColDicVbl$, Optional NoIxCol As Boolean) As String()
'QLib.Cls.Dt.Property Get Dry() As Variant()
'QLib.Cls.Dt.Property Get Fny() As String()
'QLib.Cls.Dt.Friend Function Init(DtNm, Fny0, Dry()) As Dt
'QLib.Cls.FTIx.Property Get FmIx&()
'QLib.Cls.FTIx.Property Get ToIx&()
'QLib.Cls.FTIx.Friend Function Init(FmIx, ToIx) As FTIx
'QLib.Cls.FTIx.Property Get IsEmp() As Boolean
'QLib.Cls.FTIx.Property Get Cnt&()
'QLib.Cls.FTIx.Property Get FmNo&()
'QLib.Cls.FTIx.Property Get ToNo&()
'QLib.Cls.FTIx.Property Get IsVdt() As Boolean
'QLib.Cls.Gp.Friend Function Init(LnxAy() As Lnx) As Gp
'QLib.Cls.Gp.Property Get LnxAy() As Lnx()
'QLib.Cls.LiAct.Friend Function Init(Fx() As LiActFx, Fb() As LiActFb) As LiAct
'QLib.Cls.LiAct.Property Get Fx() As LiActFx()
'QLib.Cls.LiAct.Property Get Fb() As LiActFb()
'QLib.Cls.LiActFb.Friend Function Init(Fb$, Fbn$, T$, Fset As Aset) As LiActFb
'QLib.Cls.LiActFx.Friend Function Init(Fx, Fxn, Wsn, ShtTyDic As Dictionary) As LiActFx
'QLib.Cls.LiActFx.Property Get Fset() As Aset
'QLib.Cls.LidFb.Friend Function Init(Fbn, T, Fset As Aset, Bexpr$, Fb) As LidFb
'QLib.Cls.LidFil.Friend Function Init(FilNm$, Ffn$) As LidFil
'QLib.Cls.LidFx.Friend Function Init(Fxn$, Wsn$, T$, Fxc() As LidFxc, Bexpr$) As LidFx
'QLib.Cls.LidFx.Property Get Fxc() As LidFxc()
'QLib.Cls.LidFxc.Friend Function Init(ColNm$, ShtTyLis$, ExtNm$) As LidFxc
'QLib.Cls.LidMis.Friend Function Init(Ffn As Aset, Tbl() As LidMisTbl, Col() As LidMisCol, Ty() As LidMisTy) As LidMis
'QLib.Cls.LidMis.Property Get Ty() As LidMisTy()
'QLib.Cls.LidMis.Property Get Tbl() As LidMisTbl()
'QLib.Cls.LidMis.Property Get Col() As LidMisCol()
'QLib.Cls.LidMisCol.Friend Function Init(Ffn, T, EptFset As Aset, ActFset As Aset, Optional Wsn) As LidMisCol
'QLib.Cls.LidMisCol.Property Get MisMsg() As String()
'QLib.Cls.LidMisTbl.Friend Function Init(Ffn, FilNm, T, Optional Wsn$) As LidMisTbl
'QLib.Cls.LidMisTbl.Property Get MisMsg$()
'QLib.Cls.LidMisTy.Friend Function Init(Fx, Fxn, Wsn, Tyc() As LidMisTyc) As LidMisTy
'QLib.Cls.LidMisTy.Property Get TycAy() As LidMisTyc()
'QLib.Cls.LidMisTy.Property Get MisMsg() As String()
'QLib.Cls.LidMisTy.Private Function MisMsgTyOneFx(ColMsg$()) As String()
'QLib.Cls.LidMisTy.Private Function MisMsgColMsgAy(A() As LidMisTyc) As String()
'QLib.Cls.LidPm.Friend Function Init(Apn$, Fil() As LidFil, Fx() As LidFx, Fb() As LidFb, Optional CpyInpWsToOupFx As Boolean) As LidPm
'QLib.Cls.LidPm.Property Get AppFb$()
'QLib.Cls.LidPm.Property Get Fil() As LidFil()
'QLib.Cls.LidPm.Property Get Fx() As LidFx()
'QLib.Cls.LidPm.Property Get Fb() As LidFb()
'QLib.Cls.LiFb.Friend Function Init(Fbn, T, Fset As Aset, Bexpr) As LiFb
'QLib.Cls.LiFb.Friend Function ExistFb$(A() As LiActFb)
'QLib.Cls.LiFil.Friend Function Init(FilNm, Ffn) As LiFil
'QLib.Cls.LiFx.Friend Function Init(Fxn, Wsn, T, Fxc() As LiFxc, Bexpr) As LiFx
'QLib.Cls.LiFx.Property Get FxcAy() As LiFxc()
'QLib.Cls.LiFx.Property Get Fset() As Aset
'QLib.Cls.LiFx.Property Get Fny() As String()
'QLib.Cls.LiFx.Function EptFset() As Aset
'QLib.Cls.LiFx.Function ExtNy() As String()
'QLib.Cls.LiFx.Friend Function ExistFx$(B() As LiActFx)
'QLib.Cls.LiFxc.Friend Function Init(ColNm$, ShtTyLis$, ExtNm$) As LiFxc
'QLib.Cls.LiMis.Friend Function Init(MisFfn As Aset, MisTbl() As LiMisTbl, MisCol() As LiMisCol, MisTy() As LiMisTy) As LiMis
'QLib.Cls.LiMis.Property Get MisTy() As LiMisTy()
'QLib.Cls.LiMis.Property Get MisTbl() As LiMisTbl()
'QLib.Cls.LiMis.Property Get MisCol() As LiMisCol()
'QLib.Cls.LiMisCol.Friend Function Init(Ffn, T, EptFset As Aset, ActFset As Aset, Optional Wsn) As LiMisCol
'QLib.Cls.LiMisCol.Property Get MisMsg() As String()
'QLib.Cls.LiMisTbl.Friend Function Init(Ffn, FilNm, T, Optional Wsn$) As LiMisTbl
'QLib.Cls.LiMisTbl.Property Get MisMsg$()
'QLib.Cls.LiMisTy.Friend Function Init(Fx, Fxn, Wsn, Tyc() As LiMisTyc) As LiMisTy
'QLib.Cls.LiMisTy.Property Get TycAy() As LiMisTyc()
'QLib.Cls.LiMisTy.Property Get MisMsg() As String()
'QLib.Cls.LiMisTy.Private Function MisMsgTyOneFx(ColMsg$()) As String()
'QLib.Cls.LiMisTy.Private Function MisMsgColMsgAy(A() As LiMisTyc) As String()
'QLib.Cls.LiMisTyc.Friend Function Init(ExtNm$, ActShtTy$, EptShtTyLis$) As LiMisTyc
'QLib.Cls.LiMisTyc.Property Get MisMsg$()
'QLib.Cls.LinPm.Function Init(PmStr$) As LinPm
'QLib.Cls.LinPm.Private Sub PushPmNm(PmNm$)
'QLib.Cls.LinPm.Function WhNm(Optional NmPfx$) As WhNm
'QLib.Cls.LinPm.Function HasSw(SwNm) As Boolean
'QLib.Cls.LinPm.Property Get SwNy() As String()
'QLib.Cls.LinPm.Function SwNm$(Nm$)
'QLib.Cls.LinPm.Property Get Cnt%()
'QLib.Cls.LinPm.Function HasPm(PmNm$) As Boolean
'QLib.Cls.LinPm.Private Sub PushPm(Nm$, Optional V$)
'QLib.Cls.LinPm.Sub Dmp()
'QLib.Cls.LinPm.Function Fmt() As String()
'QLib.Cls.LinPm.Private Function FmtzNmSy$(PmNm, Sy$())
'QLib.Cls.LinPm.Function LikeAy(NmPfx) As String()
'QLib.Cls.LinPm.Function ExlLikAy(NmPfx) As String()
'QLib.Cls.LinPm.Function SyPmVal(PmNm, Optional NmPfx$) As String()
'QLib.Cls.LinPm.Function StrPmVal$(PmNm, Optional NmPfx$)
'QLib.Cls.LinPm.Function Patn$(NmPfx)
'QLib.Cls.LinPm.Private Sub Class_Initialize()
'QLib.Cls.LiPm.Friend Function Init(Apn$, Fil() As LiFil, Fx() As LiFx, Fb() As LiFb) As LiPm
'QLib.Cls.LiPm.Property Get Fil() As LiFil()
'QLib.Cls.LiPm.Property Get Fx() As LiFx()
'QLib.Cls.LiPm.Property Get Fb() As LiFb()
'QLib.Cls.LiPm.Property Get FfnAy() As String()
'QLib.Cls.LiPm.Property Get ExistFfn() As Aset
'QLib.Cls.LiPm.Property Get MisFfn() As Aset
'QLib.Cls.LiPm.Property Get ExistFilNmToFfnDic() As Dictionary
'QLib.Cls.LiPm.Property Get FmtFil() As String()
'QLib.Cls.LiPm.Function FilNmToFfnDic() As Dictionary
'QLib.Cls.LiPm.Sub Brw()
'QLib.Cls.LiPm.Function Ds() As Ds
'QLib.Cls.Lnx.Friend Sub Init(Lin, Ix&)
'QLib.Cls.Lnx.Property Get Lno&()
'QLib.Cls.LtPm.Friend Function Init(T, S, Cn) As LtPm
'QLib.Cls.LtPm.Property Get ToStr$()
'QLib.Std.MAcs.Sub DoFrm(A As Access.Application, FrmNm$)
'QLib.Std.MAcs.Sub BrwTbl(D As Database, T)
'QLib.Std.MAcs.Sub BrwTT(D As Database, TT)
'QLib.Std.MAcs.Function CAcs(D As Database) As Access.Application
'QLib.Std.MAcs.Sub SavRec()
'QLib.Std.MAcs.Function FbzAcs$(A As Access.Application)
'QLib.Std.MAcs.Sub ClsDbzAcs(A As Access.Application)
'QLib.Std.MAcs.Sub BrwFb(Fb)
'QLib.Std.MAcs.Sub ClsTTz(A As Access.Application, TT)
'QLib.Std.MAcs.Sub ClsTblz(A As Access.Application, T)
'QLib.Std.MAcs.Sub ClsAllTblz(A As Access.Application)
'QLib.Std.MAcs.Sub QuitzA(A As Access.Application)
'QLib.Std.MAcs.Function AcsVis(A As Access.Application) As Access.Application
'QLib.Std.MAcs.Function CvAcs(A) As Access.Application
'QLib.Std.MAcs.Property Get Acs() As Access.Application
'QLib.Std.MAcs.Sub CpyAllAcsFrm(A As Access.Application, Fb$)
'QLib.Std.MAcs.Sub CpyAcsMd(A As Access.Application, ToFb$)
'QLib.Std.MAcs.Sub CpyAcsObj(A As Access.Application, ToFb$)
'QLib.Std.MAcs.Sub TxtbSelPth(A As Access.TextBox)
'QLib.Std.MAcs.Sub CmdTurnOffTabStop(AcsCtl As Access.Control)
'QLib.Std.MAcs.Sub ClrMainMsg()
'QLib.Std.MAcs.Sub SetMainMsgzQnm(QryNm)
'QLib.Std.MAcs.Sub SetMainMsg(A$)
'QLib.Std.MAcs.Private Property Get MMBox() As Access.TextBox
'QLib.Std.MAcs.Private Property Get MFrm() As Access.Form
'QLib.Std.MAcs.Private Sub ZZ()
'QLib.Std.MAcs.Sub FrmSetCmdNotTabStop(A As Access.Form)
'QLib.Std.MAcs.Function CvCtl(A) As Access.Control
'QLib.Std.MAcs.Function CvBtn(A) As Access.CommandButton
'QLib.Std.MAcs.Function CvTgl(A) As Access.ToggleButton
'QLib.Std.MAcs.Sub SetTBox(A As Access.TextBox, Msg$)
'QLib.Std.MAcs.Sub AcsQuit(A As Access.Application)
'QLib.Std.MAcs.Function NewAcs(Optional Shw As Boolean) As Access.Application
'QLib.Std.MAcs.Function DbNmzAcs$(A As Access.Application)
'QLib.Std.MAcs.Sub OpnFb(A As Access.Application, Fb)
'QLib.Std.MAcs.Function DftAcs(A As Access.Application) As Access.Application
'QLib.Std.MAcs_USysRegInfo.Sub CrtTblzUSysRegInf(A As Database)
'QLib.Std.MAcs_USysRegInfo.Sub EnsTblzUSysRegInf(A As Database)
'QLib.Std.MAcs_USysRegInfo.Sub InstallAddin(A As Database, Fb$, Optional AutoFunNm$ = "AutoExec")
'QLib.Std.MApp_Git.Sub GitCmit(Optional Msg$ = "commit", Optional ReInit As Boolean)
'QLib.Std.MApp_Git.Sub GitPush()
'QLib.Std.MApp_Git.Private Function GitCmitCdLines$(CmitgPth, Msg$, ReInit As Boolean)
'QLib.Std.MApp_Git.Private Sub Z_FcmdWaitzCdLines()
'QLib.Std.MApp_Git.Private Function FcmdWaitzCdLines$(CdLines)
'QLib.Std.MApp_Git.Function HasInternet() As Boolean
'QLib.Std.MApp_Git.Private Function GitPushCdLines$(CmitgPth)
'QLib.Std.MApp_Git.Sub BrwGitCmitCdLines()
'QLib.Std.MApp_Git.Sub BrwGitPushCdLines()
'QLib.Std.MApp_Git.Private Function PjNm$(CmitgPth)
'QLib.Std.MApp_Git.Private Sub XX1()
'QLib.Std.MApp_Git.Private Sub Z()
'QLib.Std.MApp_Git.Sub PowerRun(Ps1, ParamArray PmAp())
'QLib.Std.MApp_Rpt.Function OupFxzLidPm$(A As LidPm) 'Gen&Vis OupFx using LidPm as NxtFfn.
'QLib.Std.MApp_Rpt.Private Sub ClsOupWb(Apn$)
'QLib.Std.MApp_Rpt.Private Function OupWbzNxt(Apn$) As Workbook
'QLib.Std.MApp_Rpt.Sub CpyWszLidPm(A As LidPm, ToOupWb As Workbook, Optional Vis As Boolean)
'QLib.Std.MApp_EApp.Property Get MHDAppFbDic() As Dictionary
'QLib.Std.MApp_EApp.Property Get AppFbAy() As String()
'QLib.Std.MApp_Fun.Function AppDb(Apn) As Database
'QLib.Std.MApp_Fun.Function OupFxzNxt$(Apn)
'QLib.Std.MApp_Fun.Function OupFx$(Apn)
'QLib.Std.MApp_Fun.Function AppFb$(Apn)
'QLib.Std.MApp_Fun.Property Get AppHom$()
'QLib.Std.MApp_Fun.Sub Ens()
'QLib.Std.MApp_Fun.Property Get AutoExec()
'QLib.Std.MApp_Fun.Function DocLy(DclLy$()) As String()
'QLib.Std.MApp_Fun.Function StrConstVal$(Lin)
'QLib.Std.MApp_Fun.Function ConstNm$(Lin)
'QLib.Std.MApp_Fun.Function DocDicOfPj() As Dictionary
'QLib.Std.MApp_Fun.Function IsDocNm(S) As Boolean
'QLib.Std.MApp_Fun.Sub AsgStrConstNmAaaVal(OStrConstNm$, OStrConstVal$, Lin)
'QLib.Std.MApp_Fun.Function DocDiczDcl(Dcl) As Dictionary
'QLib.Std.MApp_Fun.Function DocDiczPj(A As VBProject) As Dictionary
'QLib.Std.MApp_Fun.Sub Doc(Nm$)
'QLib.Std.MApp_Fun.Property Get IsDev() As Boolean
'QLib.Std.MApp_Fun.Property Get IsProd() As Boolean
'QLib.Std.MApp_Fun.Function PgmDb_DtaDb(A As Database) As Database
'QLib.Std.MApp_Fun.Function PgmDb_DtaFb$(A As Database)
'QLib.Std.MApp_Fun.Property Get ProdPth$()
'QLib.Std.MApp_Fun.Private Sub ZZ()
'QLib.Std.MApp_Fun.Property Let ApnzDb(A As Database, V$)
'QLib.Std.MApp_Fun.Property Get ApnzDb$(A As Database)
'QLib.Std.MApp_NDrive.Sub MapNDrive()
'QLib.Std.MApp_NDrive.Sub RmvNDrive()
'QLib.Std.MApp_Pm.Function PnmOupPth$(A As Database)
'QLib.Std.MApp_Pm.Function PnmPth$(Db As Database, Pnm)
'QLib.Std.MApp_Pm.Function PnmFn$(Db As Database, Pnm)
'QLib.Std.MApp_Pm.Function PnmFfn(Db As Database, Pnm)
'QLib.Std.MApp_Pm.Property Get PnmVal$(Db As Database, Pnm$)
'QLib.Std.MApp_Pm.Property Let PnmVal(Db As Database, Pnm$, V$)
'QLib.Std.MApp_Pm.Sub BrwTblPm(Apn$)
'QLib.Std.MApp_Pm.Private Sub ZZ()
'QLib.Std.MApp_Pm.Private Sub Z()
'QLib.Std.MApp_SalRpt.Property Get DftSrpDic() As Dictionary
'QLib.Std.MApp_Tp.Function TpFn$(Apn) 'Fst Fn in Tbl.Fld.Ssk-Att.Att.Tp
'QLib.Std.MApp_Tp.Function Tp$(Apn)
'QLib.Std.MApp_Tp.Function HasTp(Apn) As Boolean
'QLib.Std.MApp_Tp.Sub ImpTp(Apn)
'QLib.Std.MApp_Tp.Private Function TpFx$(Apn)
'QLib.Std.MApp_Tp.Private Function TpFxm$(Apn)
'QLib.Std.MApp_Tp.Sub OpnTp(Apn)
'QLib.Std.MApp_Tp.Property Get TpWsCdNy() As String()
'QLib.Std.MApp_Tp.Property Get TpPth$()
'QLib.Std.MApp_Tp.Sub RfhTp(Apn)
'QLib.Std.MApp_Tp.Sub RfhWcTp(Apn)
'QLib.Std.MApp_Tp.Function TpWb(Apn) As Workbook
'QLib.Std.MApp_Tp.Function TpWcSy(Apn) As String()
'QLib.Std.MApp_Tp.Sub ExpTp(Apn$, ToFfn$)
'QLib.Std.MApp_Wrk.Property Get W() As Database
'QLib.Std.MApp_Wrk.Sub WCls()
'QLib.Std.MApp_Wrk.Sub WIniOpn(Apn$)
'QLib.Std.MApp_Wrk.Sub WIni(Apn$)
'QLib.Std.MApp_Wrk.Sub WOpn(Apn$)
'QLib.Std.MApp_Wrk.Sub WRun(QQ, ParamArray Ap())
'QLib.Std.MApp_Wrk.Sub WDrp(TT)
'QLib.Std.MApp_Wrk.Sub WBrw(Apn$)
'QLib.Std.MApp_Wrk.Function WAcs() As Access.Application
'QLib.Std.MApp_Wrk.Function WPth$(Apn)
'QLib.Std.MApp_Wrk.Function WFb$(Apn)
'QLib.Std.MDao_Ado.Function CvTc(A) As ADOX.Table
'QLib.Std.MDao_Ado.Function NoReczAdo(A As ADODB.Recordset) As Boolean
'QLib.Std.MDao_Ado.Function HasReczAdo(A As ADODB.Recordset) As Boolean
'QLib.Std.MDao_Ado.Function TnyzCat(A As Catalog) As String()
'QLib.Std.MDao_Ado.Function CatzFb(Fb) As Catalog
'QLib.Std.MDao_Ado.Function CatCn(A As ADODB.Connection) As Catalog
'QLib.Std.MDao_Ado.Function CatzFx(Fx) As Catalog
'QLib.Std.MDao_Ado.Function FnyCatTbl(Cat As ADOX.Catalog, T) As String()
'QLib.Std.MDao_Ado.Function DrsFxw(Fx, Wsn) As Drs
'QLib.Std.MDao_Ado.Function ArsFxw(Fx, Wsn) As ADODB.Recordset
'QLib.Std.MDao_Ado.Sub RunCnSqy(A As ADODB.Connection, Sqy$())
'QLib.Std.MDao_Ado.Private Sub Z_DrsCnq()
'QLib.Std.MDao_Ado.Function ArsCnq(A As ADODB.Connection, Q) As ADODB.Recordset
'QLib.Std.MDao_Ado.Function DrsCnq(A As ADODB.Connection, Q) As Drs
'QLib.Std.MDao_Ado.Function DrsFbqAdo(A$, Q$) As Drs
'QLib.Std.MDao_Ado.Private Sub Z_DrsFbqAdo()
'QLib.Std.MDao_Ado.Function ArszFbq(Fb$, Q$) As ADODB.Recordset
'QLib.Std.MDao_Ado.Function DrsArs(A As ADODB.Recordset) As Drs
'QLib.Std.MDao_Ado.Function DryzArs(A As ADODB.Recordset) As Variant()
'QLib.Std.MDao_Ado.Private Sub Z_DryArs()
'QLib.Std.MDao_Ado.Function FnyzArs(A As ADODB.Recordset) As String()
'QLib.Std.MDao_Ado.Function IntAyzArs(A As ADODB.Recordset, Optional Col = 0) As Integer()
'QLib.Std.MDao_Ado.Private Function HasCatT(A As Catalog, T) As Boolean
'QLib.Std.MDao_Ado.Private Sub Z_TnyzFb()
'QLib.Std.MDao_Ado.Private Sub Z_WsNyzFx()
'QLib.Std.MDao_Ado.Function HasTblzFfnTblNm(Ffn, TblNm) As Boolean
'QLib.Std.MDao_Ado.Function HasFbt(Fb, T) As Boolean
'QLib.Std.MDao_Ado.Function HasFxw(Fx, W) As Boolean
'QLib.Std.MDao_Ado.Function TnyzFb(Fb) As String()
'QLib.Std.MDao_Ado.Function TnyzAdoFb(Fb) As String()
'QLib.Std.MDao_Ado.Function WsNyzFx(Fx) As String()
'QLib.Std.MDao_Ado.Function FnyzFbt(Fb, T) As String()
'QLib.Std.MDao_Ado.Private Sub Z_CnStrzFbAdo()
'QLib.Std.MDao_Ado.Private Sub Z_Cn()
'QLib.Std.MDao_Ado.Function Cn(AdoCnStr) As ADODB.Connection
'QLib.Std.MDao_Ado.Function DftWsNy(WsNy0, Fx$) As String()
'QLib.Std.MDao_Ado.Function DftTny(Tny0, Fb$) As String()
'QLib.Std.MDao_Ado.Function FxDftWsNy(A, WsNy0) As String()
'QLib.Std.MDao_Ado.Function FxDftWsn$(A, Wsn0$)
'QLib.Std.MDao_Ado.Function CnStrzFbAdo$(A)
'QLib.Std.MDao_Ado.Function CnStrzFxAdo$(A)
'QLib.Std.MDao_Ado.Function DtaSrczScl(DtaSrcScl$)
'QLib.Std.MDao_Ado.Function DtaSrc$(A As Database, T)
'QLib.Std.MDao_Ado.Function CnzFx(Fx) As ADODB.Connection
'QLib.Std.MDao_Ado.Function CnzFb(A) As ADODB.Connection
'QLib.Std.MDao_Ado.Private Sub Z_CnzFb()
'QLib.Std.MDao_Ado.Function FFzFxw$(Fx$, Wsn$)
'QLib.Std.MDao_Ado.Function FnyzFfnTblNm(Ffn, TblNm) As String()
'QLib.Std.MDao_Ado.Function FnyzFxw(Fx, W) As String()
'QLib.Std.MDao_Ado.Function CvAdoTy(A) As ADODB.DataTypeEnum
'QLib.Std.MDao_Ado.Function CatT$(Wsn)
'QLib.Std.MDao_Ado.Function WsnzCatT$(CatT)
'QLib.Std.MDao_Ado.Private Sub Z()
'QLib.Std.MDao_Ado.Function IntoColzArs(A As ADODB.Recordset, Into, Optional Col = 0)
'QLib.Std.MDao_Ado.Function SyzArs(A As ADODB.Recordset, Optional Col = 0) As String()
'QLib.Std.MDao_Ado.Sub ArunzFbq(A$, Q$)
'QLib.Std.MDao_Ado.Private Sub Z_ArunzFbq()
'QLib.Std.MDao_Ado.Function DrzAfds(A As ADODB.Fields, Optional N%) As Variant()
'QLib.Std.MDao_Ado.Function FnyzAfds(A As ADODB.Fields) As String()
'QLib.Std.MDao_Att.Function FstAttFfn$(A As Database, Att)
'QLib.Std.MDao_Att.Function FnyzAttFld(A As Database) As String()
'QLib.Std.MDao_Att.Function IsOldAtt(A As Database, Att$, Ffn$) As Boolean
'QLib.Std.MDao_Att.Function AttSz&(A As Database, Att)
'QLib.Std.MDao_Att.Function AttTim(A As Database, Att) As Date
'QLib.Std.MDao_Att.Function AttFilCntzAttd%(A As Attd)
'QLib.Std.MDao_Att.Function AttFilCnt%(Db As Database, Att)
'QLib.Std.MDao_Att.Function AttFnAy(A As Database, Att) As String()
'QLib.Std.MDao_Att.Function FnyzTblAtt(A As Database) As String()
'QLib.Std.MDao_Att.Function AttFn$(A As Database, Att)
'QLib.Std.MDao_Att.Function HasOneFilAtt(A As Database, Att) As Boolean
'QLib.Std.MDao_Att.Function AttNy(A As Database) As String()
'QLib.Std.MDao_Att.Private Sub Z_AttFnAy()
'QLib.Std.MDao_Att.Private Sub Z()
'QLib.Std.MDao_Att.Function AttNm$(A As Attd)
'QLib.Std.MDao_Att.Function AttFnzAttd$(A As Attd)
'QLib.Std.MDao_Att.Function Attd(A As Database, Att) As Attd
'QLib.Std.MDao_Att_Op_Dlt.Sub DltAtt(A As Database, Att)
'QLib.Std.MDao_Att_Op_Exp.Private Function ExpAttzAttd$(A As Attd, ToFfn) 'Export the only File in {Attds} {ToFfn}
'QLib.Std.MDao_Att_Op_Exp.Function ExpAtt$(Db As Database, Att, ToFfn) 'Exporting the first File in [Att] to [ToFfn]. |If no or more than one file in att, error |If any, export and return ToFfn
'QLib.Std.MDao_Att_Op_Exp.Function ExpAttzFn$(A As Database, Att$, AttFn$, ToFfn)
'QLib.Std.MDao_Att_Op_Exp.Private Function AttFd2(A As Database, Att, AttFn) As Dao.Field2
'QLib.Std.MDao_Att_Op_Exp.Private Sub ZZ_ExpAtt()
'QLib.Std.MDao_Att_Op_Exp.Private Sub Z()
'QLib.Std.MDao_Att_Op_Exp.Private Sub ZZ()
'QLib.Std.MDao_Att_Op_Imp.Private Sub ImpAttzAttd(A As Attd, Ffn$)
'QLib.Std.MDao_Att_Op_Imp.Sub ImpAtt(Db As Database, Att, FmFfn$)
'QLib.Std.MDao_Att_Op_Imp.Private Sub Z_ImpAtt()
'QLib.Std.MDao_Att_Op_Imp.Private Sub Z()
'QLib.Std.MDao_Ccm.Private Sub Z_LnkCcm()
'QLib.Std.MDao_Ccm.Sub LnkCcm(Db As Database, IsLcl As Boolean)
'QLib.Std.MDao_Ccm.Private Sub LnkCcmzTny(Db As Database, CcmTny$(), IsLcl As Boolean)
'QLib.Std.MDao_Ccm.Private Sub Chk(Db As Database, CcmTny$(), IsLcl As Boolean)
'QLib.Std.MDao_Ccm.Private Function Chk1(A As Database, CcmTny$()) As String()
'QLib.Std.MDao_Ccm.Private Sub Chk2(Db As Database, CcmTny$())
'QLib.Std.MDao_Ccm.Private Function Chk3(Db As Database, CcmTbl) As String()
'QLib.Std.MDao_Ccm.Private Function CcmTny(Db As Database) As String()
'QLib.Std.MDao_Ccm.Private Sub Z_CcmTny()
'QLib.Std.MDao_Ccm.Private Sub Z()
'QLib.Std.MDao_Chk_PkSk.Function ChkPk$(A As Database, T)
'QLib.Std.MDao_Chk_PkSk.Function ChkSsk$(A As Database, T)
'QLib.Std.MDao_Chk_PkSk.Function ChkPkSk(A As Database) As String()
'QLib.Std.MDao_Chk_PkSk.Function ChkPkSkzT(A As Database, T) As String()
'QLib.Std.MDao_Chk_PkSk.Function ChkSk$(A As Database, T)
'QLib.Std.MDao_CnStr.Function CnStrzFbDao$(A)
'QLib.Std.MDao_CnStr.Function CnStrzFxDAO$(A)
'QLib.Std.MDao_CnStr.Function CnStrzFbAdoOle$(A) 'Return a connection used as WbConnection
'QLib.Std.MDao_CnStr.Function CnStrzFbForWbCn$(Fb$)
'QLib.Std.MDao_Const.Private Property Get LnkSpecTp$()
'QLib.Std.MDao_Db.Function IsOkDb(A As Database) As Boolean
'QLib.Std.MDao_Db.Sub AddTmpTbl(A As Database)
'QLib.Std.MDao_Db.Function PthzDb$(A As Database)
'QLib.Std.MDao_Db.Function IsTmpDb(A As Database) As Boolean
'QLib.Std.MDao_Db.Sub DrpDbIfTmp(A As Database)
'QLib.Std.MDao_Db.Sub BrwDb(A As Database)
'QLib.Std.MDao_Db.Function TnyzTT(TT) As String()
'QLib.Std.MDao_Db.Function StruzTT(A As Database, TT)
'QLib.Std.MDao_Db.Function Stru(Db As Database) As String()
'QLib.Std.MDao_Db.Function OupTny(A As Database) As String()
'QLib.Std.MDao_Db.Sub DrpTT(A As Database, TT, Optional NoReOpn As Boolean)
'QLib.Std.MDao_Db.Sub DrpTmp(A As Database)
'QLib.Std.MDao_Db.Sub DrpT(A As Database, T, Optional NoReOpn As Boolean)
'QLib.Std.MDao_Db.Sub CrtTbl(A As Database, T$, FldDclAy)
'QLib.Std.MDao_Db.Function DszDb(A As Database, Optional DsNm$) As Ds
'QLib.Std.MDao_Db.Function DszTT(A As Database, TT, Optional DsNm$) As Ds
'QLib.Std.MDao_Db.Sub EnsTmpTblz(A As Database)
'QLib.Std.MDao_Db.Sub RunQ(A As Database, Q)
'QLib.Std.MDao_Db.Sub RunQQAv(A As Database, QQ, Av())
'QLib.Std.MDao_Db.Sub RunQQ(A As Database, QQ, ParamArray Ap())
'QLib.Std.MDao_Db.Function RszQQ(A As Database, QQ, ParamArray Ap()) As Dao.Recordset
'QLib.Std.MDao_Db.Function Rs(A As Database, Q) As Dao.Recordset
'QLib.Std.MDao_Db.Function HasReczQ(A As Database, Q) As Boolean
'QLib.Std.MDao_Db.Function HasQryz(A As Database, Q) As Boolean
'QLib.Std.MDao_Db.Function HasTbl(A As Database, T, Optional ReOpn As Boolean)
'QLib.Std.MDao_Db.Function HasTblzMSysObjDb(A As Database, T) As Boolean
'QLib.Std.MDao_Db.Function IsDbOk(A As Database) As Boolean
'QLib.Std.MDao_Db.Function DbPth$(A As Database)
'QLib.Std.MDao_Db.Function Qny(A As Database) As String()
'QLib.Std.MDao_Db.Function RszQry(A As Database, Qry) As Dao.Recordset
'QLib.Std.MDao_Db.Function SrcTnAy(A As Database) As String()
'QLib.Std.MDao_Db.Function TmpTny(A As Database) As String()
'QLib.Std.MDao_Db.Function Tni(A As Database)
'QLib.Std.MDao_Db.Function Tny(A As Database) As String()
'QLib.Std.MDao_Db.Function TnyzADO(A As Database) As String()
'QLib.Std.MDao_Db.Function TnyzDaoFb(Fb) As String()
'QLib.Std.MDao_Db.Function TnyzDaoDb(A As Database, Optional NoReOpn As Boolean) As String()
'QLib.Std.MDao_Db.Function TnyzMSysObj(A As Database) As String()
'QLib.Std.MDao_Db.Private Sub ZZ_Qny()
'QLib.Std.MDao_Db.Private Sub Z_Ds()
'QLib.Std.MDao_Db.Private Sub ZZ()
'QLib.Std.MDao_Db.Sub RenTblzAddPfx(A As Database, TT, Pfx$)
'QLib.Std.MDao_Db.Function TdStrAy(A As Database, TT) As String()
'QLib.Std.MDao_Db.Sub CrtTblzTmp(A As Database)
'QLib.Std.MDao_Db.Sub RenTbl(A As Database, T, ToNm$)
'QLib.Std.MDao_Db.Sub RenTblzFmPfx(A As Database, FmPfx$, ToPfx$)
'QLib.Std.MDao_Db.Sub DrpzTT(TT)
'QLib.Std.MDao_Db.Sub DrpzAp(ParamArray TblAp())
'QLib.Std.MDao_Db.Property Get TblDes$(A As Database, T)
'QLib.Std.MDao_Db.Property Let TblDes(A As Database, T, Des$)
'QLib.Std.MDao_Db.Property Get TblAttDes$(A As Dao.Database)
'QLib.Std.MDao_Db.Property Let TblAttDes(A As Dao.Database, Des$)
'QLib.Std.MDao_Db.Property Set TblDesDic(A As Database, D As Dictionary)
'QLib.Std.MDao_Db.Property Get FldDesDic(A As Database) As Dictionary
'QLib.Std.MDao_Db.Property Set FldDesDic(A As Database, D As Dictionary)
'QLib.Std.MDao_Db.Sub ClsDb(Db As Database)
'QLib.Std.MDao_Db.Property Get TblDesDic(A As Database) As Dictionary
'QLib.Std.MDao_Db.Function JnStrDiczTwoColSql(TwoColSql) As Dictionary 'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Sy
'QLib.Std.MDao_Db.Private Sub ZZ_BrwTbl()
'QLib.Std.MDao_Db.Function TnizInp(A As Database)
'QLib.Std.MDao_Db.Function TnyzInp(A As Database) As String()
'QLib.Std.MDao_Db.Sub CpyInpTblAsTmpz(A As Database)
'QLib.Std.MDao_Db.Function ReOpnDb(A As Database, Optional NoReOpn As Boolean) As Database
'QLib.Std.MDao_Db.Function DbNm$(A As Database)
'QLib.Std.MDao_Db_DbInf.Sub BrwDbInf(A As Database)
'QLib.Std.MDao_Db_DbInf.Function DbInfDs(A As Database) As Ds
'QLib.Std.MDao_Db_DbInf.Private Sub Z_BrwDbInf()
'QLib.Std.MDao_Db_DbInf.Private Sub Z_XTbl()
'QLib.Std.MDao_Db_DbInf.Private Function XTbl(A As Database, Tny$()) As Dt
'QLib.Std.MDao_Db_DbInf.Private Function XLnk(A As Database, Tny$()) As Dt
'QLib.Std.MDao_Db_DbInf.Private Function XPrp(A As Database) As Dt
'QLib.Std.MDao_Db_DbInf.Private Function XFld(A As Database, Tny$()) As Dt
'QLib.Std.MDao_Db_DbInf.Private Function XTblF(D As Database, Tny$()) As Dt
'QLib.Std.MDao_Db_DbInf.Private Function XTblFDry(D As Database, T) As Variant()
'QLib.Std.MDao_Db_DbInf.Private Function XTblFDr(T, Seq%, F As Dao.Field2) As Variant()
'QLib.Std.MDao_Db_DbInf.Private Sub Z()
'QLib.Std.MDao_Db_DbInf.Private Function XLnkLy(A As Database) As String()
'QLib.Std.MDao_Db_DbInf_Stru.Function DbInfDtStru(A As Database) As Dt
'QLib.Std.MDao_Db_DbInf_Stru.Sub DmpStru(A As Database)
'QLib.Std.MDao_Db_DbInf_Stru.Function StruFld(ParamArray Ap()) As Drs
'QLib.Std.MDao_Db_DbInf_Stru.Sub DmpStruTT(A As Database, TT)
'QLib.Std.MDao_Def_Fd.Function FdClone(A As Dao.Field2, FldNm) As Dao.Field2
'QLib.Std.MDao_Def_Fd.Function FdVal(A As Dao.Field)
'QLib.Std.MDao_Def_Fd.Function IsEqFd(A As Dao.Field2, B As Dao.Field2) As Boolean
'QLib.Std.MDao_Def_Fd.Function CvFd(A) As Dao.Field
'QLib.Std.MDao_Def_Fd.Function CvFd2(A As Dao.Field) As Dao.Field2
'QLib.Std.MDao_Def_Fds.Function CsvzFds$(A As Dao.Fields)
'QLib.Std.MDao_Def_Fds.Function NzEmpty(A)
'QLib.Std.MDao_Def_Fds.Function DrzFds(A As Dao.Fields, Optional FF = "") As Variant()
'QLib.Std.MDao_Def_Fds.Function FnyzFds(A As Fields) As String()
'QLib.Std.MDao_Def_Fds.Function VyzFds(A As Dao.Fields, Optional FF = "") As Variant()
'QLib.Std.MDao_Def_Fds.Private Sub Z_DrzFds()
'QLib.Std.MDao_Def_Fds.Private Sub Z_VyzFds()
'QLib.Std.MDao_Def_Fds.Private Sub Z()
'QLib.Std.MDao_Def_Fd_New.Function FdzStr(FdStr$) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function Fd(F, Optional Ty As Dao.DataTypeEnum = dbText, Optional Req As Boolean, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional VRul$, Optional VTxt$) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzBool(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzByt(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzCrtDte(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzCur(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzChr(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzDbl(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzDte(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzDec(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzEle(Ele, F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Private Function FdzTnnn(F, EleTnnn) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzFk(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzId(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzInt(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzLng(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzAtt(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzMem(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzNm(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzPk(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzSng(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzTim(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzShtTys(ShtTys, Fld) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzFld(StdFld, Optional T$) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Private Sub Z_FdzFdStr()
'QLib.Std.MDao_Def_Fd_New.Private Sub Z_FdzFdStr1()
'QLib.Std.MDao_Def_Fd_New.Function FdzFdStr(FdStr) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzTxt(F, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Function FdzTy(F) As Dao.Field2
'QLib.Std.MDao_Def_Fd_New.Private Sub ZZ()
'QLib.Std.MDao_Def_Fd_New.Private Sub Z()
'QLib.Std.MDao_Def_Lin.Function IdxStr$(A As Dao.Index)
'QLib.Std.MDao_Def_Lin.Function IdxStrAyIdxs(A As Dao.Indexes) As String()
'QLib.Std.MDao_Def_Td.Function CvTd(A) As Dao.TableDef
'QLib.Std.MDao_Def_Td.Sub AddFdy(A As TableDef, Fdy() As Dao.Field2)
'QLib.Std.MDao_Def_Td.Sub AddFldzId(A As Dao.TableDef)
'QLib.Std.MDao_Def_Td.Sub AddFldzLng(A As Dao.TableDef, FF)
'QLib.Std.MDao_Def_Td.Sub AddFldzTimstmp(A As Dao.TableDef, F$)
'QLib.Std.MDao_Def_Td.Sub AddFldzTxt(A As Dao.TableDef, FF, Optional Req As Boolean, Optional Si As Byte = 255)
'QLib.Std.MDao_Def_Td.Function FnyzTd(A As Dao.TableDef) As String()
'QLib.Std.MDao_Def_Td.Function IsEqTd(A As Dao.TableDef, B As Dao.TableDef) As Boolean
'QLib.Std.MDao_Def_Td.Sub ThwIfNETd(A As Dao.TableDef, B As Dao.TableDef)
'QLib.Std.MDao_Def_Td.Sub DmpTdAy(TdAy() As Dao.TableDef)
'QLib.Std.MDao_Def_Td.Function TdLyzDb(A As Database) As String()
'QLib.Std.MDao_Def_Td.Function TdLyzT(A As Database, T) As String()
'QLib.Std.MDao_Def_Td.Function TdLy(Td) As String()
'QLib.Std.MDao_Def_Td.Private Function Fdy(FF, T As Dao.DataTypeEnum) As Dao.Field2()
'QLib.Std.MDao_Def_Td.Private Sub ZZ()
'QLib.Std.MDao_Def_Td.Function IsSysTd(A As Dao.TableDef) As Boolean
'QLib.Std.MDao_Def_Td.Function IsHidTd(A As Dao.TableDef) As Boolean
'QLib.Std.MDao_Def_Td_New.Private Function CvIdxfds(A) As Dao.IndexFields
'QLib.Std.MDao_Def_Td_New.Private Function IsIdFd(A As Dao.Field2, T) As Boolean
'QLib.Std.MDao_Def_Td_New.Function NewSkIdx(T As Dao.TableDef, SkFny$()) As Dao.Index
'QLib.Std.MDao_Def_Td_New.Function TdzFdy(T, Fdy() As Field2, Optional SkFF) As Dao.TableDef
'QLib.Std.MDao_Def_Td_New.Private Sub AddPk(A As Dao.TableDef)
'QLib.Std.MDao_Def_Td_New.Private Function NewPkIdx(T) As Dao.Index
'QLib.Std.MDao_Def_Td_New.Private Sub AddSk(A As Dao.TableDef, SkFF)
'QLib.Std.MDao_Def_ToStr.Function FdStrAyFds(A As Dao.Fields) As String()
'QLib.Std.MDao_Def_ToStr.Function TdStr$(A As Dao.TableDef)
'QLib.Std.MDao_Def_ToStr.Function FnyzTdLy(TdLy$()) As String()
'QLib.Std.MDao_Def_ToStr.Function TdStrzT$(A As Database, T)
'QLib.Std.MDao_Def_ToStr.Function SkFnyzTdLin(A) As String()
'QLib.Std.MDao_Def_ToStr.Function FdStr$(A As Dao.Field2)
'QLib.Std.MDao_Dic.Function AyDaoTy(A As Dao.DataTypeEnum)
'QLib.Std.MDao_Dic.Function AyDic_RsKF(A As Dao.Recordset, DicKeyFld, AyFld) As Dictionary 'Return a dictionary of Ay using KeyFld and AyFld.  The Val-of-returned-Dic is Ay using the AyFld.Type to create
'QLib.Std.MDao_Dic.Function JnStrDicTwoFldRs(A As Dao.Recordset, Optional Sep$ = " ") As Dictionary
'QLib.Std.MDao_Dic.Function JnStrDicRsKeyJn(A As Dao.Recordset, KeyFld, JnStrFld, Optional Sep$ = " ") As Dictionary
'QLib.Std.MDao_Dic.Function CntDiczRs(A As Dao.Recordset, Optional Fld = 0) As Dictionary
'QLib.Std.MDao_Fb.Sub CrtFb(Fb)
'QLib.Std.MDao_Fb.Function DbCrt(Fb) As Database
'QLib.Std.MDao_Fb.Private Sub Z_BrwFb()
'QLib.Std.MDao_Fb.Function DbzFb(Fb) As Database
'QLib.Std.MDao_Fb.Function CntrNyzFb(Fb) As String()
'QLib.Std.MDao_Fb.Function CntrItmNyzFb(Fb) As String()
'QLib.Std.MDao_Fb.Function Db(Fb) As Database
'QLib.Std.MDao_Fb.Sub EnsFb(Fb)
'QLib.Std.MDao_Fb.Function OupTnyzFb(Fb) As String()
'QLib.Std.MDao_Fb.Sub AsgFbtStr(FbtStr$, OFb$, OT$)
'QLib.Std.MDao_Fb.Sub DrpvFbt(Fb$, T$)
'QLib.Std.MDao_Fb.Private Sub ZZ_HasFbt()
'QLib.Std.MDao_Fb.Private Sub ZZ_OupTnyzFb()
'QLib.Std.MDao_Fb.Private Sub ZZ_TnyzFb()
'QLib.Std.MDao_Fb.Private Sub Z()
'QLib.Std.MDao_Fb_Fbq.Private Sub Z_WszFbq()
'QLib.Std.MDao_Fb_Fbq.Function WszFbq(Fb, Sql, Optional Wsn$, Optional Vis As Boolean) As Worksheet
'QLib.Std.MDao_Fb_Fbq.Function DrszQ(A As Database, Q) As Drs
'QLib.Std.MDao_Fb_Fbq.Function DrszFbq(Fb, Q) As Drs
'QLib.Std.MDao_Fb_Fbq.Function ArszFbq(A, Sql) As ADODB.Recordset
'QLib.Std.MDao_Fb_Fbq.Sub RunFbq(A, Sql)
'QLib.Std.MDao_Idx.Function CvIdx(A) As Dao.Index
'QLib.Std.MDao_Idx.Function FnyzIdx(A As Dao.Index) As String()
'QLib.Std.MDao_Idx.Function IsEqIdx(A As Dao.Index, B As Dao.Index) As Boolean
'QLib.Std.MDao_Idx.Function IdxIsSk(A As Dao.Index, T) As Boolean
'QLib.Std.MDao_Idx.Function IsEqIdxs(A As Dao.Indexes, B As Dao.Indexes) As Boolean
'QLib.Std.MDao_Idx.Function IdxIsUniq(A As Dao.Index) As Boolean
'QLib.Std.MDao_Lg.Sub CurLgLis(Optional Sep$ = " ", Optional Top% = 50)
'QLib.Std.MDao_Lg.Function CurLgLy(Optional Sep$ = " ", Optional Top% = 50) As String()
'QLib.Std.MDao_Lg.Private Function RsLy(A As Dao.Database, Sep$) As String()
'QLib.Std.MDao_Lg.Function CurLgRs(Optional Top% = 50) As Dao.Recordset
'QLib.Std.MDao_Lg.Sub CurSessLis(Optional Sep$ = " ", Optional Top% = 50)
'QLib.Std.MDao_Lg.Function CurSessLy(Optional Sep$, Optional Top% = 50) As String()
'QLib.Std.MDao_Lg.Function CurSessRs(Optional Top% = 50) As Dao.Recordset
'QLib.Std.MDao_Lg.Private Function CvSess&(A&)
'QLib.Std.MDao_Lg.Private Sub EnsMsg(Fun$, MsgTxt$)
'QLib.Std.MDao_Lg.Private Sub Ens_Sess()
'QLib.Std.MDao_Lg.Private Property Get L() As Database
'QLib.Std.MDao_Lg.Sub Lg(Fun$, MsgTxt$, ParamArray Ap())
'QLib.Std.MDao_Lg.Private Sub AsgRs(A As Dao.Recordset, ParamArray OAp())
'QLib.Std.MDao_Lg.Sub LgAsg(A&, OSess&, OTimStr_Dte$, OFun$, OMsgTxt$)
'QLib.Std.MDao_Lg.Sub LgBeg()
'QLib.Std.MDao_Lg.Sub LgBrw()
'QLib.Std.MDao_Lg.Property Get LgFt$()
'QLib.Std.MDao_Lg.Sub LgCls()
'QLib.Std.MDao_Lg.Sub LgCrt()
'QLib.Std.MDao_Lg.Sub LgCrt_v1()
'QLib.Std.MDao_Lg.Property Get LgDb() As Database
'QLib.Std.MDao_Lg.Sub LgEnd()
'QLib.Std.MDao_Lg.Property Get LgFb$()
'QLib.Std.MDao_Lg.Property Get LgFn$()
'QLib.Std.MDao_Lg.Private Sub X(A$)
'QLib.Std.MDao_Lg.Property Get LgSchm() As String()
'QLib.Std.MDao_Lg.Sub LgKill()
'QLib.Std.MDao_Lg.Function LgLinesAy(A&) As Variant()
'QLib.Std.MDao_Lg.Sub LgLis(Optional Sep$ = " ", Optional Top% = 50)
'QLib.Std.MDao_Lg.Function LgLy(A&) As String()
'QLib.Std.MDao_Lg.Private Sub LgOpn()
'QLib.Std.MDao_Lg.Property Get LgPth$()
'QLib.Std.MDao_Lg.Sub FTIxDmp(A As FTIx)
'QLib.Std.MDao_Lg.Sub SessBrw(Optional A&)
'QLib.Std.MDao_Lg.Function SessLgAy(A&) As Long()
'QLib.Std.MDao_Lg.Sub SessLis(Optional Sep$ = " ", Optional Top% = 50)
'QLib.Std.MDao_Lg.Function SessLy(Optional A&) As String()
'QLib.Std.MDao_Lg.Function SessNLg%(A&)
'QLib.Std.MDao_Lg.Private Sub WrtLg(Fun$, MsgTxt$)
'QLib.Std.MDao_Lg.Private Sub Z_Lg()
'QLib.Std.MDao_Lg.Private Sub Z()
'QLib.Std.MDao_Li.Private Sub Z_LnkImp()
'QLib.Std.MDao_Li.Function LnkImp(A As LiPm) As Database
'QLib.Std.MDao_Li.Private Sub Z_ChkColzLiPm()
'QLib.Std.MDao_Li.Private Function ChkColzLiPm(A As LiPm) As String()
'QLib.Std.MDao_Li.Private Function ImpSqyzLi(A As LiPm) As String()
'QLib.Std.MDao_Li.Private Function ImpSqyFb(A() As LiFb) As String()
'QLib.Std.MDao_Li.Private Function ImpSqyFx(A() As LiFx) As String()
'QLib.Std.MDao_Li.Private Function ImpSqlFx$(A As LiFx)
'QLib.Std.MDao_Li.Private Function ImpSqlFb$(A As LiFb)
'QLib.Std.MDao_Lid.Sub LnkImpzLidPm(A As LidPm)
'QLib.Std.MDao_Lid.Private Function ErzLidPm(A As LidPm) As String()
'QLib.Std.MDao_Lid.Private Function ImpSqyzLidPm(A As LidPm) As String()
'QLib.Std.MDao_Lid.Private Function ImpSqyzFbAy(A() As LidFb) As String()
'QLib.Std.MDao_Lid.Private Function ImpSqyzFxAy(A() As LidFx) As String()
'QLib.Std.MDao_Lid.Private Function ImpSqyzFx(A As LidFx) As String()
'QLib.Std.MDao_Lid.Private Function ExTnyzLidFxcAy(A() As LidFxc) As String()
'QLib.Std.MDao_Lid.Private Function FnyzLidFxcAy(A() As LidFxc) As String()
'QLib.Std.MDao_Lid.Private Function ImpSqyzFb(A As LiFb) As String()
'QLib.Std.MDao_Lid.Private Sub Z_ImpSqy()
'QLib.Std.MDao_Lid.Private Sub Z_ErzLidPm()
'QLib.Std.MDao_Lid.Private Sub Z()
'QLib.Std.MDao_Li_Act.Function LiAct(A As LiPm) As LiAct
'QLib.Std.MDao_Li_Act.Private Function LiActFxOpt(A As LiFx, Fx$) As LiActFx
'QLib.Std.MDao_Li_Act.Private Function LiActFxAy(A() As LiFx, ExistFilNmToFfnDic As Dictionary) As LiActFx()
'QLib.Std.MDao_Li_Act.Private Function LiActFbAy(A() As LiFb, ExistFilNmToFfnDic As Dictionary) As LiActFb()
'QLib.Std.MDao_Li_Act.Private Function LiActFbOpt(A As LiFb, Fb$) As LiActFb
'QLib.Std.MDao_Li_Act_Brw.Sub BrwLiAct(A As LiAct)
'QLib.Std.MDao_Li_Act_Brw.Function LiActDrs(A As LiAct) As Drs
'QLib.Std.MDao_Li_Act_Brw.Private Function LiActDry(A As LiAct) As Variant()
'QLib.Std.MDao_Li_Act_Brw.Private Function LiActDryFxAy(A() As LiActFx) As Variant()
'QLib.Std.MDao_Li_Act_Brw.Private Function LiActDryFbAy(A() As LiActFb) As Variant()
'QLib.Std.MDao_Li_Act_Brw.Private Function LiActDrFb(A As LiActFb) As Variant()
'QLib.Std.MDao_Li_Act_Brw.Private Function LiActDryFx(A As LiActFx) As Variant()
'QLib.Std.MDao_Li_Act_Brw.Private Sub Z_LiAct()
'QLib.Std.MDao_Lid_Brw.Private Sub Z_BrwLidPm()
'QLib.Std.MDao_Lid_Brw.Sub BrwLidPm(A As LidPm)
'QLib.Std.MDao_Lid_Brw.Function DszLidPm(A As LidPm) As Ds
'QLib.Std.MDao_Lid_Brw.Private Function FbColDt(A() As LidFb) As Dt
'QLib.Std.MDao_Lid_Brw.Private Function FilDt(A() As LidFil) As Dt
'QLib.Std.MDao_Lid_Brw.Private Function FilDry(A() As LidFil) As Variant()
'QLib.Std.MDao_Lid_Brw.Private Function FxColDt(A() As LidFx) As Dt
'QLib.Std.MDao_Lid_Brw.Private Function FxColDry(A() As LidFx) As Variant()
'QLib.Std.MDao_Lid_Brw.Private Function FxColDr(A As LidFx) As Variant()
'QLib.Std.MDao_Lid_Mis_MsgzLIdMis.Function MsgzLidMis(A As LidMis) As String()
'QLib.Std.MDao_Lid_Mis_MsgzLIdMis.Private Function Tbl(A() As LidMisTbl) As String()
'QLib.Std.MDao_Lid_Mis_MsgzLIdMis.Private Function Col(A() As LidMisCol) As String()
'QLib.Std.MDao_Lid_Mis_MsgzLIdMis.Private Function Ty(A() As LidMisTy) As String()
'QLib.Std.MDao_Lid_Mis_MsgzLIdMis.Private Sub Z_MsgzLidMis()
'QLib.Std.MDao_Li_Pm.Sub BrwSampLiPm()
'QLib.Std.MDao_Li_Pm.Property Get SampLiPm() As LiPm
'QLib.Std.MDao_Li_Pm.Function LiPm(Src$()) As LiPm
'QLib.Std.MDao_Li_Pm.Private Property Get Apn$()
'QLib.Std.MDao_Li_Pm.Private Property Get Fil() As LiFil()
'QLib.Std.MDao_Li_Pm.Private Property Get Fb() As LiFb()
'QLib.Std.MDao_Li_Pm.Private Property Get Fx() As LiFx()
'QLib.Std.MDao_Li_Pm.Private Function FxcAy(T$) As LiFxc()
'QLib.Std.MDao_Li_Pm.Function SampLiPmSrc() As String()
'QLib.Std.MDao_Li_Pm.Private Function LtPmAyFx(A() As LiFx, FfnDic As Dictionary) As LtPm()
'QLib.Std.MDao_Li_Pm.Private Function LtPmAyFb(A() As LiFb, FfnDic As Dictionary) As LtPm()
'QLib.Std.MDao_Li_Pm.Function LtPm(A As LiPm) As LtPm()
'QLib.Std.MDao_Lid_Lnk.Function ErzLnkTblzLtPm(Db As Database, A() As LtPm) As String()
'QLib.Std.MDao_Lid_Lnk.Sub LnkTblzLtPm(Db As Database, A() As LtPm)
'QLib.Std.MDao_Lid_Lnk.Function TdzTSCn(T, Src, Cn) As Dao.TableDef
'QLib.Std.MDao_Lid_Lnk.Sub LnkTblzTSCn(Db As Database, T, S, Cn)
'QLib.Std.MDao_Lid_Lnk.Sub LnkFxw(A As Database, T, Fx, Wsn)
'QLib.Std.MDao_Lid_Lnk.Sub LnkFbzTT(Db As Database, TTCrt$, Fb$, Optional Fbtt$)
'QLib.Std.MDao_Lid_Lnk.Function LnkTnyDb(Db As Database) As String()
'QLib.Std.MDao_Lid_Lnk.Sub LnkFb(A As Database, T, Fb$, Optional Fbt)
'QLib.Std.MDao_Lid_LnkChk.Function ErzLnkTblzTSrcCn(A As Database, T, S$, Cn$) As String()
'QLib.Std.MDao_Lid_LnkChk.Function ChkFxww(Fx$, Wsnss$, Optional FxKind$ = "Excel file") As String()
'QLib.Std.MDao_Lid_LnkChk.Function ChkWs(Fx, Wsn, FxKind$) As String()
'QLib.Std.MDao_Lid_LnkChk.Function ChkFxw(Fx, Wsn, Optional FxKind$ = "Excel file") As String()
'QLib.Std.MDao_Lid_LnkChk.Function ChkLnkWs(A As Database, T, Fx, Wsn, Optional FxKind$ = "Excel file") As String()
'QLib.Std.MDao_Lid_LtPm.Function LtPmzLid(A As LidPm) As LtPm()
'QLib.Std.MDao_Lid_LtPm.Private Function LtPmAyFx(Pm As LidPm, FfnDic As Dictionary, WPth$) As LtPm()
'QLib.Std.MDao_Lid_LtPm.Private Function LtPmAyFb(Pm As LidPm, FfnDic As Dictionary, WPth$) As LtPm()
'QLib.Std.MDao_Lid_LtPm.Private Function FilNmToFfnDiczLidPm(A() As LidFil) As Dictionary
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Function LidMis(A As LidPm) As LidMis
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function MisTbl(T() As Tbl) As LidMisTbl()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function MisCol(T() As Tbl) As LidMisCol()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function MisTy(T() As Tbl) As LidMisTy()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function T1Ay(A As LidPm) As Tbl()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function FfnDic(A() As LidFil) As Dictionary
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function T1Fx(A As LidFx, B() As LidFil, FfnDic As Dictionary) As Tbl
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function FsetzFxc(A() As LidFxc) As Aset
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function FldNmToEptShtTyLisDiczFxc(A() As LidFxc) As Dictionary
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function T1Fb(A As LidFb, B() As LidFil) As Tbl
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function T2Ay(T() As Tbl, ExistFfnAy$()) As Tbl()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function T3Ay(T() As Tbl) As Tbl()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function T4Ay(T() As Tbl) As Tbl()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function T5Ay(T() As Tbl) As Tbl()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function Tyc(ExistFset As Aset, FldNmToEptShtTyLisDic As Dictionary, Ffn$, TblNm$) As LidMisTyc()
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function Tyci(ActShtTy$, EptShtTyLis$, ExtNm) As TycOpt
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function FfnzFilNm$(FilNm$, A() As LidFil)
'QLib.Std.MDao_Lid_Mis_Fm_LidPm.Private Function FfnAyzLidFil(A() As LidFil) As String()
'QLib.Std.MDao_Li_MisCol.Function MisCol(A As LiPm, B As LiAct) As LiMisCol()
'QLib.Std.MDao_Li_MisCol.Private Function MisColFx(A() As LiFx, B() As LiActFx) As LiMisCol()
'QLib.Std.MDao_Li_MisCol.Private Function MisColFb(A() As LiFb, B() As LiActFb) As LiMisCol()
'QLib.Std.MDao_Li_MisCol.Private Function MisColFxOpt(A As LiFx, Fx$, ActFset As Aset) As LiMisCol
'QLib.Std.MDao_Li_MisCol.Private Function MisColFbOpt(A As LiFb, Fb$, ActFset As Aset) As LiMisCol
'QLib.Std.MDao_Li_MisCol.Private Sub AsgActFsetFb(OActFsetOpt As Aset, OFb$, T$, B() As LiActFb)
'QLib.Std.MDao_Li_MisCol.Private Sub AsgActFsetFx(OActFsetOpt As Aset, OFx$, Wsn$, B() As LiActFx)
'QLib.Std.MDao_Li_MisTy.Function MisTyFxAy(A() As LiFx, B() As LiActFx) As LiMisTy()
'QLib.Std.MDao_Li_MisTy.Private Function MisTyFxOpt(A As LiFx, B() As LiActFx) As LiMisTy
'QLib.Std.MDao_Li_MisTy.Private Function MisColAy(Ept() As LiFxc, ActShtTyDic As Dictionary) As LiMisTyc()
'QLib.Std.MDao_Li_MisTy.Private Function MisColOpt(Ept As LiFxc, ActShtTyDic As Dictionary) As LiMisTyc
'QLib.Std.MDao_Li_MisTy.Private Function IsMisTy(EptShtTyLis$, ActShtTy$) As Boolean
'QLib.Std.MDao_Li_MisTy.Private Function LiActFxOpt(Fxn$, Wsn$, A() As LiActFx) As LiActFx
'QLib.Std.MDao_Lid_Pm.Function LidPm(Src$()) As LidPm
'QLib.Std.MDao_Lid_Pm.Private Property Get Apn$()
'QLib.Std.MDao_Lid_Pm.Private Function LyT1(T1$) As String()
'QLib.Std.MDao_Lid_Pm.Private Property Get Fx() As LidFx()
'QLib.Std.MDao_Lid_Pm.Private Function Fxi(WsLin) As LidFx
'QLib.Std.MDao_Lid_Pm.Private Function FxcAy(T$) As LidFxc()
'QLib.Std.MDao_Lid_Pm.Private Function Fxc(WsColLin) As LidFxc
'QLib.Std.MDao_Lid_Pm.Private Function LyzWsCol(TblNm$)
'QLib.Std.MDao_Lid_Pm.Private Function FfnDic() As Dictionary
'QLib.Std.MDao_Lid_Pm.Private Property Get Fb() As LidFb()
'QLib.Std.MDao_Lid_Pm.Private Function Fbi(TblLin, FfnDic As Dictionary) As LidFb
'QLib.Std.MDao_Lid_Pm.Private Property Get Fil() As LidFil()
'QLib.Std.MDao_Lid_Pm.Private Function Fili(L) As LidFil
'QLib.Std.MDao_Lid_PmFmt.Function FmtLidPm(A As LidPm) As String()
'QLib.Std.MDao_Lid_PmFmt.Private Property Get ApnLin$()
'QLib.Std.MDao_Lid_PmFmt.Private Function Fil() As LidFil()
'QLib.Std.MDao_Lid_PmFmt.Private Property Get Fx() As String()
'QLib.Std.MDao_Lid_PmFmt.Private Function FxLin$(A As LidFx)
'QLib.Std.MDao_Lid_PmFmt.Private Property Get Fb() As String()
'QLib.Std.MDao_Lid__Is_LnkImpDef.Function LiFxc(ColNm$, ShtTyLis$, ExtNm$) As LiFxc
'QLib.Std.MDao_Lid__Is_LnkImpDef.Function LiFxcLnkColStr(LnkColStr) As LiFxc
'QLib.Std.MDao_Lid__Is_LnkImpDef.Function LiFxcAy(LnkColVbl$) As LiFxc()
'QLib.Std.MDao_Lid__Is_LnkImpDef.Function LiFxcAyLnkColAy(A$()) As LiFxc()
'QLib.Std.MDao_Lid__Is_LnkImpDef.Function LiFx(Fxn$, T$, Wsn$, Fxc() As LiFxc, Optional Bexpr) As LiFx
'QLib.Std.MDao_Lid__Is_LnkImpDef.Function LiFxLnkColVbl(Fxn$, T$, Wsn$, LnkColVbl$, Optional Bexpr$) As LiFx
'QLib.Std.MDao_Lid__Is_LnkImpDef.Function LiFil(FilNm$, Ffn$) As LiFil
'QLib.Std.MDao_Lid__Is_LnkImpDef.Function LiFb(Fbn, T, Fset As Aset, Bexpr$) As LiFb
'QLib.Std.MDao_Prp.Property Get FldDes$(A As Database, T, F)
'QLib.Std.MDao_Prp.Property Let FldDes(A As Database, T, F, Des$)
'QLib.Std.MDao_Prp.Private Sub Z_PrpNy()
'QLib.Std.MDao_Prp.Function PrpNyFd(A As Dao.Field) As String()
'QLib.Std.MDao_Prp.Private Sub Z_PrpDryzFd()
'QLib.Std.MDao_Prp.Function PrpDryzFd(A As Dao.Field) As Variant()
'QLib.Std.MDao_Prp.Property Let PrpVal(O, P, V)
'QLib.Std.MDao_Prp.Property Get PrpVal(O, P)
'QLib.Std.MDao_Prp.Private Sub Z_FldPrp()
'QLib.Std.MDao_Prp.Property Get FldDeszTd$(A As Dao.Field)
'QLib.Std.MDao_Prp.Property Let FldDeszTd(A As Dao.Field, Des$)
'QLib.Std.MDao_Prp.Private Sub Z()
'QLib.Std.MDao_Prp.Private Sub ZZ_TblPrp()
'QLib.Std.MDao_Prp.Function HasDbtPrp(A As Database, T, P) As Boolean
'QLib.Std.MDao_Prp.Property Get TblPrp(A As Database, T, P)
'QLib.Std.MDao_Prp.Property Let TblPrp(A As Database, T, P, V)
'QLib.Std.MDao_Prp.Sub SetDaoPrp(DaoObj, Prps As Dao.Properties, P, V)
'QLib.Std.MDao_Prp.Property Let FldPrp(A As Database, T, F, P, V)
'QLib.Std.MDao_Prp.Property Get FldPrp(A As Database, T, F, P)
'QLib.Std.MDao_Prp.Function HasFldPrp(A As Database, T, F, P) As Boolean
'QLib.Std.MDao_Rs.Sub UpdRs(Rs As Dao.Recordset, Dr)
'QLib.Std.MDao_Rs.Private Sub ZZ_AsgRs()
'QLib.Std.MDao_Rs.Function CvRs(A) As Dao.Recordset
'QLib.Std.MDao_Rs.Function NoRec(A As Dao.Recordset) As Boolean
'QLib.Std.MDao_Rs.Function HasRec(A As Dao.Recordset) As Boolean
'QLib.Std.MDao_Rs.Sub AsgRs(A As Dao.Recordset, ParamArray OAp())
'QLib.Std.MDao_Rs.Sub BrwRs(A As Dao.Recordset)
'QLib.Std.MDao_Rs.Sub BrwSngRec(A As Dao.Recordset)
'QLib.Std.MDao_Rs.Sub RsDlt(A As Dao.Recordset)
'QLib.Std.MDao_Rs.Function CsvLinzRs$(A As Dao.Recordset)
'QLib.Std.MDao_Rs.Function CsvLyzRs1(A As Dao.Recordset) As String()
'QLib.Std.MDao_Rs.Function CsvLyzRs(A As Dao.Recordset, Optional FF) As String()
'QLib.Std.MDao_Rs.Function AsetzRs(Rs As Dao.Recordset, Optional Fld = 0) As Aset
'QLib.Std.MDao_Rs.Function RsMovFst(Rs As Dao.Recordset) As Dao.Recordset
'QLib.Std.MDao_Rs.Sub DmpRs(A As Recordset, Optional FF)
'QLib.Std.MDao_Rs.Function DrzRs(A As Dao.Recordset, Optional FF = "") As Variant()
'QLib.Std.MDao_Rs.Function DrszRs(A As Dao.Recordset) As Drs
'QLib.Std.MDao_Rs.Function DryzRs(A As Dao.Recordset, Optional ExlFldNm As Boolean) As Variant()
'QLib.Std.MDao_Rs.Function FnyzRs(A As Dao.Recordset) As String()
'QLib.Std.MDao_Rs.Function HasReczFEv(Rs As Dao.Recordset, F, Ev) As Boolean
'QLib.Std.MDao_Rs.Function IntAyzRs(A As Dao.Recordset, Optional Fld = 0) As Integer()
'QLib.Std.MDao_Rs.Function ShouldBrkvRs(A As Dao.Recordset, GpKy$(), LasVy()) As Boolean
'QLib.Std.MDao_Rs.Function RsLin$(A As Dao.Recordset, Optional Sep$ = " ")
'QLib.Std.MDao_Rs.Function LngAyzRs(A As Dao.Recordset, Optional Fld = 0) As Long()
'QLib.Std.MDao_Rs.Function RsLy(A As Dao.Recordset, Optional Sep$ = " ") As String()
'QLib.Std.MDao_Rs.Function FmtRec(A As Dao.Recordset)
'QLib.Std.MDao_Rs.Function NReczRs&(A As Dao.Recordset)
'QLib.Std.MDao_Rs.Sub SetSqrzRs(OSq, R, A As Dao.Recordset, Optional NoTxtSngQ As Boolean)
'QLib.Std.MDao_Rs.Function SyzRs(A As Dao.Recordset, Optional F = 0) As String()
'QLib.Std.MDao_Rs.Function StruzRs$(A As Dao.Recordset)
'QLib.Std.MDao_Rs.Function NzEmpty(A)
'QLib.Std.MDao_Rs.Function DrzRsFny(Fny$(), Rs As Dao.Recordset) As Variant()
'QLib.Std.MDao_Rs.Function IntozRs(Into, Rs As Recordset, Optional Fld = 0)
'QLib.Std.MDao_Rs.Function AvRsCol(A As Dao.Recordset, Optional Fld = 0) As Variant()
'QLib.Std.MDao_Rs.Function ColSetzRs(A As Dao.Recordset, Optional Fld = 0) As Aset
'QLib.Std.MDao_Rs.Function SqzRs(A As Dao.Recordset, Optional ExlFldNm As Boolean) As Variant()
'QLib.Std.MDao_Rs_Mdy.Sub InsRszDry(A As Dao.Recordset, Dry())
'QLib.Std.MDao_Rs_Mdy.Sub SetRs(Rs As Dao.Recordset, Dr)
'QLib.Std.MDao_Rs_Mdy.Sub InsRszAp(Rs As Dao.Recordset, ParamArray Ap())
'QLib.Std.MDao_Rs_Mdy.Sub InsRs(Rs As Dao.Recordset, Dr)
'QLib.Std.MDao_Rs_Mdy.Sub UpdRszAp(Rs As Dao.Recordset, ParamArray Ap())
'QLib.Std.MDao_Rs_Mdy.Sub DltRs(A As Dao.Recordset)
'QLib.Std.MDao_Schm.Sub CrtSchmzVbl(A As Database, SchmVbl$)
'QLib.Std.MDao_Schm.Sub CrtSchm(A As Database, Schm$())
'QLib.Std.MDao_Schm.Private Function EFzSchm(Schm$()) As EF
'QLib.Std.MDao_Schm.Private Function PkTny(TdLy$()) As String()
'QLib.Std.MDao_Schm.Private Function SqyCrtSk(TdLy$()) As String()
'QLib.Std.MDao_Schm.Private Function SkFny(TdLin) As String()
'QLib.Std.MDao_Schm.Private Function TdAy(TdLy$(), A As EF) As Dao.TableDef()
'QLib.Std.MDao_Schm.Private Function TdzLin(TdLin, A As EF) As Dao.TableDef
'QLib.Std.MDao_Schm.Function T1zItm_T1LikssAy$(Itm, T1LikssAy$())
'QLib.Std.MDao_Schm.Private Function FdzEF(F, A As EF) As Dao.Field2
'QLib.Std.MDao_Schm.Private Function FdzEle(Ele$, EleLy$(), F) As Dao.Field2
'QLib.Std.MDao_Schm.Private Function EleStr$(EleLy$(), Ele$)
'QLib.Std.MDao_Schm.Private Function EleStrzStd$(Ele)
'QLib.Std.MDao_Schm.Private Property Get Schm1() As String()
'QLib.Std.MDao_Schm.Private Sub Z_CrtSchm()
'QLib.Std.MDao_Schm.Sub AA()
'QLib.Std.MDao_Schm.Private Sub Z()
'QLib.Std.MDao_Schm.Sub AppTdAy(A As Database, TdAy() As Dao.TableDef)
'QLib.Std.MDao_Schm.Function FnyzTdLin(TdLin) As String()
'QLib.Std.MDao_Schm_Ens.Sub EnsSchm(A As Database, Schm$())
'QLib.Std.MDao_Schm_Er.Function ErzSchm(Schm$()) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErD_LinEr(A() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErDT_InvalidFld(D() As Lnx, T$, Fny$()) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErDT_InvalidFld1$(D As Lnx, T$, Fny$())
'QLib.Std.MDao_Schm_Er.Private Function ErDF_Er(A() As Lnx, T() As Lnx) As String() 'Given A is D-Lnx having fmt = $Tbl $Fld $D, 'This Sub checks if $Fld is in Fny
'QLib.Std.MDao_Schm_Er.Private Function ErDF_Er1(A As Lnx, Tny$(), FnyAy$()) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErDT_Tbl_NotIn_Tny1$(D As Lnx, Tny$())
'QLib.Std.MDao_Schm_Er.Private Function ErDT_Tbl_NotIn_Tny(D() As Lnx, Tny$()) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErE_DupE(E() As Lnx, Eny$()) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErE_ELnx(A As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErE_ELnxAy(A() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErT_FldHasNoEle(T() As Lnx, E() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErF_Ele_NotIn_Eny(F() As Lnx, Eny$()) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErF_EleHasNoDef(F() As Lnx, AllEny$()) As String()
'QLib.Std.MDao_Schm_Er.Private Function Er_1_OneLiner(F As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErF_1_LinEr(A() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErT_DupTbl(T() As Lnx, Tny$()) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErT_NoTLin(A() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErT_LinEr_zLnx(T As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErT_LinEr(A() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function LnoAyzEle(E() As Lnx, Ele) As Long()
'QLib.Std.MDao_Schm_Er.Private Function LnoAyzTbl(A() As Lnx, T) As Long()
'QLib.Std.MDao_Schm_Er.Private Function MsgMultiLno(LnoAy&(), M$)
'QLib.Std.MDao_Schm_Er.Private Function Msg$(A As Lnx, M$)
'QLib.Std.MDao_Schm_Er.Private Function MsgDF_InvalidFld$(ErLin As Lnx, ErFld$, VdtFny$())
'QLib.Std.MDao_Schm_Er.Private Function MsgD_NTermShouldBe3OrMore$(D As Lnx)
'QLib.Std.MDao_Schm_Er.Private Function MsgDT_Tbl_NotIn_Tny$(A As Lnx, T, Tblss$)
'QLib.Std.MDao_Schm_Er.Private Function MsgE_DupE$(LnoAy&(), E)
'QLib.Std.MDao_Schm_Er.Private Function MsgT_DupT$(LnoAy&(), T)
'QLib.Std.MDao_Schm_Er.Private Function MsgE_ExcessEleItm$(A As Lnx, ExcessEle$)
'QLib.Std.MDao_Schm_Er.Private Function MsgF_ExcessTxtSz$(A As Lnx)
'QLib.Std.MDao_Schm_Er.Private Function MsgF_Ele_NotIn_Eny$(A As Lnx, E$, Eless$)
'QLib.Std.MDao_Schm_Er.Private Function MsgE_FldEleEr$(A As Lnx, E$, Eless$)
'QLib.Std.MDao_Schm_Er.Private Function MsgFzDLy_NotIn_Fny$(A As Lnx, T$, F$, Fssl$)
'QLib.Std.MDao_Schm_Er.Private Function MsgT_DupF$(A As Lnx, T$, Fny$())
'QLib.Std.MDao_Schm_Er.Private Function MsgT_FldIsNotANmEr$(A As Lnx, F)
'QLib.Std.MDao_Schm_Er.Private Function MsgT_IdFld$(A As Lnx)
'QLib.Std.MDao_Schm_Er.Private Function MsgT_NoFld(A As Lnx)
'QLib.Std.MDao_Schm_Er.Private Property Get MsgT_NoTLin$()
'QLib.Std.MDao_Schm_Er.Private Function MsgT_TblIsNotNm$(A As Lnx)
'QLib.Std.MDao_Schm_Er.Private Function MsgT_Vbar_Cnt$(A As Lnx)
'QLib.Std.MDao_Schm_Er.Private Function MsgT_FldEr$(A As Lnx, F$)
'QLib.Std.MDao_Schm_Er.Private Function Msg_LinTyEr$(A As Lnx, Ty$)
'QLib.Std.MDao_Schm_Er.Private Function ErD(Tny$(), T() As Lnx, D() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErE(E() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErF(AllEny$(), F() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Function ErT(Tny$(), T() As Lnx, E() As Lnx) As String()
'QLib.Std.MDao_Schm_Er.Private Sub Z_ErT_LinEr_zLnx()
'QLib.Std.MDao_Schm_Er.Private Sub Z()
'QLib.Std.MDao_Schm_Std.Function StdSchmFldLy() As String()
'QLib.Std.MDao_Schm_Std.Function StdSchmEleLy() As String()
'QLib.Std.MDao_Spec.Sub ImpSpec(A As Database, Spnm)
'QLib.Std.MDao_Spec.Function SpecPth$(Apn$)
'QLib.Std.MDao_Spec.Sub BrwSpecPth(Apn$)
'QLib.Std.MDao_Spec.Sub ClrSpecPth(Apn$)
'QLib.Std.MDao_Spec.Sub EnsTblSpec(A As Database)
'QLib.Std.MDao_Spec.Sub CrtTblSpec(A As Database)
'QLib.Std.MDao_Spec.Sub ExpSpec(Apn$)
'QLib.Std.MDao_Spec.Function SpecNy(Apn$) As String()
'QLib.Std.MDao_Sql.Private Property Get C_Fm$()
'QLib.Std.MDao_Sql.Private Property Get C_Into$()
'QLib.Std.MDao_Sql.Private Property Get C_NL$() ' New Line
'QLib.Std.MDao_Sql.Private Property Get C_T$()
'QLib.Std.MDao_Sql.Private Property Get C_NLT$() ' New Line Tabe
'QLib.Std.MDao_Sql.Private Property Get C_NLTT$() ' New Line Tabe
'QLib.Std.MDao_Sql.Private Function AyQ(Ay) As String()
'QLib.Std.MDao_Sql.Function FldInVy_Str$(F, InAy)
'QLib.Std.MDao_Sql.Function FFJnComma$(FF)
'QLib.Std.MDao_Sql.Function SqpInto_T$(T)
'QLib.Std.MDao_Sql.Function BexprRecId$(T, RecId)
'QLib.Std.MDao_Sql.Function SqpSet_Fny_Vy$(Fny$(), Vy())
'QLib.Std.MDao_Sql.Function SqpAnd_Bexpr$(Bexpr$)
'QLib.Std.MDao_Sql.Private Function AyAddPfxNLTT$(Ay)
'QLib.Std.MDao_Sql.Private Function ExprInLis_InLisBexpr$(Expr$, InLis$)
'QLib.Std.MDao_Sql.Function SqpSel_F$(F)
'QLib.Std.MDao_Sql.Function SqpSel_X$(X, Optional Dis As Boolean)
'QLib.Std.MDao_Sql.Function SqpFm$(T)
'QLib.Std.MDao_Sql.Function SqpGp_ExprVblAy$(ExprVblAy$())
'QLib.Std.MDao_Sql.Private Sub Z_SqpGp_ExprVblAy()
'QLib.Std.MDao_Sql.Function WdtLines%(Lines)
'QLib.Std.MDao_Sql.Function WdtzLinesAy%(LinesAy)
'QLib.Std.MDao_Sql.Function LinesFmtAyL(LinesAy$()) As String()
'QLib.Std.MDao_Sql.Function SqpSelX_FF_ExtNy$(FF, ExtNy$())
'QLib.Std.MDao_Sql.Function SqpSel_FF_Ey$(FF, ExprAy$())
'QLib.Std.MDao_Sql.Function JnCommaSpcFF$(FF)
'QLib.Std.MDao_Sql.Function SqpSel_FF$(FF, Optional Dis As Boolean)
'QLib.Std.MDao_Sql.Function SqpSel_Dis$(Dis As Boolean)
'QLib.Std.MDao_Sql.Private Sub Z_SqpSel()
'QLib.Std.MDao_Sql.Function SqlSel_FF$(FF, Optional IsDis As Boolean)
'QLib.Std.MDao_Sql.Function SqpSet_FF_Ey$(FF, Ey$())
'QLib.Std.MDao_Sql.Private Sub Z_SqpSetFFEqvy()
'QLib.Std.MDao_Sql.Function SqpSet_FF_Evy$(FF, EqVy)
'QLib.Std.MDao_Sql.Private Function QNm$(T)
'QLib.Std.MDao_Sql.Function SqpUpd_T$(T)
'QLib.Std.MDao_Sql.Function SqpWhfv(F, V) ' Ssk is single-Sk-value
'QLib.Std.MDao_Sql.Function SqpWhK$(K&, T)
'QLib.Std.MDao_Sql.Function SqpWhBet_F_Fm_To$(F, FmV, ToV)
'QLib.Std.MDao_Sql.Private Function QV$(V)
'QLib.Std.MDao_Sql.Private Property Get C_And$()
'QLib.Std.MDao_Sql.Private Property Get C_Wh$()
'QLib.Std.MDao_Sql.Function SqpWh_F_InVy$(F, InVy)
'QLib.Std.MDao_Sql.Private Sub Z_SqpWhFldInVy_Str()
'QLib.Std.MDao_Sql.Private Function FnyEqVy_Bexpr$(Fny$(), EqVy)
'QLib.Std.MDao_Sql.Function SqpWh_FnyEqVy$(Fny$(), EqVy)
'QLib.Std.MDao_Sql.Function SqpWh$(A$)
'QLib.Std.MDao_Sql.Private Sub Z_SqpSet_Fny_VyFmt()
'QLib.Std.MDao_Sql.Private Sub Z_SqpWhFldInVy_StrSqpAy()
'QLib.Std.MDao_Sql.Function VblFmtAyAsLines$(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAy, Optional Sep$ = ",")
'QLib.Std.MDao_Sql.Function VblFmtAyAsLy(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAyOpt, Optional Sep$ = ",") As String()
'QLib.Std.MDao_Sql.Function SqlSel_FF_EDic_Fm$(FF, EDic As Dictionary, T, Optional IsDis As Boolean)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Fm$(FF, T, Optional IsDis As Boolean, Optional Bexpr$)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Ey_Fm$(FF, Ey$(), T, Optional IsDis As Boolean, Optional Bexpr$)
'QLib.Std.MDao_Sql.Function ItrTT(TT)
'QLib.Std.MDao_Sql.Function FnyzPfxN(Pfx$, N%) As String()
'QLib.Std.MDao_Sql.Function NsetzNN(FF) As Aset
'QLib.Std.MDao_Sql.Function NyzNNDft(NN, DftFny$()) As String()
'QLib.Std.MDao_Sql.Function NyzNN(NN) As String()
'QLib.Std.MDao_Sql.Function QuoteSql$(V)
'QLib.Std.MDao_Sql.Function AyQuoteSql(Ay) As String()
'QLib.Std.MDao_Sql.Function SqlUpd_T_FF_EqDr_Whff_Eqvy$(T$, FF, Dr, WhFF, EqVy)
'QLib.Std.MDao_Sql.Function SqpWh_FF_Eqvy$(FF, EqVy)
'QLib.Std.MDao_Sql.Function SqpSet_FF_EqDr$(FF, EqDr)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Fm_Bexpr$(FF, T, Bexpr$)
'QLib.Std.MDao_Sql.Function QAddCol$(T, Fny0, F As Drs, E As Dictionary)
'QLib.Std.MDao_Sql.Function SqlCrtPkzT$(T)
'QLib.Std.MDao_Sql.Function SqlCrtSk_T_SkFF$(T, SkFF)
'QLib.Std.MDao_Sql.Function SqlCrtTbl_T_X$(T, X$)
'QLib.Std.MDao_Sql.Function SqlDrpCol_T_F$(T, F)
'QLib.Std.MDao_Sql.Function SqlDrpTbl_T$(T)
'QLib.Std.MDao_Sql.Function SqlIns_T_FF_Dr$(T, FF, Dr)
'QLib.Std.MDao_Sql.Function SqlSel_T$(T)
'QLib.Std.MDao_Sql.Function SqlSel_T_Wh$(T, Bexpr$)
'QLib.Std.MDao_Sql.Function SqlSel_Into_Fm_WhFalse(Into, T)
'QLib.Std.MDao_Sql.Function SqlSel_F$(F)
'QLib.Std.MDao_Sql.Function SqlSel_F_Fm$(F, T, Optional Bexpr$)
'QLib.Std.MDao_Sql.Function SqpOrd_FFMinus$(OrdFFMinus)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Fm_Ord(FF, T, OrdFFMinus)
'QLib.Std.MDao_Sql.Function SqlUpd_T_Sk_Fny_Dr$(T, Sk$(), Fny$(), Dr)
'QLib.Std.MDao_Sql.Function SqpSet_Fny_Vy1$(Fny$(), Vy())
'QLib.Std.MDao_Sql.Function FnyAlignQuote(Fny$()) As String()
'QLib.Std.MDao_Sql.Private Sub Z_SqlDtlTWhfInAset()
'QLib.Std.MDao_Sql.Function SqlDlt_Fm$(T)
'QLib.Std.MDao_Sql.Function SqlDlt_Fm_Wh$(T, Bexpr$)
'QLib.Std.MDao_Sql.Function SqyDlt_Fm_WhFld_InAset(T, F, S As Aset, Optional SqlWdt% = 3000) As String()
'QLib.Std.MDao_Sql.Function SqpFldInX_F_InAset_Wdt(F, S As Aset, Wdt%) As String()
'QLib.Std.MDao_Sql.Function LyJnSqlCommaAsetW(A As Aset, W%) As String()
'QLib.Std.MDao_Sql.Function SqpBexpr_F_Ev$(F, Ev)
'QLib.Std.MDao_Sql.Function SqpBktFF$(FF)
'QLib.Std.MDao_Sql.Function JnCommaFF$(FF)
'QLib.Std.MDao_Sql.Function SqlIns_T_FF_Valap$(T, FF, ParamArray ValAp())
'QLib.Std.MDao_Sql.Function SqpIns_T$(T)
'QLib.Std.MDao_Sql.Function SqpBktAv$(Av())
'QLib.Std.MDao_Sql.Function SqlSel_Fny_Fm(Fny$(), Fm, Optional Bexpr$, Optional IsDis As Boolean)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Fm_WhF_InAy$(FF, Fm, WhF, InAy, Optional IsDis As Boolean)
'QLib.Std.MDao_Sql.Function QSelDis_FF_Fm$(FF, T)
'QLib.Std.MDao_Sql.Function SqlSel_FF_ExprDic_Fm$(FF, E As Dictionary, Fm, Optional IsDis As Boolean)
'QLib.Std.MDao_Sql.Function SqlSel_Fm_WhId$(T, Id)
'QLib.Std.MDao_Sql.Function SqpSel_Fm$(T)
'QLib.Std.MDao_Sql.Function SqpWh_T_Id$(T, Id)
'QLib.Std.MDao_Sql.Function QSelDis_FF_ExprDic_Fm$(FF, E As Dictionary, Fm)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Into_Fm$(FF, Into, Fm, Optional Bexpr$, Optional Dis As Boolean)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Fm_WhFny_EqVy$(FF, Fm, Fny$(), EqVy)
'QLib.Std.MDao_Sql.Function SqlSel_FF_ExtNy_Into_Fm$(FF, ExtNy$(), Into, Fm, Optional Bexpr$)
'QLib.Std.MDao_Sql.Function SqlSel_Fm$(Fm, Optional Bexpr$)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Into_Fm_WhFalse$(FF, Into, T)
'QLib.Std.MDao_Sql.Function SqlSel_X_Into_Fm$(X, Into, Fm, Optional Bexpr$)
'QLib.Std.MDao_Sql.Function SqlSel_X_Fm$(X, Fm, Optional Bexpr$)
'QLib.Std.MDao_Sql.Function SqlSel_FF_Fm_OrdFF$(FF, T, OrdFFMinsu)
'QLib.Std.MDao_Sql.Function SqlSelCnt_T$(Fm, Optional Bexpr$)
'QLib.Std.MDao_Sql.Function SqyCrtPkzTny(A$()) As String()
'QLib.Std.MDao_Sql.Function SqlSel_F_Fm_F_Ev$(F, Fm, WhFld, Ev)
'QLib.Std.MDao_Sql.Function BexprzFnyVy$(Fny$(), Vy())
'QLib.Std.MDao_Sql.Function Bexpr$(F, Ev)
'QLib.Std.MDao_Ssk.Function SkFnyzTd(T As Dao.TableDef) As String()
'QLib.Std.MDao_Ssk.Function SkFny(A As Database, T) As String()
'QLib.Std.MDao_Ssk.Function Sskv(A As Database, T) As Aset
'QLib.Std.MDao_Ssk.Function SkIdxzTd(T As Dao.TableDef) As Dao.Index
'QLib.Std.MDao_Ssk.Function SkIdx(A As Database, T) As Dao.Index
'QLib.Std.MDao_Ssk.Function SskFld$(Db As Database, T)
'QLib.Std.MDao_Ssk.Sub DltRecNotInSskv(Db As Database, SskTbl, NotInSSskv As Aset) 'Delete Db-T record for those record's Sk not in NotInSSskv, 'Assume T has single-fld-sk
'QLib.Std.MDao_Ssk.Function AsetzDbtf(A As Database, T, F) As Aset
'QLib.Std.MDao_Ssk.Function SskVset(Db As Database, T) As Aset
'QLib.Std.MDao_Ssk.Sub InsReczSskv(A As Database, T, ToInsSSskv As Aset) 'Insert Single-Field-Secondary-Key-Aset-A into Dbt
'QLib.Std.MDao_Ssk.Private Sub Z()
'QLib.Std.MDao_Tbl.Function Fny(A As Database, T, Optional NoReOpn As Boolean) As String()
'QLib.Std.MDao_Tbl.Function ColzRs(A As Database, T, Optional F = 0) As Dao.Recordset
'QLib.Std.MDao_Tbl.Function ColSetzT(A As Database, T, Optional F = 0) As Aset
'QLib.Std.MDao_Tbl.Function CntDiczTF(A As Database, T, F) As Dictionary
'QLib.Std.MDao_Tbl.Function IdxzTd(A As Dao.TableDef, Nm) As Dao.Index
'QLib.Std.MDao_Tbl.Function Idx(A As Database, T, Nm) As Dao.Index
'QLib.Std.MDao_Tbl.Function HasSk(A As Database, T) As Boolean
'QLib.Std.MDao_Tbl.Function HasIdx(A As Database, T, IdxNm) As Boolean
'QLib.Std.MDao_Tbl.Function FstUniqIdx(A As Database, T) As Dao.Index
'QLib.Std.MDao_Tbl.Function HasFld(A As Database, T, F) As Boolean
'QLib.Std.MDao_Tbl.Function HasPk(A As Database, T) As Boolean
'QLib.Std.MDao_Tbl.Function HasPkzTd(A As Dao.TableDef) As Boolean
'QLib.Std.MDao_Tbl.Function HasStdSkzTd(A As Dao.TableDef) As Boolean
'QLib.Std.MDao_Tbl.Function HasStdPkzTd(A As Dao.TableDef) As Boolean
'QLib.Std.MDao_Tbl.Function HasStdPk(A As Database, T) As Boolean
'QLib.Std.MDao_Tbl.Function HasIdz(A As Database, T, Id&) As Boolean
'QLib.Std.MDao_Tbl.Function DryzTFF(A As Database, T, FF) As Variant()
'QLib.Std.MDao_Tbl.Sub AsgColApzDrsFF(Drs As Drs, FF, ParamArray OColAp())
'QLib.Std.MDao_Tbl.Function RszId(A As Database, T, Id) As Dao.Recordset
'QLib.Std.MDao_Tbl.Function CsvLyzDbt(A As Database, T) As String()
'QLib.Std.MDao_Tbl.Function DrszT(A As Database, T) As Drs
'QLib.Std.MDao_Tbl.Function DryzT(A As Database, T) As Variant()
'QLib.Std.MDao_Tbl.Function DtzT(A As Database, T) As Dt
'QLib.Std.MDao_Tbl.Function FdStrAy(A As Database, T) As String()
'QLib.Std.MDao_Tbl.Function Fds(A As Database, T) As Dao.Fields
'QLib.Std.MDao_Tbl.Sub ReSeqFldzFny(A As Database, T, Fny$())
'QLib.Std.MDao_Tbl.Function SrcFbzT$(A As Database, T)
'QLib.Std.MDao_Tbl.Function NColzT&(A As Database, T)
'QLib.Std.MDao_Tbl.Function NReczDbtBexpr&(A As Database, T, Bexpr$)
'QLib.Std.MDao_Tbl.Function PkFnyzTd(A As Dao.TableDef) As String()
'QLib.Std.MDao_Tbl.Function PkFny(A As Database, T) As String()
'QLib.Std.MDao_Tbl.Function PkIdxNm$(A As Database, T)
'QLib.Std.MDao_Tbl.Function NewPkIdxd(A As Dao.TableDef) As Dao.Index
'QLib.Std.MDao_Tbl.Function PkIdx(A As Database, T) As Dao.Index
'QLib.Std.MDao_Tbl.Function RszTFF(A As Database, T, FF) As Dao.Recordset
'QLib.Std.MDao_Tbl.Function RszTF(A As Database, T, F) As Dao.Recordset
'QLib.Std.MDao_Tbl.Function RszT(A As Database, T) As Dao.Recordset
'QLib.Std.MDao_Tbl.Function FdzTF(A As Database, T, F) As Dao.Field2
'QLib.Std.MDao_Tbl.Function SqzT(A As Database, T, Optional ExlFldNm As Boolean) As Variant()
'QLib.Std.MDao_Tbl.Function SrcTn$(A As Database, T)
'QLib.Std.MDao_Tbl.Function StruzT$(A As Database, T)
'QLib.Std.MDao_Tbl.Function LasUpdTimz(A As Database, T) As Date
'QLib.Std.MDao_Tbl.Sub InsDrsz(A As Database, T, Drs As Drs)
'QLib.Std.MDao_Tbl.Sub AddFd(A As Database, T, Fd As Dao.Fields)
'QLib.Std.MDao_Tbl.Sub AddFld(A As Database, T, F, Ty As DataTypeEnum, Optional Si%, Optional Precious%)
'QLib.Std.MDao_Tbl.Sub RenTbl(A As Database, T, ToNm)
'QLib.Std.MDao_Tbl.Sub RenTblzAddPfx(A As Database, T, Pfx)
'QLib.Std.MDao_Tbl.Sub BrwTblzByDt(A As Database, T)
'QLib.Std.MDao_Tbl.Function IsSysTbl(A As Database, T) As Boolean
'QLib.Std.MDao_Tbl.Function IsHidTbl(A As Database, T) As Boolean
'QLib.Std.MDao_Tbl.Function Lnkinf(A As Database) As String()
'QLib.Std.MDao_Tbl.Function LnkinfzT$(A As Database, T)
'QLib.Std.MDao_Tbl.Function Acs() As Access.Application
'QLib.Std.MDao_Tbl.Function CrtTblzDupKey$(A As Database, Into, FmTbl, KK$)
'QLib.Std.MDao_Tbl.Private Sub Z_CrtTblzDrs()
'QLib.Std.MDao_Tbl.Sub CrtTblzDrs(A As Database, T, Drs As Drs)
'QLib.Std.MDao_Tbl.Sub CrtTblzDrszAllStr(A As Database, T, Drs As Drs)
'QLib.Std.MDao_Tbl.Sub CrtTblzDrszEmpzAllStr(A As Database, T, Drs As Drs)
'QLib.Std.MDao_Tbl.Sub CrtTblzDrszEmp(A As Database, T, Drs As Drs)
'QLib.Std.MDao_Tbl.Private Sub Z_ShtTyBqlzDrs()
'QLib.Std.MDao_Tbl.Function ShtTyBqlzDrs$(Drs As Drs)
'QLib.Std.MDao_Tbl.Private Function ShtTyscfzCol$(Col, F)
'QLib.Std.MDao_Tbl.Function ShtTyzNumCol$(Col)
'QLib.Std.MDao_Tbl.Function IsColzMem(Col) As Boolean
'QLib.Std.MDao_Tbl.Function IsColzStr(Col) As Boolean
'QLib.Std.MDao_Tbl.Function DaoTyzNumCol$(NumCol)
'QLib.Std.MDao_Tbl.Function IsColzNum(Col) As Boolean
'QLib.Std.MDao_Tbl.Function IsNumzVbTy(A As VbVarType) As Boolean
'QLib.Std.MDao_Tbl.Private Function MaxNumVbTy(A As VbVarType, B As VbVarType) As VbVarType
'QLib.Std.MDao_Tbl.Function ShtTyzNumVbTy$(NumVbTy As VbVarType)
'QLib.Std.MDao_Tbl.Function IsColzBool(Col) As Boolean
'QLib.Std.MDao_Tbl.Function IsColzDte(Col) As Boolean
'QLib.Std.MDao_Tbl.Sub InsTblzDry(A As Database, T, Dry())
'QLib.Std.MDao_Tbl.Sub CrtTblzJnFld(A As Database, T, KK, JnFld$, Optional Sep$ = " ")
'QLib.Std.MDao_Tbl.Function FldIx%(A As Database, T, Fld)
'QLib.Std.MDao_Tbl.Sub CrtPk(A As Database, T)
'QLib.Std.MDao_Tbl.Function JnQSqCommaSpcAp$(ParamArray Ap())
'QLib.Std.MDao_Tbl.Function CommaSpcSqAv$(Av())
'QLib.Std.MDao_Tbl.Function JnCommaSpcFF$(FF)
'QLib.Std.MDao_Tbl.Sub CrtSk(A As Database, T, SkFF)
'QLib.Std.MDao_Tbl.Sub DrpFld(A As Database, T, FF)
'QLib.Std.MDao_Tbl.Sub RenFld(A As Database, T, F, ToFld)
'QLib.Std.MDao_Tbl.Sub UpdValIdFldz(A As Database, T, ValFld, ValIdFld)
'QLib.Std.MDao_Tbl.Function FdStrzTF$(A As Database, T, F)
'QLib.Std.MDao_Tbl.Function IntAyzDbtf(A As Database, T, F) As Integer()
'QLib.Std.MDao_Tbl.Function NxtId&(Db As Database, T)
'QLib.Std.MDao_Tbl.Function DaoTyzTF(A As Database, T, F) As Dao.DataTypeEnum
'QLib.Std.MDao_Tbl.Function ShtTyzTF$(Db As Database, T, F)
'QLib.Std.MDao_Tbl.Function CnStrzLnkTbl$(Db As Database, T)
'QLib.Std.MDao_Tbl.Sub AddFldzExpr(Db As Database, T, F, Expr$, Ty As Dao.DataTypeEnum)
'QLib.Std.MDao_Tbl.Function ValOfTFRecId(Db As Database, T, F, RecId&) ' K is Pk value
'QLib.Std.MDao_Tbl.Sub CrtTblzEmpClone(Db As Database, T, FmTbl)
'QLib.Std.MDao_Tbl.Sub KillTmpDb(Db As Database)
'QLib.Std.MDao_Tbl.Private Sub Z_CrtDupKeyTbl()
'QLib.Std.MDao_Tbl.Private Sub Z_PkFny()
'QLib.Std.MDao_Tbl.Private Sub ZZ()
'QLib.Std.MDao_Tbl.Function ValOfArs(A As ADODB.Recordset)
'QLib.Std.MDao_Tbl.Function ValOfCnq(A As ADODB.Connection, Q)
'QLib.Std.MDao_Tbl.Function NReczFxw&(Fx, Wsn, Optional Bexpr$)
'QLib.Std.MDao_Tbl.Function NReczT&(A As Database, T, Optional Bexpr$)
'QLib.Std.MDao_Tbl.Property Get LofVblPrp$(A As Database, T)
'QLib.Std.MDao_Tbl.Property Let LofVblPrp(A As Database, T, LofVbl$)
'QLib.Std.MDao_Tbl.Function IsLnk(A As Database, T) As Boolean
'QLib.Std.MDao_Tbl.Function CnStrzT$(A As Database, T)
'QLib.Std.MDao_Tbl.Function IsLnkzFb(A As Database, T) As Boolean
'QLib.Std.MDao_Tbl.Function IsLnkzFx(A As Database, T) As Boolean
'QLib.Std.MDao_Tbl_ReseqFld.Sub ReSeqFldzSpec(A As Database, T, ReSeqSpec$)
'QLib.Std.MDao_Tbl_ReseqFld.Private Sub ZZ_FnyzReseqSpec()
'QLib.Std.MDao_Tbl_ReseqFld.Function FnyzReseqSpec(ReSeqSpec$) As String()
'QLib.Std.MDao_Tbl_ReseqFld.Sub ReSeqFldzFny(A As Database, T, ByFny$())
'QLib.Std.MDao_Tbl_ReseqFld.Function AyReSeq(Ay, ByAy)
'QLib.Std.MDao_Tbl_Upd_EndDteFld.Sub UpdEndDte(A As Database, T, EndDteFld$, BegDteFld$, GpFF)
'QLib.Std.MDao_Tbl_Upd_SeqFld.Sub UpdSeqFld(A As Database, T, SeqFld$, GpFF, OrdFFMinus$)
'QLib.Std.MDao_Tbl_Upd_SeqFld.Private Sub ZZ_UpdSeqFld()
'QLib.Std.MDao_Tmp.Property Get TmpTd() As Dao.TableDef
'QLib.Std.MDao_Tmp.Property Get TmpDbPth$()
'QLib.Std.MDao_Tmp.Function TmpDb(Optional Fdr$, Optional Fnn$) As Database
'QLib.Std.MDao_Tmp.Function TmpFb$()
'QLib.Std.MDao_Ty.Property Get ShtTyDrs() As Drs
'QLib.Std.MDao_Ty.Property Get VdtShtTyAy() As String()
'QLib.Std.MDao_Ty.Property Get VdtShtTyDtaTyAy() As String()
'QLib.Std.MDao_Ty.Property Get VdtDtaTyAy() As String()
'QLib.Std.MDao_Ty.Function IsShtTy(A) As Boolean
'QLib.Std.MDao_Ty.Function DaoTyzShtTy(ShtTy) As Dao.DataTypeEnum
'QLib.Std.MDao_Ty.Function SqlTyzDao$(T As Dao.DataTypeEnum, Optional Si%, Optional Precious%)
'QLib.Std.MDao_Ty.Function ShtTyzDao$(A As Dao.DataTypeEnum)
'QLib.Std.MDao_Ty.Function DtaTyzTF$(A As Database, T, F)
'QLib.Std.MDao_Ty.Function DtaTy$(T As Dao.DataTypeEnum)
'QLib.Std.MDao_Ty.Function DaoTyzDtaTy(DtaTy) As Dao.DataTypeEnum
'QLib.Std.MDao_Ty.Function DaoTyzVbTy(A As VbVarType) As Dao.DataTypeEnum
'QLib.Std.MDao_Ty.Function DaoTyzVal(V) As Dao.DataTypeEnum
'QLib.Std.MDao_Ty.Function CvDaoTy(A) As Dao.DataTypeEnum
'QLib.Std.MDao_Ty.Function ShtTyLiszDaoTyAy$(A() As DataTypeEnum)
'QLib.Std.MDao_Ty.Function ErzShtTyLis(ShtTyLis$) As String()
'QLib.Std.MDao_Ty.Function IsVdtShtTy(A) As Boolean
'QLib.Std.MDao_Ty.Function DtaTyAyzShtTyAy(ShtTyAy$()) As String()
'QLib.Std.MDao_Ty.Function DtaTyzShtTy$(ShtTy)
'QLib.Std.MDao_Ty.Function ShtTyzAdo$(A As ADODB.DataTypeEnum)
'QLib.Std.MDao_Ty.Function ShtTyAyzShtTyLis(ShtTyLis$) As String()
'QLib.Std.MDao_Ty.Sub ThwShtTyEr(Fun$, ShtTy)
'QLib.Std.MDao_Ty_Ado.Function ShtAdoTyAy(A() As ADODB.DataTypeEnum) As String()
'QLib.Std.MDao_Ty_Ado.Function ShtAdoTy$(A As ADODB.DataTypeEnum)
'QLib.Std.MDao_Val.Private Sub Z_ValOfQ()
'QLib.Std.MDao_Val.Property Get ValOfQ(A As Database, Sql)
'QLib.Std.MDao_Val.Property Let ValOfQ(A As Database, Sql, V)
'QLib.Std.MDao_Val.Property Let ValOfSsk(Db As Database, T, F, Sskv, V)
'QLib.Std.MDao_Val.Property Get ValOfSsk(Db As Database, T, F, Sskv)
'QLib.Std.MDao_Val.Function ValOfTF(A As Database, T, F)
'QLib.Std.MDao_Val.Function ValOfQQ(A As Database, QQSql, ParamArray Ap())
'QLib.Std.MDao_Val.Property Let ValOfRs(A As Dao.Recordset, V)
'QLib.Std.MDao_Val.Property Get ValOfRs(A As Dao.Recordset)
'QLib.Std.MDao_Val.Property Let ValOfRsFld(Rs As Dao.Recordset, Fld, V)
'QLib.Std.MDao_Val.Property Get ValOfRsFld(Rs As Dao.Recordset, Fld)
'QLib.Std.MDta_Ay.Function AyDt(A, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
'QLib.Std.MDta_Ay.Function DryGpCntzAy(A) As Variant()
'QLib.Std.MDta_Ay.Function DryGpCntzAyWhDup(A) As Variant()
'QLib.Std.MDta_Ay.Sub BrwDryGpCntzAy(Ay)
'QLib.Std.MDta_Ay.Function FmtDryGpCntzAy(Ay) As String()
'QLib.Std.MDta_Ay.Private Sub ZZ_FmtDryGpCntzAy()
'QLib.Std.MDta_Ay.Private Sub ZZ_CntDryzAy()
'QLib.Std.MDta_Dic.Function DrszDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val") As Drs
'QLib.Std.MDta_Dic.Function DtzDic(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValOptTy As Boolean) As Dt
'QLib.Std.MDta_Dic.Function FnyzDic(Optional InclValTy As Boolean) As String()
'QLib.Std.MDta_Dr.Function DrzTLinVbTyAy(TLin, VbTyAy() As VbVarType) As Variant()
'QLib.Std.MDta_Dr.Function VbTyzShtTy(ShtTy$) As VbVarType
'QLib.Std.MDta_Dr.Function VbTyAyzShtTyLis(ShtTyLis$) As VbVarType()
'QLib.Std.MDta_Dr.Function DrzTLinShtTyLis(TLin, ShtTyLis$) As Variant()
'QLib.Std.MDta_Dr.Function DrzDrs(A As Drs, Optional CC, Optional Row&)
'QLib.Std.MDta_Drs.Function CvDrs(A) As Drs
'QLib.Std.MDta_Drs.Function Drs(FF, Dry()) As Drs
'QLib.Std.MDta_Drs.Function DrsAddCol(A As Drs, ColNm$, ConstVal) As Drs
'QLib.Std.MDta_Drs.Function DrsAddIxCol(A As Drs, HidIxCol As Boolean) As Drs
'QLib.Std.MDta_Drs.Function IsDrs(A) As Boolean
'QLib.Std.MDta_Drs.Function AvDrsC(A As Drs, C) As Variant()
'QLib.Std.MDta_Drs.Function IntoDrsC(Into, A As Drs, C)
'QLib.Std.MDta_Drs.Sub DmpDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$)
'QLib.Std.MDta_Drs.Function DrsDrpCC(A As Drs, CC) As Drs
'QLib.Std.MDta_Drs.Function DrsSelCC(A As Drs, CC) As Drs
'QLib.Std.MDta_Drs.Function DrySelColIxAy(Dry(), IxAy&()) As Variant()
'QLib.Std.MDta_Drs.Function DtDrsDtnm(A As Drs, DtNm$) As Dt
'QLib.Std.MDta_Drs.Function DrsInsCV(A As Drs, C$, V) As Drs
'QLib.Std.MDta_Drs.Function DrsInsCVAft(A As Drs, C$, V, AftFldNm$) As Drs
'QLib.Std.MDta_Drs.Function DrsInsCVBef(A As Drs, C$, V, BefFldNm$) As Drs
'QLib.Std.MDta_Drs.Private Function DrsInsCVIsAftFld(A As Drs, C$, V, IsAft As Boolean, FldNm$) As Drs
'QLib.Std.MDta_Drs.Function IsEqDrs(A As Drs, B As Drs) As Boolean
'QLib.Std.MDta_Drs.Sub BrwCnt(Ay, Optional IgnCas As Boolean, Optional Opt As eCntOpt)
'QLib.Std.MDta_Drs.Function DicItmWdt%(A As Dictionary)
'QLib.Std.MDta_Drs.Private Function CntLyzCntDic(CntDic As Dictionary, CntWdt%) As String()
'QLib.Std.MDta_Drs.Function CntLy(Ay, Optional Opt As eCntOpt, Optional SrtOpt As eCntSrtOpt, Optional IsDesc As Boolean, Optional IgnCas As Boolean) As String()
'QLib.Std.MDta_Drs.Function CntDic(Ay, Optional IgnCas As Boolean, Optional Opt As eCntOpt) As Dictionary
'QLib.Std.MDta_Drs.Function CntDiczDrs(A As Drs, C$) As Dictionary
'QLib.Std.MDta_Drs.Function NColzDrs%(A As Drs)
'QLib.Std.MDta_Drs.Function NRowDrs&(A As Drs)
'QLib.Std.MDta_Drs.Function DrwIxAy(Dr, IxAy)
'QLib.Std.MDta_Drs.Function DrySelIxAy(Dry(), IxAy) As Variant()
'QLib.Std.MDta_Drs.Function DrsReOrdBy(A As Drs, BySubFF) As Drs
'QLib.Std.MDta_Drs.Function NRowDrsCEv&(A As Drs, ColNm$, EqVal)
'QLib.Std.MDta_Drs.Function SqzDrs(A As Drs) As Variant()
'QLib.Std.MDta_Drs.Function SyDrsC(A As Drs, ColNm) As String()
'QLib.Std.MDta_Drs.Sub PushDrs(O As Drs, A As Drs)
'QLib.Std.MDta_Drs.Private Sub ZZ_GpDicDKG()
'QLib.Std.MDta_Drs.Private Sub ZZ_CntDiczDrs()
'QLib.Std.MDta_Drs.Private Sub ZZ_DrsSel()
'QLib.Std.MDta_Drs.Private Property Get Z_FmtDrs()
'QLib.Std.MDta_Drs.Private Sub ZZ()
'QLib.Std.MDta_Drs.Private Sub Z()
'QLib.Std.MDta_Drs.Function DrsAddCC(A As Drs, FF, C1, C2) As Drs
'QLib.Std.MDta_Drs_Dup.Function DrswDup(A As Drs, FF) As Drs
'QLib.Std.MDta_Drs_Dup.Function DrseDup(A As Drs, FF) As Drs
'QLib.Std.MDta_Drs_Dup.Private Function RowIxAyzDupzDrs(A As Drs, FF) As Long()
'QLib.Std.MDta_Drs_Dup.Private Function RowIxAyzDupzDry(Dry()) As Long()
'QLib.Std.MDta_Drs_Dup.Function DrywDup(Dry()) As Variant()
'QLib.Std.MDta_Drs_Dup.Function DrywDist(Dry()) As Variant()
'QLib.Std.MDta_Drs_Dup.Function DryGpCnt(Dry()) As Variant()
'QLib.Std.MDta_Drs_Dup.Private Function DryGpCntzQuick(Dry()) As Variant()
'QLib.Std.MDta_Drs_Dup.Private Function DryGpCntzSlow(Dry()) As Variant()
'QLib.Std.MDta_Drs_Dup.Private Function IxOptzDryDr(Dry(), Dr) As LngRslt
'QLib.Std.MDta_Drs_Dup.Private Sub Z_DrswDup()
'QLib.Std.MDta_Drs_Dup.Private Function RowIxAyzDupzDryColIx(Dry(), ColIx&) As Long()
'QLib.Std.MDta_Drs_Dup.Private Sub Z_RowIxAyzDupzDryColIx()
'QLib.Std.MDta_Dry.Function IxAyzCC(CC) As Integer()
'QLib.Std.MDta_Dry.Function IntAyzIIStr(IIStr$) As Integer()
'QLib.Std.MDta_Dry.Function CntDryWhGt1(CntDry()) As Variant()
'QLib.Std.MDta_Dry.Function DrywColInAy(A(), ColIx%, InAy) As Variant()
'QLib.Std.MDta_Dry.Sub C3DryDo3(C3Dry(), Do3$)
'QLib.Std.MDta_Dry.Sub C4DryDo4(C4Dry(), Do4$)
'QLib.Std.MDta_Dry.Function DotNyDry(A()) As String()
'QLib.Std.MDta_Dry.Function DryDotNy(DotNy$()) As Variant()
'QLib.Std.MDta_Dry.Private Sub ZZ_FmtA()
'QLib.Std.MDta_Dry.Function DrywColHasDup(A(), C) As Variant()
'QLib.Std.MDta_Dry.Private Sub Z_DryzJnFldKK()
'QLib.Std.MDta_Dry.Function DrsJnFldKKFld(Drs As Drs, KK, JnFld, Optional Sep$ = " ") As Drs
'QLib.Std.MDta_Dry.Function DryzJnFldKK(Dry(), KKIxAy, JnFldIx, Optional Sep$ = " ") As Variant()
'QLib.Std.MDta_Dry.Function RowIxOptzDryDr&(Dry(), Dr)
'QLib.Std.MDta_Dry.Function DryJnFldNFld(Dry(), FstNFld%, Optional Sep$ = " ") As Variant()
'QLib.Std.MDta_Dry.Function DryzSslAy(SslAy) As Variant()
'QLib.Std.MDta_Dry.Function CntDiczDry(A(), C) As Dictionary
'QLib.Std.MDta_Dry.Function SeqCntDiczDry(A(), C) As Dictionary
'QLib.Std.MDta_Dry.Function SqlTyzDryC$(A(), C)
'QLib.Std.MDta_Dry.Function SqlTyzAv$(Av())
'QLib.Std.MDta_Dry.Function SqlTyzVbTy$(A As VbVarType, Optional IsMem As Boolean)
'QLib.Std.MDta_Dry.Function IsBrkDryIxC(A(), DrIx&, BrkColIx) As Boolean
'QLib.Std.MDta_Dry.Function NColzDry%(A)
'QLib.Std.MDta_Dry.Function NRowDryCEv&(A(), C, Ev)
'QLib.Std.MDta_Dry.Function DrywCEv(A(), C, Ev) As Variant()
'QLib.Std.MDta_Dry.Function DrywCCNe(A, C1, C2) As Variant()
'QLib.Std.MDta_Dry.Sub ThwIfNEDry(A(), B())
'QLib.Std.MDta_Dry.Function DrywColEq(A, C%, V) As Variant()
'QLib.Std.MDta_Dry.Function DrywCGt(A, C%, GtV) As Variant()
'QLib.Std.MDta_Dry.Function DrywDupCC(Dry(), CC) As Variant()
'QLib.Std.MDta_Dry.Private Function DrywDupCol(Dry(), ColIx%) As Variant()
'QLib.Std.MDta_Dry.Function DrywIxAyzy(A, IxAy, EqVy) As Variant()
'QLib.Std.MDta_Dry.Function DistSyzDry(A(), C) As String()
'QLib.Std.MDta_Dry.Function DryzSqCol(Sq, ColIxAy) As Variant()
'QLib.Std.MDta_Dry.Function DryzSq(Sq) As Variant()
'QLib.Std.MDta_Dry.Function HasDr(Dry(), Dr) As Boolean
'QLib.Std.MDta_Dry_Col_Add.Function DryAddColz(Dry(), C) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryAddColzC3(A(), C1, C2, C3) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryAddColzBy(Dry(), Optional ByNCol% = 1) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryAddColzC(Dry(), C) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryAddColzCC(Dry(), V1, V2) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryAddColzAv(Dry(), Av()) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryInsColzAv(Dry(), Av()) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryInsColzV3(Dry(), V1, V2, V3) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryInsColzV(A(), V, Optional At& = 0) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryInsColzV4(A(), V1, V2, V3, V4) As Variant()
'QLib.Std.MDta_Dry_Col_Add.Function DryInsColzV2(A(), V1, V2) As Variant()
'QLib.Std.MDta_Dry_Col_Rmv.Function RmvColzDryC(A(), C) As Variant()
'QLib.Std.MDta_Dry_Col_Rmv.Function RmvColzDryIxAy(Dry(), IxAy) As Variant()
'QLib.Std.MDta_Ds.Function DsAddDt(A As Ds, T As Dt) As Ds
'QLib.Std.MDta_Ds.Function CvDs(A) As Ds
'QLib.Std.MDta_Ds.Function DsAddDtAy(A As Ds, B() As Dt) As Ds
'QLib.Std.MDta_Ds.Function DsDt(A As Ds, Ix%) As Dt
'QLib.Std.MDta_Ds.Function Ds(A() As Dt, Optional DsNm$ = "Ds") As Ds
'QLib.Std.MDta_Ds.Function DsHasDt(A As Ds, DtNm) As Boolean
'QLib.Std.MDta_Ds.Function DsIsEmp(A As Ds) As Boolean
'QLib.Std.MDta_Ds.Function DsNDt%(A As Ds)
'QLib.Std.MDta_Dt.Function DtAddAp(A As Dt, ParamArray DtAp()) As Dt()
'QLib.Std.MDta_Dt.Function IsDt(A) As Boolean
'QLib.Std.MDta_Dt.Sub BrwDs(A As Ds, Optional Fnn$)
'QLib.Std.MDta_Dt.Sub BrwDt(A As Dt, Optional Fnn$)
'QLib.Std.MDta_Dt.Function CsvQQStrDr$(Dr)
'QLib.Std.MDta_Dt.Function CsvLyDt(A As Dt) As String()
'QLib.Std.MDta_Dt.Function DtSelCol(A As Dt, CC, Optional DtNm$) As Dt
'QLib.Std.MDta_Dt.Function DtDrpCol(A As Dt, CC, Optional DtNm$) As Dt
'QLib.Std.MDta_Dt.Function DrszDt(A As Dt) As Drs
'QLib.Std.MDta_Dt.Function DtzDrs(A As Drs, Optional DtNm$ = "Dt") As Dt
'QLib.Std.MDta_Dt.Function NRowzDt&(A As Dt)
'QLib.Std.MDta_Dt.Function NRowzDrs&(A As Drs)
'QLib.Std.MDta_Dt.Sub DmpDt(A As Dt)
'QLib.Std.MDta_Dt.Property Get EmpDtAy() As Dt()
'QLib.Std.MDta_Dt.Function IsEmpDt(A As Dt) As Boolean
'QLib.Std.MDta_Dt.Function DtReOrd(A As Dt, BySubFF) As Dt
'QLib.Std.MDta_Dt.Function Dt(DtNm, Fny0, Dry()) As Dt
'QLib.Std.MDta_Dt.Function CvDt(A) As Dt
'QLib.Std.MDta_Dt.Function DtAy(ParamArray Ap()) As Dt()
'QLib.Std.MDta_Dt.Private Sub ZZ_DtAy()
'QLib.Std.MDta_ExpLines.Function DrExpLinesCol(Dr, LinesColIx%) As Variant()
'QLib.Std.MDta_ExpLines.Function DrsExpLinesCol(A As Drs, LinesColNm$) As Drs
'QLib.Std.MDta_Fmt.Sub VcDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNN, Optional Fnn$)
'QLib.Std.MDta_Fmt.Sub BrwDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNN, Optional Fnn$, Optional UseVc As Boolean)
'QLib.Std.MDta_Fmt.Function FmtDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNN, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
'QLib.Std.MDta_Fmt.Function FmtDs(A As Ds, Optional MaxColWdt% = 100, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
'QLib.Std.MDta_Fmt.Function FmtDt(A As Dt, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
'QLib.Std.MDta_Fmt.Private Sub Z_FmtDrs()
'QLib.Std.MDta_Fmt.Private Sub Z_FmtDt()
'QLib.Std.MDta_Fmt.Private Sub ZZ()
'QLib.Std.MDta_Fmt.Private Sub Z()
'QLib.Std.MDta_Fmt_Dry.Private Sub A_Main()
'QLib.Std.MDta_Fmt_Dry.Sub BrwDry(A(), Optional MaxColWdt% = 100, Optional BrkCC, Optional ShwZer As Boolean)
'QLib.Std.MDta_Fmt_Dry.Sub BrwDryzSpc(A(), Optional MaxColWdt% = 100, Optional ShwZer As Boolean)
'QLib.Std.MDta_Fmt_Dry.Function FmtDry(Dry(), Optional MaxColWdt% = 100, Optional BrkCC, Optional ShwZer As Boolean) As String()
'QLib.Std.MDta_Fmt_Dry.Sub DmpDryAsSpcSep(Dry())
'QLib.Std.MDta_Fmt_Dry.Sub DmpDry(Dry())
'QLib.Std.MDta_Fmt_Dry.Function FmtDryAsSpcSep(Dry(), Optional MaxColWdt% = 100, Optional ShwZer As Boolean) As String()
'QLib.Std.MDta_Fmt_Dry_Fun.Function AyExtend(Ay, ToSamSzAsThisAy)
'QLib.Std.MDta_Fmt_Dry_Fun.Function AyFmtToWdtAy(Dr, ToWdt%()) As String() 'Fmt-Dr-ToWdt
'QLib.Std.MDta_Fmt_Dry_Fun.Function DryzAySepSS(Ay, SepSS$) As Variant()
'QLib.Std.MDta_Fmt_Dry_Fun.Function DryFmtCellAsStr(A()) As Variant()
'QLib.Std.MDta_Fmt_Dry_Fun.Function DrFmtCellSpcSep(Dr) As String()
'QLib.Std.MDta_Fmt_Dry_Fun.Private Sub Z_DrzLinSepAy()
'QLib.Std.MDta_Fmt_Dry_Fun.Function ShfTermFmSep$(OLin, FmPos%, Sep$)
'QLib.Std.MDta_Fmt_Dry_Fun.Function DrzLinSepAy(Lin, SepAy$()) As String()
'QLib.Std.MDta_Fmt_Dry_Fun.Function FmtAyzSepSS(Ay, SepSS$) As String()
'QLib.Std.MDta_Fmt_Dry_Fun.Function DryFmtCommon(Dry(), MaxColWdt%, ShwZer As Boolean) As Variant()
'QLib.Std.MDta_Fmt_Dry_Fun.Function DrAlign(Dr, W%()) As String()
'QLib.Std.MDta_Fmt_Dry_Fun.Function DryAlignColzWdt(Dry(), W%()) As Variant()
'QLib.Std.MDta_Fmt_Dry_Fun.Function DryAlignCol(Dry()) As Variant()
'QLib.Std.MDta_Fmt_Dry_Fun.Function DryMkCell(Dry(), ShwZer As Boolean, MaxColWdt%) As Variant()
'QLib.Std.MDta_Fmt_Dry_Fun.Private Function DrMkCell(Dr, ShwZer As Boolean, MaxWdt%) As String()
'QLib.Std.MDta_Fmt_Dry_Fun.Function MkCell$(V, Optional ShwZer As Boolean, Optional MaxWdt% = 30) ' Convert V into a string fit in a cell
'QLib.Std.MDta_Fmt_Dry_Fun.Function IsEqDrCC(Dr1, Dr2, CC%()) As Boolean
'QLib.Std.MDta_Fmt_Dry_Fun.Function DryInsBrk(SrtedDry, BrkCC, SepDr$()) As Variant()
'QLib.Std.MDta_Fmt_Dry_Fun.Function WdtAyzDry(A()) As Integer()
'QLib.Std.MDta_Fmt_Dry_Fun.Function SepLin$(W%(), Sep$)
'QLib.Std.MDta_Fmt_Dry_Fun.Function SepDr(W%()) As String()
'QLib.Std.MDta_Fmt_Dry_Fun.Function SepLinzSepDr$(SepDr$(), Sep$)
'QLib.Std.MDta_Fmt_Dry_Fun.Function LinFmDrByJnCell$(Dr, Optional Sep$ = "|", Optional Pfx$ = "|", Optional Sfx$ = "|")
'QLib.Std.MDta_Fmt_Dry_Fun.Function FmtDryByJnCell(Dry(), Optional Sep$ = "|", Optional Pfx$ = "|", Optional Sfx$ = "|") As String()
'QLib.Std.MDta_Fmt_Wrp.Function WrpDrNRow%(WrpDr())
'QLib.Std.MDta_Fmt_Wrp.Function WrpDrPad(WrpDr, W%()) As Variant() 'Some Cell in WrpDr may be an array, Wrping each element to cell if their width can fit its W%(?)
'QLib.Std.MDta_Fmt_Wrp.Function WrpDrSq(WrpDr()) As Variant()
'QLib.Std.MDta_Fmt_Wrp.Function WrpDryWdt(WrpDry(), WrpWdt%) As Integer() 'WrpDry is dry having 1 or more wrpCol, which mean need Wrpping.
'QLib.Std.MDta_Fmt_Wrp.Function WrpCellDr(A, ColWdt%()) As String()
'QLib.Std.MDta_Fmt_Wrp.Function FmtDrWrp(WrpDr, W%()) As String() 'Each Itm of WrpDr may be an array.  So a AyFmtToWdtAy return Ly not string.
'QLib.Std.MDta_Fmt_Wrp.Function DryWrpCell(A(), Optional WrpWdt% = 40) As String() 'WrpWdt is for wrp-col.  If maxWdt of an ele of wrp-col > WrpWdt, use the maxWdt
'QLib.Std.MDta_Fmt_Wrp.Function SqAlign(Sq(), W%()) As Variant()
'QLib.Std.MDta_ObjPrp.Function DrszItrPP(Itr, PP_MayWith_NewFldEqQuoteFmFld$) As Drs
'QLib.Std.MDta_ObjPrp.Function DrszOyPP(Oy, PP_MayWith_NewFldEqQuoteFmFld$) As Drs
'QLib.Std.MDta_ObjPrp.Private Function WFmlEr(PrpAy$(), PPzFml$()) As String()
'QLib.Std.MDta_ObjPrp.Private Sub WAsg3PP(PP_with_NewFldEqQuoteFmFld$, OPPzPrp$(), OPPzFml$(), OPPzAll$())
'QLib.Std.MDta_ObjPrp.Private Function DrsAddFml(A As Drs, PPzFml$()) As Drs
'QLib.Std.MDta_ObjPrp.Function AddColzFmlDrs(A As Drs, NewFld, FunNm$, PmAy$()) As Drs
'QLib.Std.MDta_ObjPrp.Private Function DrszItrPPzPure(Oy, PP) As Drs
'QLib.Std.MDta_ObjPrp.Private Function DryzItrPPzPure(Itr, PP) As Variant()
'QLib.Std.MDta_ObjPrp.Private Sub Z_DrszItrPP()
'QLib.Std.MDta_ObjPrp.Private Sub Z()
'QLib.Std.MDta_Piv.Function DryGpAy(A, KIx%, GIx%) As Variant()
'QLib.Std.MDta_Piv.Private Function KKDrIx&(KKDr, FstColIsKKDrDry)
'QLib.Std.MDta_Piv.Private Function KKDrToItmAyDualColDry(Dry(), KKColIx%(), ItmColIx%) As Variant()
'QLib.Std.MDta_Piv.Function KKCntMulItmColDry(Dry(), KKColIx%(), ItmColIx%) As Variant()
'QLib.Std.MDta_Piv.Private Function KKCntMulItmColDryD(KKDrToItmAyDualColDry()) As Variant()
'QLib.Std.MDta_Piv.Function GpDicDKG(A As Drs, KK, G$) As Dictionary
'QLib.Std.MDta_Piv.Function TwoColDryzDotLy(DotLy$()) As Variant()
'QLib.Std.MDta_Piv.Function DryzDotLy(DotAy) As Variant()
'QLib.Std.MDta_Piv.Function DryzLyWithColon(LyWithColon$()) As Variant()
'QLib.Std.MDta_Piv.Function DryGpDic(A, KeyIxAy, G) As Dictionary
'QLib.Std.MDta_Piv.Function DrszFbt(Fb, T) As Drs
'QLib.Std.MDta_Piv.Function KE24Drs() As Drs
'QLib.Std.MDta_S1S2.Function S1S2DrSumSi(A() As S1S2) As Drs
'QLib.Std.MDta_S1S2.Function S1S2AyDry(A() As S1S2) As Variant()
'QLib.Std.MDta_Samp.Property Get SampDr1() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDr2() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDr3() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDr4() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDr5() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDr6() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDrs1() As Drs
'QLib.Std.MDta_Samp.Property Get SampDrs2() As Drs
'QLib.Std.MDta_Samp.Property Get SampDrs() As Drs
'QLib.Std.MDta_Samp.Property Get SampDFnyRs() As String()
'QLib.Std.MDta_Samp.Property Get SampDry1() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDry2() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDry() As Variant()
'QLib.Std.MDta_Samp.Property Get SampDs() As Ds
'QLib.Std.MDta_Samp.Property Get SampDt1() As Dt
'QLib.Std.MDta_Samp.Property Get SampDt2() As Dt
'QLib.Std.MDta_Sel.Function DrySel(A, IxAy) As Variant()
'QLib.Std.MDta_Sel.Function DrySelIxAp(A, ParamArray IxAp()) As Variant()
'QLib.Std.MDta_Sel.Function DrsSel(A As Drs, FF) As Drs
'QLib.Std.MDta_Sel.Private Sub Z_DrsSel()
'QLib.Std.MDta_Sel.Function DtSel(A As Dt, FF) As Dt
'QLib.Std.MDta_Sel.Private Sub Z()
'QLib.Std.MDta_Srt.Private Sub Asg(Fny$(), SrtByFF$, OColIxAy%(), OIsDesAy() As Boolean) 'SrtByFF may have - as sfx which means IsDes
'QLib.Std.MDta_Srt.Function DrsSrt(A As Drs, Optional SrtByFF$ = "") As Drs 'If SrtByFF is blank use fst col.
'QLib.Std.MDta_Srt.Function DrySrt(Dry(), ColIxAy%(), IsDesAy() As Boolean) As Variant()
'QLib.Std.MDta_Srt.Function DtSrt(A As Dt, Optional SrtByFF$ = "") As Dt
'QLib.Std.MDta_Srt.Function DrySrtzCol(Dry(), ColIx%, Optional IsDes As Boolean) As Variant()
'QLib.Std.MDta_Srt.Private Function DrySrtzColIxAy(Dry(), SrtColIxAy%(), IsDesAy() As Boolean) As Variant()
'QLib.Std.MDta_Srt.Private Function IsGT(Dr1, Dr2) As Boolean
'QLib.Std.MDta_Srt.Private Function Partition&(ODry, L&, H&)
'QLib.Std.MDta_Srt.Private Sub DrySrtLH(ODry, L&, H&)
'QLib.Std.MDta_Wh.Function DrswFldEqV(A As Drs, F, EqVal) As Drs
'QLib.Std.MDta_Wh.Function DrswFFNe(A As Drs, F1, F2) As Drs 'FFNe = Two Fld Not Eq
'QLib.Std.MDta_Wh.Function DrswColEq(A As Drs, C$, V) As Drs
'QLib.Std.MDta_Wh.Function DrswColGt(A As Drs, C$, V) As Drs
'QLib.Std.MDta_Wh.Function DrseRowIxAy(A As Drs, RowIxAy&()) As Drs
'QLib.Std.MDta_Wh.Function DrswNotRowIxAy(A As Drs, RowIxAy&()) As Drs
'QLib.Std.MDta_Wh.Function DrswRowIxAy(A As Drs, RowIxAy) As Drs
'QLib.Std.MIde_Btn.Sub DltClr(A As CommandBar)
'QLib.Std.MIde_Btn.Function CtlAy(A As CommandBar) As CommandBarControl()
'QLib.Std.MIde_Btn.Function CtlCapAy(A As CommandBar) As String()
'QLib.Std.MIde_Btn.Function Bar(Nm$) As CommandBar
'QLib.Std.MIde_Btn.Property Get BarNy() As String()
'QLib.Std.MIde_Btn.Property Get BrwObjWin() As Vbide.Window
'QLib.Std.MIde_Btn.Property Get CompileBtn() As CommandBarButton
'QLib.Std.MIde_Btn.Private Function CvCtl(A) As CommandBarControl
'QLib.Std.MIde_Btn.Property Get DbgPopup() As CommandBarPopup
'QLib.Std.MIde_Btn.Property Get EdtclrBtn() As Office.CommandBarButton
'QLib.Std.MIde_Btn.Property Get MnuBar() As CommandBar
'QLib.Std.MIde_Btn.Property Get SelAllBtn() As Office.CommandBarButton
'QLib.Std.MIde_Btn.Function IsBtn(A) As Boolean
'QLib.Std.MIde_Btn.Property Get NxtStmtBtn() As CommandBarButton
'QLib.Std.MIde_Btn.Private Property Get EditPopup() As Office.CommandBarPopup
'QLib.Std.MIde_Btn.Property Get SavBtn() As CommandBarButton
'QLib.Std.MIde_Btn.Property Get JmpNxtStmtBtn() As CommandBarButton
'QLib.Std.MIde_Btn.Property Get StdBar() As Office.CommandBar
'QLib.Std.MIde_Btn.Function MnuBarzVbe(A As Vbe) As CommandBar
'QLib.Std.MIde_Btn.Function SavBtnz(A As Vbe) As CommandBarButton
'QLib.Std.MIde_Btn.Function StdBarz(A As Vbe) As Office.CommandBar
'QLib.Std.MIde_Btn.Function BarNyvVbe(A As Vbe) As String()
'QLib.Std.MIde_Btn.Function CmdBarAyzVbe(A As Vbe) As Office.CommandBar()
'QLib.Std.MIde_Btn.Function CmdBarNyvVbe(A As Vbe) As String()
'QLib.Std.MIde_Btn.Property Get WinPop() As CommandBarPopup
'QLib.Std.MIde_Btn.Property Get WinTileVertBtn() As Office.CommandBarButton
'QLib.Std.MIde_Btn.Property Get XlsBtn() As Office.CommandBarControl
'QLib.Std.MIde_Btn.Private Sub ZZ_DbgPop()
'QLib.Std.MIde_Btn.Private Sub ZZ_MnuBar()
'QLib.Std.MIde_Btn.Private Sub ZZ()
'QLib.Std.MIde_Btn.Private Sub Z()
'QLib.Std.MIde_VbCd.Function CdLyPj() As String()
'QLib.Std.MIde_VbCd.Function CdLyzMd(A As CodeModule) As String()
'QLib.Std.MIde_VbCd.Function CdLyzPj(A As VBProject) As String()
'QLib.Std.MIde_VbCd.Function CdLyzSrc(Src$()) As String()
'QLib.Std.MIde_VbCd.Function IsCdLin(A) As Boolean
'QLib.Std.MIde_VbCd.Function IsNonOptCdLin(A) As Boolean
'QLib.Std.MIde_Cmd.Function HasBar(Nm$)
'QLib.Std.MIde_Cmd.Function CvCmdBtn(A) As Office.CommandBarButton
'QLib.Std.MIde_Cmd_Action.Sub TileH()
'QLib.Std.MIde_Cmd_Action.Sub TileV()
'QLib.Std.MIde_Cmd_Action.Property Get TileVBtn() As CommandBarButton
'QLib.Std.MIde_Cmd_Lis_Src.Sub LisSrc(Patn$)
'QLib.Std.MIde_Cmd_Lis_Src.Sub CurPjLisSrc(Patn$)
'QLib.Std.MIde_Cmd_Lis_Src.Sub PjLisSrc(A As VBProject, Patn$)
'QLib.Std.MIde_Cmd_Lis_Src.Function MdNmLnoGoStr$(MdDNm$, Lno&)
'QLib.Std.MIde_Cmd_Lis_Src.Function CmpReSrc(A As VBComponent, R As RegExp) As String()
'QLib.Std.MIde_Cmd_MovMth.Property Get CmdBarNy() As String()
'QLib.Std.MIde_Cmd_MovMth.Private Sub Z_Mov_MthBar()
'QLib.Std.MIde_Cmd_MovMth.Function Vbe_Bars(A As Vbe) As Office.CommandBars
'QLib.Std.MIde_Cmd_MovMth.Property Get CurVbe_Bars() As Office.CommandBars
'QLib.Std.MIde_Cmd_MovMth.Function CurVbe_BarsHas(A) As Boolean
'QLib.Std.MIde_Cmd_MovMth.Function CmdBar(A) As Office.CommandBar
'QLib.Std.MIde_Cmd_MovMth.Sub RmvCmdBar(A)
'QLib.Std.MIde_Cmd_MovMth.Function CmdBar_HasBtn(A As Office.CommandBar, BtnCaption)
'QLib.Std.MIde_Cmd_MovMth.Sub Ens_CmdBarBtn(CmdBarNm, BtnCaption)
'QLib.Std.MIde_Cmd_MovMth.Sub Ens_CmdBar(A$)
'QLib.Std.MIde_Cmd_MovMth.Sub AddCmdBar(A)
'QLib.Std.MIde_Cmd_MovMth.Private Property Get XMov_MthBar() As Office.CommandBar
'QLib.Std.MIde_Cmd_MovMth.Private Property Get XMov_MthBtn() As Office.CommandBarControl
'QLib.Std.MIde_Cmd_MovMth.Private Sub Z()
'QLib.Std.MIde_Cmp.Function Cmp(CmpNm$) As VBComponent
'QLib.Std.MIde_Cmp.Function PjzCmp(A As VBComponent) As VBProject
'QLib.Std.MIde_Cmp.Function HasCmpzPj(A As VBProject, CmpNm) As Boolean
'QLib.Std.MIde_Cmp.Function PjNmzCmp$(A As VBComponent)
'QLib.Std.MIde_Cmp.Property Get CurCmp() As VBComponent
'QLib.Std.MIde_Cmp.Function CvCmp(A) As VBComponent
'QLib.Std.MIde_Cmp.Private Function HasCmpzPjTy(A As VBProject, Nm, Ty As vbext_ComponentType) As Boolean
'QLib.Std.MIde_Cmp.Function MdAyzCmp(A() As VBComponent) As CodeModule()
'QLib.Std.MIde_Cmp_is.Function IsModCmp(A As VBComponent) As Boolean
'QLib.Std.MIde_Cmp_is.Function IsClsCmp(A As VBComponent) As Boolean
'QLib.Std.MIde_Cmp_is.Function IsMd(A As VBComponent) As Boolean
'QLib.Std.MIde_Cmp_Itr.Function ClsAyPj(A As VBProject, Optional WhStr$) As CodeModule()
'QLib.Std.MIde_Cmp_Itr.Function ClsNyPj(A As VBProject) As String()
'QLib.Std.MIde_Cmp_Itr.Private Sub Z_CmpAyzPj()
'QLib.Std.MIde_Cmp_Itr.Function CmpAyzPj(A As VBProject, Optional WhStr$) As VBComponent()
'QLib.Std.MIde_Cmp_Itr.Function MdAyzCmp(CmpAy() As VBComponent) As CodeModule()
'QLib.Std.MIde_Cmp_Itr.Function IsNoClsNoModPj(A As VBProject) As Boolean
'QLib.Std.MIde_Cmp_Itr.Function ModItrzPj(A As VBProject, Optional WhStr$)
'QLib.Std.MIde_Cmp_Itr.Function ModAyzPj(A As VBProject, Optional WhStr$) As CodeModule()
'QLib.Std.MIde_Cmp_Itr.Function MdNyOfPj(Optional WhStr$) As String()
'QLib.Std.MIde_Cmp_Itr.Function MdNyWiPrpOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Cmp_Itr.Function MdNyWiPrpzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Cmp_Itr.Function IsMdWiPrp(MdNm) As Boolean
'QLib.Std.MIde_Cmp_Itr.Function MdNyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Cmp_Itr.Function MdNyzMth(MthNm$) As String()
'QLib.Std.MIde_Cmp_Itr.Function MdNsetzMth(MthNm$) As Aset
'QLib.Std.MIde_Cmp_Itr.Function MdNyzPj(A As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Cmp_Itr.Function MdNyzVbe(A As Vbe, WhStr$) As String()
'QLib.Std.MIde_Cmp_Itr.Function ModNy(Optional WhStr$) As String()
'QLib.Std.MIde_Cmp_Itr.Function ModNyzPj(A As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Cmp_Itr.Private Sub Z_ClsNyPj()
'QLib.Std.MIde_Cmp_Itr.Private Sub Z_MdAy()
'QLib.Std.MIde_Cmp_Itr.Private Sub Z_MdzPjNy()
'QLib.Std.MIde_Cmp_Itr.Private Sub Z()
'QLib.Std.MIde_Cmp_Itr.Function CmpAy(Optional WhStr$) As VBComponent()
'QLib.Std.MIde_Cmp_Itr.Function MdAy(Optional WhStr$) As CodeModule()
'QLib.Std.MIde_Cmp_Itr.Function CmpItr(A As VBProject, Optional WhStr$)
'QLib.Std.MIde_Cmp_Itr.Function MdItr(A As VBProject, Optional WhStr$)
'QLib.Std.MIde_Cmp_Itr.Function MdAyzPj(A As VBProject, Optional WhStr$) As CodeModule()
'QLib.Std.MIde_Cmp_Op_Add.Function AddCmpM(Nm) As VBComponent
'QLib.Std.MIde_Cmp_Op_Add.Function AddCmpC(Nm) As VBComponent
'QLib.Std.MIde_Cmp_Op_Add.Function AddCmp(Nm, Ty As vbext_ComponentType) As VBComponent
'QLib.Std.MIde_Cmp_Op_Add.Function AddCmpzPj(A As VBProject, Nm, Ty As vbext_ComponentType) As VBComponent
'QLib.Std.MIde_Cmp_Op_Add.Function AddModzPj(A As VBProject, ModNm) As CodeModule
'QLib.Std.MIde_Cmp_Op_Add.Sub AddMod(ModNN$)
'QLib.Std.MIde_Cmp_Op_Add.Function IsErDmp(Er$()) As Boolean
'QLib.Std.MIde_Cmp_Op_Add.Sub AddCls(ClsNN$) 'To CurPj
'QLib.Std.MIde_Cmp_Op_Add.Function MdApdLines(A As CodeModule, Lines) As CodeModule
'QLib.Std.MIde_Cmp_Op_Add.Sub AddFun(FunNm$)
'QLib.Std.MIde_Cmp_Op_Add.Function CmpNew(Nm$, Ty As vbext_ComponentType) As VBComponent
'QLib.Std.MIde_Cmp_Op_Add.Function EmpFunLines$(FunNm)
'QLib.Std.MIde_Cmp_Op_Add.Function EmpSubLines$(SubNm)
'QLib.Std.MIde_Cmp_Op_Add.Sub AddSub(SubNm$)
'QLib.Std.MIde_Cmp_Op_Add.Function AddOptExpLinMd(A As CodeModule) As CodeModule
'QLib.Std.MIde_Cmp_Op_Add.Function HasCmp(CmpNm) As Boolean
'QLib.Std.MIde_Cmp_Op_Add.Function AddCmpzLines(Nm, Lines$) As VBComponent
'QLib.Std.MIde_Cmp_Op_Add.Function AddCmpzSrcLineszPj(A As VBProject, Nm, Lines$) As VBComponent
'QLib.Std.MIde_Cmp_Op_Add.Sub RenAddCmpPfx_CmpPfx(A As VBComponent, AddPfx$)
'QLib.Std.MIde_Cmp_Op_Add.Function ModCmpItr(Pj As VBProject)
'QLib.Std.MIde_Cmp_Op_Add.Function ModCmpAy(Pj As VBProject) As VBComponent()
'QLib.Std.MIde_Cmp_Op_Add.Sub RenCmpRplPfx(A As VBComponent, FmPfx$, ToPfx$)
'QLib.Std.MIde_Cmp_Op_Add.Sub CrtPjNmzCmpTy(A As VBProject, Nm, Ty As vbext_ComponentType)
'QLib.Std.MIde_Cmp_Op_Add.Sub CrtCls(A As VBProject, Nm$)
'QLib.Std.MIde_Cmp_Op_Add.Function EnsCls(A As VBProject, ClsNm$) As CodeModule
'QLib.Std.MIde_Cmp_Op_Add.Function EnsCmp(A As VBProject, Nm, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As CodeModule
'QLib.Std.MIde_Cmp_Op_Add.Function EnsMdzPj(A As VBProject, MdNm) As CodeModule
'QLib.Std.MIde_Cmp_Op_Add.Function EnsMod(A As VBProject, ModNm$) As CodeModule
'QLib.Std.MIde_Cmp_Op_Add.Private Sub ZZ()
'QLib.Std.MIde_Cmp_Op_Add.Private Sub Z()
'QLib.Std.MIde_Cmp_Op_Cpy.Sub ThwNotCls(A As CodeModule, Fun$)
'QLib.Std.MIde_Cmp_Op_Cpy.Private Sub CpyCls(A As CodeModule, ToPj As VBProject)
'QLib.Std.MIde_Cmp_Op_Cpy.Sub CpyModAyToPj(ModAy() As CodeModule, ToPj As VBProject)
'QLib.Std.MIde_Cmp_Op_Cpy.Sub CpyClsAyToPj(ClsAy() As CodeModule, ToPj As VBProject)
'QLib.Std.MIde_Cmp_Op_Cpy.Sub CpyCmp(A As VBComponent, ToPj As VBProject)
'QLib.Std.MIde_Cmp_Op_Cpy.Sub ThwNotMod(A As CodeModule, Fun$)
'QLib.Std.MIde_Cmp_Op_Cpy.Sub CpyMod(A As CodeModule, ToPj As VBProject)
'QLib.Std.MIde_Cmp_Op_Cpy.Private Sub ZZ()
'QLib.Std.MIde_Cmp_Op_Cpy.Private Sub Z()
'QLib.Std.MIde_Cmp_Op_Ren.Sub RmvModPfx(Pj As VBProject, Pfx$)
'QLib.Std.MIde_Cmp_Op_Ren.Sub RplModPfx(FmPfx$, ToPfx$)
'QLib.Std.MIde_Cmp_Op_Ren.Sub RenCmp(A As VBComponent, NewNm$)
'QLib.Std.MIde_Cmp_Op_Ren.Sub RplModPfxzPj(Pj As VBProject, FmPfx$, ToPfx$)
'QLib.Std.MIde_Cmp_Op_Ren.Sub AddCmpSfxPj(Sfx)
'QLib.Std.MIde_Cmp_Op_Ren.Sub AddCmpSfx(A As VBProject, Sfx)
'QLib.Std.MIde_Cmp_Op_Ren.Function SetCmpNm(A As VBComponent, Nm, Optional Fun$ = "SetCmpNm") As VBComponent
'QLib.Std.MIde_Cmp_Op_Rmv.Sub DltCmpz(A As VBProject, MdNm$)
'QLib.Std.MIde_Cmp_Op_Rmv.Sub RmvMdzPfx(Pfx$)
'QLib.Std.MIde_Cmp_Op_Rmv.Sub RmvMd(A As CodeModule)
'QLib.Std.MIde_Cmp_Op_Rmv.Sub RmvCmp(A As VBComponent)
'QLib.Std.MIde_Cnt_Cmp.Property Get NCls%()
'QLib.Std.MIde_Cnt_Cmp.Property Get NCmpPj%()
'QLib.Std.MIde_Cnt_Cmp.Property Get NModPj%()
'QLib.Std.MIde_Cnt_Cmp.Function LockedCmpCnt() As CmpCnt
'QLib.Std.MIde_Cnt_Cmp.Function CmpCnt(A As VBProject) As CmpCnt
'QLib.Std.MIde_Cnt_Cmp.Property Get CmpCntPj() As CmpCnt
'QLib.Std.MIde_Cnt_Cmp.Function NCmpzPj%(A As VBProject)
'QLib.Std.MIde_Cnt_Cmp.Function NModzPj%(Pj As VBProject)
'QLib.Std.MIde_Cnt_Cmp.Function NDocOfPj%(A As VBProject)
'QLib.Std.MIde_Cnt_Cmp.Function NClszPj%(A As VBProject)
'QLib.Std.MIde_Cnt_Cmp.Function NCmpzTy%(A As VBProject, Ty As vbext_ComponentType)
'QLib.Std.MIde_Cnt_Cmp.Function NOthCmpzPj%(A As VBProject)
'QLib.Std.MIde_Cnt_Mth.Function NMthzSrc%(Src$())
'QLib.Std.MIde_Cnt_Mth.Function NMthPj%()
'QLib.Std.MIde_Cnt_Mth.Function NMthMd%()
'QLib.Std.MIde_Cnt_Mth.Function NMthzPj%(Pj As VBProject)
'QLib.Std.MIde_Cnt_Mth.Function MthDotCmlNyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Cnt_Mth.Private Function MthDotCmlNyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Cnt_Mth.Function MthCmlGpAsetOfVbe(Optional WhStr$) As Aset
'QLib.Std.MIde_Cnt_Mth.Function MthCmlGpAsetzVbe(A As Vbe, Optional WhStr$) As Aset
'QLib.Std.MIde_Cnt_Mth.Function MthCmlAsetzPj(A As VBProject, Optional WhStr$) As Aset
'QLib.Std.MIde_Cnt_Mth.Function MthCnt(A As CodeModule) As MthCnt
'QLib.Std.MIde_Cnt_Mth.Function MthCntMd() As MthCnt
'QLib.Std.MIde_Cnt_Mth.Sub MthCntPjBrw()
'QLib.Std.MIde_Cnt_Mth.Function MthCntPj() As MthCnt()
'QLib.Std.MIde_Cnt_Mth.Function LyzMthCntAy(A() As MthCnt) As String()
'QLib.Std.MIde_Cnt_Mth.Function CvMthCnt(A) As MthCnt
'QLib.Std.MIde_Cnt_Mth.Function MthCntAy(A As VBProject) As MthCnt()
'QLib.Std.MIde_Cnt_SrcLin.Property Get NSrcLin&()
'QLib.Std.MIde_Cnt_SrcLin.Function NSrcLinzPj&(A As VBProject)
'QLib.Std.MIde_ConstMth.Sub EdtConst(ConstQNm$)
'QLib.Std.MIde_ConstMth.Sub UpdConst(ConstQNm$, Optional IsPub As Boolean)
'QLib.Std.MIde_ConstMth.Private Property Get Z_CrtSchm1() As String()
'QLib.Std.MIde_ConstMth.Private Property Get C_A() As String()
'QLib.Std.MIde_ConstMth_MthLines.Function ConstPrpLines$(ConstQNm$, IsPub As Boolean)
'QLib.Std.MIde_ConstMth_MthLines.Private Function ConstPrpLy(ConstQNm$, IsPub As Boolean) As String() 'Ret Ly from ConstPrpFt
'QLib.Std.MIde_ConstMth_MthLines.Private Sub Z_ExprLyzStr()
'QLib.Std.MIde_ConstMth_MthLines.Private Function CdLyzPushStr(S, ByVal Fst As Boolean) As String()
'QLib.Std.MIde_ConstMth_MthLines.Private Sub Z_ConstPrpLines()
'QLib.Std.MIde_ConstMth_MthLines.Private Sub ZZ()
'QLib.Std.MIde_ConstMth_MthLines.Private Sub Z()
'QLib.Std.MIde_ConstMth_MthLines.Private Property Get C_A$()
'QLib.Std.MIde_ConstMth_Val.Function ConstValOfFt(ConstNm$)
'QLib.Std.MIde_ConstMth_Val.Function ConstVal$(ConstQNm$)
'QLib.Std.MIde_ConstMth_Val.Private Sub AsgMdAndConstNm(OMd As CodeModule, OConstNm$, ConstQNm$)
'QLib.Std.MIde_ConstMth_Val.Function ConstValOfMd$(Md As CodeModule, ConstNm$)
'QLib.Std.MIde_ConstMth_Val.Private Function IsConstPrp(MthLines$) As Boolean
'QLib.Std.MIde_ConstMth_Val.Function ConstValOfMth$(MthLines$)
'QLib.Std.MIde_ConstMth_Val.Private Function ConstValOfConst$(C)
'QLib.Std.MIde_ConstMth_Val.Private Function ConstLinesAy(ConstPrpLines$) As String()
'QLib.Std.MIde_ConstMth_Val.Private Sub Z_ConstValOfMth()
'QLib.Std.MIde_ConstMth_Val.Private Sub Z()
'QLib.Std.MIde_ConstMth_Val.Function ConstValOfMd1$(A As CodeModule, ConstNm$)
'QLib.Std.MIde_ConstMth_Val.Function ConstValOfLinNm$(Lin, ConstNm)
'QLib.Std.MIde_ContLin.Function ContLinzMdLno$(A As CodeModule, Lno)
'QLib.Std.MIde_ContLin.Function NxtSrcIx&(Src$(), Ix&)
'QLib.Std.MIde_ContLin.Private Sub Z_ContLin()
'QLib.Std.MIde_ContLin.Function ContLinCnt%(Src$(), Ix)
'QLib.Std.MIde_ContLin.Private Function JnContLin$(ContLy$())
'QLib.Std.MIde_ContLin.Function ContLin$(A$(), Ix, Optional OneLin As Boolean)
'QLib.Std.MIde_ContLin.Function ContFTIxzSrc(Src$(), Ix) As FTIx
'QLib.Std.MIde_ContLin.Function ContFTIxzMd(A As CodeModule, Lno&) As FTIx
'QLib.Std.MIde_ContLin.Private Function ContToLno&(A As CodeModule, Lno&)
'QLib.Std.MIde_ContLin.Function ContLinzMd$(A As CodeModule, Lno&)
'QLib.Std.MIde_Cur_CdPne_Md_Mth.Property Get CurMdNm$()
'QLib.Std.MIde_Cur_CdPne_Md_Mth.Property Get CurLno&()
'QLib.Std.MIde_Cur_CdPne_Md_Mth.Property Get CurMthNm$()
'QLib.Std.MIde_Cur_CdPne_Md_Mth.Function CurMthNmMd$(A As CodeModule)
'QLib.Std.MIde_Cur_CdPne_Md_Mth.Property Get CurWinzMd() As Vbide.Window
'QLib.Std.MIde_Cur_CdPne_Md_Mth.Property Get CurCdPne() As Vbide.CodePane
'QLib.Std.MIde_Cur_CdPne_Md_Mth.Property Get CurMthLin$()
'QLib.Std.MIde_Dcl_Const.Function ShfConst(O) As Boolean
'QLib.Std.MIde_Dcl_Const.Function HasConstNm(A As CodeModule, ConstNm$) As Boolean
'QLib.Std.MIde_Dcl_Const.Function ConstNmzSrcLin$(SrcLin)
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmBdyLyzSrc(Src$(), EnmNm$) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmBdyLy(EnmLy$()) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmFTIx(Src$(), EnmNm) As FTIx
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmLy(Src$(), EnmNm$) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmFmIx&(Src$(), EnmNm)
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmNyMd(A As CodeModule) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmNyPj(Pj As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmNy(Src$()) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function HasUsrTyNm(Src$(), Nm$) As Boolean
'QLib.Std.MIde_Dcl_EnmAndTy.Function NEnm%(Src$())
'QLib.Std.MIde_Dcl_EnmAndTy.Function UsrTyFTIx(Src$(), TyNm$) As FTIx
'QLib.Std.MIde_Dcl_EnmAndTy.Function EndEnmIx&(Src$(), FmIx)
'QLib.Std.MIde_Dcl_EnmAndTy.Function EndTyIx&(Src$(), FmIx)
'QLib.Std.MIde_Dcl_EnmAndTy.Function UsrTyLines$(Src$(), UsrTyNm$)
'QLib.Std.MIde_Dcl_EnmAndTy.Function UsrTyLy(Src$(), TyNm$) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function UsrTyFmIx&(Src$(), TyNm)
'QLib.Std.MIde_Dcl_EnmAndTy.Function UsrTyNy(Src$()) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function IsEmnLin(A) As Boolean
'QLib.Std.MIde_Dcl_EnmAndTy.Function IsUsrTyLin(A) As Boolean
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmNm$(Lin)
'QLib.Std.MIde_Dcl_EnmAndTy.Function UsrTyNm$(Lin)
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmLyMd(Md As CodeModule, EnmNm$) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function NEnmMbrMd%(A As CodeModule, EnmNm$)
'QLib.Std.MIde_Dcl_EnmAndTy.Function EnmMbrLyMd(A As CodeModule, EnmNm$) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function NEnmMd%(A As CodeModule)
'QLib.Std.MIde_Dcl_EnmAndTy.Function UsrTyNyMd(A As CodeModule) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function UsrTyNyPj(A As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Dcl_EnmAndTy.Function ShfXEnm(O) As Boolean
'QLib.Std.MIde_Dcl_EnmAndTy.Function ShfXTy(O) As Boolean
'QLib.Std.MIde_Dcl_EnmAndTy.Private Sub Z()
'QLib.Std.MIde_Dcl_EnmAndTy.Private Sub Z_NEnmMbrMd()
'QLib.Std.MIde_Dcl_Lines.Private Sub Z_DclLinCnt()
'QLib.Std.MIde_Dcl_Lines.Sub BrwDclLinCntDryPj()
'QLib.Std.MIde_Dcl_Lines.Function DclLinCntDryzPj(A As VBProject) As Variant()
'QLib.Std.MIde_Dcl_Lines.Function DclLinCntzMd%(Md As CodeModule) 'Assume FstMth cannot have TopRmk
'QLib.Std.MIde_Dcl_Lines.Function DclLinCnt%(Src$()) 'Assume FstMth cannot have TopRmk
'QLib.Std.MIde_Dcl_Lines.Function Dcl$(Src$())
'QLib.Std.MIde_Dcl_Lines.Function DclDicOfPj() As Dictionary
'QLib.Std.MIde_Dcl_Lines.Function DclDiczPj(A As VBProject) As Dictionary
'QLib.Std.MIde_Dcl_Lines.Function DclLy(Src$()) As String()
'QLib.Std.MIde_Dcl_Lines.Function DclzMd$(A As CodeModule)
'QLib.Std.MIde_Dcl_Lines.Function DclLyzMd(A As CodeModule) As String()
'QLib.Std.MIde_Dcl_Lines.Private Sub Z()
'QLib.Std.MIde_Dft.Function DftMd(A As CodeModule) As CodeModule
'QLib.Std.MIde_Dft.Function DftPj(A As VBProject) As VBProject
'QLib.Std.MIde_Dft.Function SizPj&(A As VBProject)
'QLib.Std.MIde_Dft.Function SiOfPj&()
'QLib.Std.MIde_Dft.Private Sub Z()
'QLib.Std.MIde_EnmAndTy.Function CvLinPos(A) As LinPos
'QLib.Std.MIde_EnmAndTy.Function LinPos(Lno, Optional Cno1 = 0, Optional Cno2 = 0) As LinPos
'QLib.Std.MIde_EnmAndTy.Function SubStrPos(S, SubStr) As Pos
'QLib.Std.MIde_EnmAndTy.Function Pos(Optional Cno1 = 0, Optional Cno2 = 0) As Pos
'QLib.Std.MIde_EnmAndTy.Function MdPosStr$(A As MdPos)
'QLib.Std.MIde_EnmAndTy.Function MdPos(Md As CodeModule, Pos As LinPos) As MdPos
'QLib.Std.MIde_Ens_CLib.Sub EnsCLib(A As CodeModule, Optional B As eLibNmTy = eeByDic)
'QLib.Std.MIde_Ens_CLib.Private Function LinzEptCLib$(A As CodeModule, Optional B As eLibNmTy = eeByDic)
'QLib.Std.MIde_Ens_CLib.Private Function LnxzActCLibOpt(A As CodeModule) As Lnx
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Function NoLibMdNy() As String()
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Function LibNm$(A As VBComponent, B As eLibNmTy)
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Function MdNmToLibNmDic() As Dictionary
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Private Function MdNmToLibNmDiczPj(A As VBProject) As Dictionary
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Private Function MdPfxToLibNmDic() As Dictionary
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Private Function MdNmToLibNmDiczDef() As Dictionary
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Function C_MdNmToLibNmLy() As String()
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Function C_MdPfxToLibNmLy() As String()
'QLib.Std.MIde_Ens_CLib_Cmp_LibNm11.Function LibDef() As String()
'QLib.Std.MIde_Ens_CLib__LibNm2.Sub EnsClsLibNmPj(Pj As VBProject)
'QLib.Std.MIde_Ens_CLib__LibNm2.Sub EnsClsLibNm()
'QLib.Std.MIde_Ens_CLib__LibNm2.Function HasClsLibLin(A As CodeModule) As Boolean
'QLib.Std.MIde_Ens_CLib__LibNm2.Function ClsAyPjLibNm(Pj As VBProject, LibNm) As CodeModule()
'QLib.Std.MIde_Ens_CLib__LibNm2.Sub EnsClsLibNmCls(A As CodeModule)
'QLib.Std.MIde_Ens_CLib__LibNm2.Function FstIxzAftOpt&(Src$())
'QLib.Std.MIde_Ens_CLib__LibNm2.Function FstLnozAftOptMd%(A As CodeModule)
'QLib.Std.MIde_Ens_CLib__LibNm2.Private Function ClsLibNmLin$(LibNm$)
'QLib.Std.MIde_Ens_CLib__LibNm2.Private Function HasClsLibNmLinMd(A As CodeModule) As Boolean
'QLib.Std.MIde_Ens_CSub.Function ActMdAy01zEnsCSub(A As CodeModule) As ActMd()
'QLib.Std.MIde_Ens_CSub.Function CvActLin(A) As ActLin
'QLib.Std.MIde_Ens_CSub.Function CvActMd(A) As ActMd
'QLib.Std.MIde_Ens_CSub.Function LyzActMdAy(A() As ActMd) As String()
'QLib.Std.MIde_Ens_CSub.Sub EnsCSub()
'QLib.Std.MIde_Ens_CSub.Sub EnsCSubMd()
'QLib.Std.MIde_Ens_CSub.Sub EnsCSubPj(Optional Silent As Boolean)
'QLib.Std.MIde_Ens_CSub.Sub EnsCSubzMd(A As CodeModule, Optional Silent As Boolean)
'QLib.Std.MIde_Ens_CSub.Sub EnsCSubzPj(A As VBProject, Optional Silent As Boolean)
'QLib.Std.MIde_Ens_CSub.Private Sub Z_ActMdAyzPj()
'QLib.Std.MIde_Ens_CSub.Private Function ActLin(Act As eActLin, Lno&, Lin$) As ActLin
'QLib.Std.MIde_Ens_CSub.Private Function ActMdAyzPj(Pj As VBProject) As ActMd()
'QLib.Std.MIde_Ens_CSub.Private Function CModAct(A As CModInf) As ActLin()
'QLib.Std.MIde_Ens_CSub.Private Function CModInf(A As CodeModule, B() As CSubInf) As CModInf
'QLib.Std.MIde_Ens_CSub.Private Function CModLnx(Md As CodeModule) As Lnx
'QLib.Std.MIde_Ens_CSub.Private Function CSubAct(A() As CSubInf) As ActLin()
'QLib.Std.MIde_Ens_CSub.Private Function CSubActzSng(A As CSubInf) As ActLin()
'QLib.Std.MIde_Ens_CSub.Private Function CSubInf(Src$(), B As MthRg) As CSubInf
'QLib.Std.MIde_Ens_CSub.Private Function CSubInfAy(Src$(), A() As MthRg) As CSubInf()
'QLib.Std.MIde_Ens_CSub.Private Function CSubInfSz%(A() As CSubInf)
'QLib.Std.MIde_Ens_CSub.Private Function CSubInfUB%(A() As CSubInf)
'QLib.Std.MIde_Ens_CSub.Private Function CSubLin$(MthNm$)
'QLib.Std.MIde_Ens_CSub.Private Function CSubLnx(Src$(), FmIx&, ToIx&) As Lnx
'QLib.Std.MIde_Ens_CSub.Private Function InsLnoOfCMod&(A As CodeModule)
'QLib.Std.MIde_Ens_CSub.Private Function InsLnoOfCSub&(Src$(), A As MthRg)
'QLib.Std.MIde_Ens_CSub.Private Function IsUsingCMod(A() As CSubInf) As Boolean
'QLib.Std.MIde_Ens_CSub.Private Function IsUsingCSub(Src$(), A As MthRg) As Boolean
'QLib.Std.MIde_Ens_CSub.Private Sub Z_ActMdyAy01zEnsCSub()
'QLib.Std.MIde_Ens_MthMdy.Function MthLinzEnsprv$(MthLin)
'QLib.Std.MIde_Ens_MthMdy.Function MthLinzEnspub$(MthLin)
'QLib.Std.MIde_Ens_MthMdy.Sub EnsMdPrvZ()
'QLib.Std.MIde_Ens_MthMdy.Sub EnsPjPrvZ()
'QLib.Std.MIde_Ens_MthMdy.Sub EnsPrvZzPj(A As VBProject)
'QLib.Std.MIde_Ens_MthMdy.Sub EnsPubMd()
'QLib.Std.MIde_Ens_MthMdy.Sub EnsPrvZzMd(A As CodeModule)
'QLib.Std.MIde_Ens_MthMdy.Function ActLinAyOfEnsPrvZ(A As CodeModule) As ActLin()
'QLib.Std.MIde_Ens_MthMdy.Function LnoAyOfPubZ(A As CodeModule) As Long()
'QLib.Std.MIde_Ens_MthMdy.Function LnoItrOfPubZ(A As CodeModule)
'QLib.Std.MIde_Ens_MthMdy.Sub EnsPubzMd(A As CodeModule)
'QLib.Std.MIde_Ens_MthMdy.Function ActLinAyOfEnsPubZ(A As CodeModule) As ActLin()
'QLib.Std.MIde_Ens_MthMdy.Function LnoItrPrvZ(A As CodeModule)
'QLib.Std.MIde_Ens_MthMdy.Function ActLinzEnsMdy(A As CodeModule, MthNm, Mdy) As ActLin
'QLib.Std.MIde_Ens_MthMdy.Sub EnsMdy(A As CodeModule, MthNm$, Optional Mdy$)
'QLib.Std.MIde_Ens_MthMdy.Sub EnsPrv(A As CodeModule, MthNm$)
'QLib.Std.MIde_Ens_MthMdy.Function ActLinzEnspub(A, MthNm) As ActLin
'QLib.Std.MIde_Ens_MthMdy.Sub EnsPub(A As CodeModule, MthNm$)
'QLib.Std.MIde_Ens_MthMdy.Function ActLinzEnsprv(A As CodeModule, MthNm) As ActLin
'QLib.Std.MIde_Ens_MthMdy.Private Function MthLinzEnsMdy$(OldMthLin$, ShtMdy$)
'QLib.Std.MIde_Ens_MthMdy.Private Sub Z_EnsMdy()
'QLib.Std.MIde_Ens_MthMdy.Private Sub Z()
'QLib.Std.MIde_Ens_MthMdy_PubZ_Get.Sub BrwMdNyzWiPubZ()
'QLib.Std.MIde_Ens_MthMdy_PubZ_Get.Function MdNyzWiPubZ() As String()
'QLib.Std.MIde_Ens_MthMdy_PubZ_Get.Function MdNyzWiPubZPj(A As VBProject) As String()
'QLib.Std.MIde_Ens_MthMdy_PubZ_Get.Function MthLinAyzPubZInMd(A As CodeModule) As String()
'QLib.Std.MIde_Ens_MthMdy_PubZ_Get.Function MthLinAyzPubZ() As String()
'QLib.Std.MIde_Ens_MthMdy_PubZ_Get.Function MthLinAyzPubZInPj(A As VBProject) As String()
'QLib.Std.MIde_Ens_MthMdy_PubZ_Get.Function IsWiPubZMd(A As CodeModule) As Boolean
'QLib.Std.MIde_Ens_MthMdy_PubZ_Get.Function MthLinAyzPub(Src$()) As String()
'QLib.Std.MIde_Ens_OptLin.Sub EnsOptLinPj()
'QLib.Std.MIde_Ens_OptLin.Sub EnsMdOptLin()
'QLib.Std.MIde_Ens_OptLin.Private Sub EnsOptLinzPj(Pj As VBProject)
'QLib.Std.MIde_Ens_OptLin.Private Sub EnsOptLinzMd(A As CodeModule)
'QLib.Std.MIde_Ens_OptLin.Private Sub EnsCLibzPj(A As VBProject, Optional B As eLibNmTy)
'QLib.Std.MIde_Ens_OptLin.Function LnozAftOpt%(A As CodeModule)
'QLib.Std.MIde_Ens_OptLin.Private Function OptLno%(A As CodeModule, OptLin$)
'QLib.Std.MIde_Ens_OptLin.Private Sub EnsOptLin(A As CodeModule, OptLin$)
'QLib.Std.MIde_Ens_OptLin.Private Sub RmvOptLin(A As CodeModule, OptLin$)
'QLib.Std.MIde_Ens_PrpEr.Private Sub EnsLinzExit(A As CodeModule, PrpLno&)
'QLib.Std.MIde_Ens_PrpEr.Private Sub EnsLinzLblX(A As CodeModule, PrpLno&)
'QLib.Std.MIde_Ens_PrpEr.Private Sub EnsPrpOnErzLno(A As CodeModule, PrpLno&)
'QLib.Std.MIde_Ens_PrpEr.Private Sub EnsLinzOnEr(A As CodeModule, PrpLno&)
'QLib.Std.MIde_Ens_PrpEr.Private Function LnozExit&(A As CodeModule, PrpLno)
'QLib.Std.MIde_Ens_PrpEr.Private Function LnozInsExit&(A As CodeModule, PrpLno)
'QLib.Std.MIde_Ens_PrpEr.Private Function LinzLblX$(A As CodeModule, PrpLno)
'QLib.Std.MIde_Ens_PrpEr.Private Function LnozLblX&(A As CodeModule, PrpLno)
'QLib.Std.MIde_Ens_PrpEr.Private Function IxOfOnEr&(A As CodeModule, PrpLno)
'QLib.Std.MIde_Ens_PrpEr.Private Function LnozEndPrp&(A As CodeModule, PrpLno)
'QLib.Std.MIde_Ens_PrpEr.Private Sub EnsPrpOnErzMd(A As CodeModule)
'QLib.Std.MIde_Ens_PrpEr.Private Sub RmvPrpOnErzLno(A As CodeModule, PrpLno&)
'QLib.Std.MIde_Ens_PrpEr.Private Sub RmvPrpOnErzMd(A As CodeModule)
'QLib.Std.MIde_Ens_PrpEr.Sub RmvPrpOnErOfMd()
'QLib.Std.MIde_Ens_PrpEr.Sub EnsPrpOnErOfMd()
'QLib.Std.MIde_Ens_PrpEr.Private Sub Z_EnsPrpOnErzMd()
'QLib.Std.MIde_Ens_SubZ.Private Function SubZEptzNy$(MthNySubZDash$()) ' Sub Z() bodylines
'QLib.Std.MIde_Ens_SubZ.Function SubZEptzMd$(A As CodeModule)
'QLib.Std.MIde_Ens_SubZZ.Function SubZZEpt$(A As CodeModule) ' SubZZ is Sub ZZ() bodyLines
'QLib.Std.MIde_Ens_SubZZ.Function ArgAyzPmAy(PmAy$()) As String()
'QLib.Std.MIde_Ens_SubZZ.Private Function ArgSfxAy(ArgAy$()) As String()
'QLib.Std.MIde_Ens_SubZZ.Private Function WCallingLin$(MthNm, CallingPm$, PrpGetAset As Aset)
'QLib.Std.MIde_Ens_SubZZ.Private Function WCallingLy(MthNy$(), PmAy$(), ArgDic As Dictionary, PrpGetAset As Aset) As String()
'QLib.Std.MIde_Ens_SubZZ.Private Function WCallingPm$(Pm, ArgDic As Dictionary)
'QLib.Std.MIde_Ens_SubZZ.Private Function WCallingPmAy(PmAy$(), ArgDic As Dictionary) As String()
'QLib.Std.MIde_Ens_SubZZ.Private Function WDimLy(ArgDic As Dictionary, HasPrp As Boolean) As String()  '1-Arg => 1-DimLin
'QLib.Std.MIde_Ens_SubZZ.Private Function WPrpGetAset(MthDclAy$()) As Aset
'QLib.Std.MIde_Ens_SubZZ.Private Sub Z_SubZZEpt()
'QLib.Std.MIde_Ens_SubZZ.Private Sub ZZ()
'QLib.Std.MIde_Ens_SubZZ.Private Sub Z()
'QLib.Std.MIde_Ens_SubZZ.Private Property Get Z_SubZZzMd__Ept2$()
'QLib.Std.MIde_Ens_SubZZ.Private Sub Z_SubZZEptzMd()
'QLib.Std.MIde_Ens_SubZZZ.Sub EnsSubZZZPj()
'QLib.Std.MIde_Ens_SubZZZ.Sub EnsSubZZZMd()
'QLib.Std.MIde_Ens_SubZZZ.Private Function SubZZZEptzMd$(A As CodeModule)
'QLib.Std.MIde_Ens_SubZZZ.Private Sub EnsSubZZZzMd(A As CodeModule)
'QLib.Std.MIde_Ens_SubZZZ.Private Sub EnsSubZZZzPj(A As VBProject)
'QLib.Std.MIde_Ens__Mdy.Sub MdyLinAy(A As CodeModule, B() As ActLin)
'QLib.Std.MIde_Ens__Mdy.Sub MdyLin(A As CodeModule, B As ActLin)
'QLib.Std.MIde_Ens__Mdy.Function SrcMdyLin(Src$(), B As ActLin) As String()
'QLib.Std.MIde_Ens__Mdy.Function PjMdy(A As VBProject, B() As ActMd, Optional Silent As Boolean) As VBProject
'QLib.Std.MIde_Ens__Mdy.Private Sub Z_SrcMdy()
'QLib.Std.MIde_Ens__Mdy.Private Sub Z_FmtEnsCSubzMd()
'QLib.Std.MIde_Ens__Mdy.Function FmtEnsCSubzMd(A As CodeModule) As String()
'QLib.Std.MIde_Ens__Mdy.Function FmtSrcMdy(Src$(), B() As ActLin) As String()
'QLib.Std.MIde_Ens__Mdy.Function SrcMdy(Src$(), B() As ActLin) As String()
'QLib.Std.MIde_Ens__Mdy.Function MdMdy(A As CodeModule, B() As ActLin, Optional Silent As Boolean) As CodeModule
'QLib.Std.MIde_Ens__Mdy.Private Sub PushActEr(O$(), Msg$, Ix, Cur As ActLin, Las As ActLin)
'QLib.Std.MIde_Ens__Mdy.Private Function ErzActLinCurLas(Ix, Cur As ActLin, Las As ActLin) As String()
'QLib.Std.MIde_Ens__Mdy.Private Function ErzActLinAy(A() As ActLin) As String()
'QLib.Std.MIde_Ens__Mdy.Private Function LyzActLinAy(A() As ActLin) As String()
'QLib.Std.MIde_Exp.Function ExpgPth$()
'QLib.Std.MIde_Exp.Sub ExpExpg()
'QLib.Std.MIde_Exp.Sub ExpPjf(Pjf, Optional Xls As Excel.Application, Optional Acs As Access.Application)
'QLib.Std.MIde_Exp.Sub Z1()
'QLib.Std.MIde_Exp.Sub ExpFb(Fb, Optional Acs As Access.Application)
'QLib.Std.MIde_Exp.Sub ExpFxa(Fxa, Optional Xls As Excel.Application)
'QLib.Std.MIde_Exp.Sub ExpPj()
'QLib.Std.MIde_Exp.Function PjExp(Pj As VBProject) As VBProject
'QLib.Std.MIde_Exp.Private Sub ExpSrc(A As VBProject)
'QLib.Std.MIde_Exp.Private Sub ExpRf(A As VBProject)
'QLib.Std.MIde_Exp.Private Sub ExpFrm(A As VBProject)
'QLib.Std.MIde_Exp.Private Sub ExpFrmzAcs(A As Access.Application, ToPth$)
'QLib.Std.MIde_Export.Sub ExpMd(A As CodeModule)
'QLib.Std.MIde_Export.Sub ExpPjRf(A As VBProject)
'QLib.Std.MIde_Export.Sub BrwPSrcp()
'QLib.Std.MIde_Export.Function SrcExtMd$(A As CodeModule)
'QLib.Std.MIde_Export.Function SrcFfnMd$(A As CodeModule)
'QLib.Std.MIde_Export.Function SrcpzPj$(A As VBProject)
'QLib.Std.MIde_Exp_SrcPth.Function SrcpPj$()
'QLib.Std.MIde_Exp_SrcPth.Function SrcpzCmp$(A As VBComponent)
'QLib.Std.MIde_Exp_SrcPth.Function SrcpzPjf$(Pjf)
'QLib.Std.MIde_Exp_SrcPth.Function SrcpzEns$(A As VBProject)
'QLib.Std.MIde_Exp_SrcPth.Function SrcpzDistPj$(DistPj As VBProject)
'QLib.Std.MIde_Exp_SrcPth.Function PthRmvFdr$(Pth)
'QLib.Std.MIde_Exp_SrcPth.Function FfnUp$(Ffn)
'QLib.Std.MIde_Exp_SrcPth.Function Srcp$(A As VBProject)
'QLib.Std.MIde_Exp_SrcPth.Function IsSrcp(Pth) As Boolean
'QLib.Std.MIde_Exp_SrcPth.Function SrcFn$(A As VBComponent)
'QLib.Std.MIde_Exp_SrcPth.Sub ThwNotSrcp(Srcp)
'QLib.Std.MIde_Exp_SrcPth.Function SrcFfn$(A As VBComponent)
'QLib.Std.MIde_Exp_SrcPth.Function IsSrcpInst(Pth) As Boolean
'QLib.Std.MIde_Exp_SrcPth.Sub ThwNotSrcpInst(Pth)
'QLib.Std.MIde_Exp_SrcPth_Pjf.Function Fxa$(FxaNm, Srcp)
'QLib.Std.MIde_Exp_SrcPth_Pjf.Function Fba$(FbaNm, Srcp)
'QLib.Std.MIde_Gen_Pjf.Function SrcRoot$(Srcp)
'QLib.Std.MIde_Gen_Pjf.Function DistPth$(Srcp)
'QLib.Std.MIde_Gen_Pjf.Function DistFba$(Srcp)
'QLib.Std.MIde_Gen_Pjf.Function DistFxa$(Srcp)
'QLib.Std.MIde_Gen_Pjf.Function DistFxazNxt$(Srcp)
'QLib.Std.MIde_Gen_Pjf.Private Sub Z_DistPjf()
'QLib.Std.MIde_Gen_Pjf_Bas.Sub LoadBas(DistPj As VBProject)
'QLib.Std.MIde_Gen_Pjf_Bas.Private Function BasFfnAy(Srcp) As String()
'QLib.Std.MIde_Gen_Pjf_Bas.Private Function IsBasFfn(Ffn) As Boolean
'QLib.Std.MIde_Gen_Pjf_Expg.Sub Z2()
'QLib.Std.MIde_Gen_Pjf_Expg.Sub GenExpg()
'QLib.Std.MIde_Gen_Pjf_Expg.Function SrcpAyzExpgzInst() As String()
'QLib.Std.MIde_Gen_Pjf_Expg.Private Sub Z_SrcpAyzExpgzInstzNoNonEmpDist()
'QLib.Std.MIde_Gen_Pjf_Expg.Private Sub Z_SrcpAyzExpgzInst()
'QLib.Std.MIde_Gen_Pjf_Expg.Function SrcpAyzExpgzInstzNoNonEmpDist() As String()
'QLib.Std.MIde_Gen_Pjf_Fba.Sub GenFba(SrcpInst, Optional Acs As Access.Application)
'QLib.Std.MIde_Gen_Pjf_Fba.Private Sub LoadFrm(A As VBProject)
'QLib.Std.MIde_Gen_Pjf_Fba.Private Sub LoadFrmzAcs(A As Access.Application, Srcp)
'QLib.Std.MIde_Gen_Pjf_Fba.Private Function FrmFfnAy(Srcp) As String()
'QLib.Std.MIde_Gen_Pjf_Fba.Private Function IsFrmFfn(Ffn) As Boolean
'QLib.Std.MIde_Gen_Pjf_Fxa.Private Sub Z_FxaCompress()
'QLib.Std.MIde_Gen_Pjf_Fxa.Function FxaCompress$(Fxa, Optional Xls As Excel.Application)
'QLib.Std.MIde_Gen_Pjf_Fxa.Function DistFxazSrcp$(Srcp, Optional Xls As Excel.Application)
'QLib.Std.MIde_Gen_Pjf_Fxa.Function PjzFxa(Fxa, A As Excel.Application) As VBProject
'QLib.Std.MIde_Gen_Pjf_Fxa.Function WbCrtNxtFxa(Fxa, A As Excel.Application) As Workbook
'QLib.Std.MIde_Gen_Rf.Property Get RffAy() As String()
'QLib.Std.MIde_Gen_Rf.Property Get FmtRfPj() As String()
'QLib.Std.MIde_Gen_Rf.Property Get RfLyPj() As String()
'QLib.Std.MIde_Gen_Rf.Function RfNyPj(A As VBProject) As String()
'QLib.Std.MIde_Gen_Rf.Property Get RfNy() As String()
'QLib.Std.MIde_Gen_Rf.Function CvRf(A) As Vbide.Reference
'QLib.Std.MIde_Gen_Rf.Sub CpyPjRfToPj(Pj As VBProject, ToPj As VBProject)
'QLib.Std.MIde_Gen_Rf.Function HasRf(Pj As VBProject, RfNm)
'QLib.Std.MIde_Gen_Rf.Function HasRfGuid(A As VBProject, RfGuid)
'QLib.Std.MIde_Gen_Rf.Function HasRff(A As VBProject, Rff) As Boolean
'QLib.Std.MIde_Gen_Rf.Sub BrwRf()
'QLib.Std.MIde_Gen_Rf.Function RffAyPj(A As VBProject) As String()
'QLib.Std.MIde_Gen_Rf.Function RfLin$(A As Vbide.Reference)
'QLib.Std.MIde_Gen_Rf.Function RffPjNm$(A As VBProject, RfNm$)
'QLib.Std.MIde_Gen_Rf.Function PjRfNy(A As VBProject) As String()
'QLib.Std.MIde_Gen_Rf.Sub RmvPjRfNN(A As VBProject, RfNN$)
'QLib.Std.MIde_Gen_Rf.Function StdRff$(StdRfNm)
'QLib.Std.MIde_Gen_Rf.Sub AddPjStdRf(A As VBProject, StdRfNm)
'QLib.Std.MIde_Gen_Rf.Function Rff$(A As Vbide.Reference)
'QLib.Std.MIde_Gen_Rf.Function RfPth$(A As Vbide.Reference)
'QLib.Std.MIde_Gen_Rf.Function RfToStr$(A As Vbide.Reference)
'QLib.Std.MIde_Gen_Rf.Private Sub ZZ()
'QLib.Std.MIde_Gen_Rf.Private Sub Z()
'QLib.Std.MIde_Gen_Rf_1.Sub AddRfzPj(DistPj As VBProject)
'QLib.Std.MIde_Gen_Rf_1.Sub AddRf(A As VBProject, RfLin)
'QLib.Std.MIde_Gen_Rf_1.Function RfFfn$(RfLin)
'QLib.Std.MIde_Gen_Rf_1.Function HasRfFfn(A As VBProject, RfFfn) As Boolean
'QLib.Std.MIde_Gen_Rf_1.Function RfSrcFfn$(A As VBProject)
'QLib.Std.MIde_Gen_Rf_1.Function RfSrcFfnzDistPj$(DistPj As VBProject)
'QLib.Std.MIde_Gen_Rf_1.Function RfSrcPj() As String()
'QLib.Std.MIde_Gen_Rf_1.Function RfSrczPj(A As VBProject) As String()
'QLib.Std.MIde_Gen_Rf_1.Function RfLin$(A As Vbide.Reference)
'QLib.Std.MIde_Gen_Rf_Add.Private Sub ZZ()
'QLib.Std.MIde_Gen_Rf_Add.Private Sub Z()
'QLib.Std.MIde_Gen_Rf_Add.Sub AddRfzRff(A As VBProject, Rff)
'QLib.Std.MIde_Gen_Rf_Add.Sub AddRfzAy(A As VBProject, RffAy$())
'QLib.Std.MIde_Gen_Rf_Dfn.Private Property Get UsrPjRfLy() As String()
'QLib.Std.MIde_Gen_Rf_Dfn.Property Get PjNmToStdRfNNDic() As Dictionary
'QLib.Std.MIde_Gen_Rf_Dfn.Private Property Get PjNmToStdRfNNLy() As String()
'QLib.Std.MIde_Gen_Rf_Dfn.Private Property Get StdGuidLy() As String()
'QLib.Std.MIde_Gen_Rf_Dfn.Private Sub Z_FAny_DPD_ORD()
'QLib.Std.MIde_Gen_Rf_Dfn.Private Sub ZZ()
'QLib.Std.MIde_Gen_Rf_Dfn.Private Sub Z()
'QLib.Std.MIde_Gen_Rf_Dfn.Sub BrwRf()
'QLib.Std.MIde_Gen_Rf_Dfn.Sub DmpRf()
'QLib.Std.MIde_Gen_Rf_Dfn.Sub DmpRfzPj(A As VBProject)
'QLib.Std.MIde_Gen_Rf_Dfn.Private Function GuidLinRfNm$(RfNm)
'QLib.Std.MIde_Gen_Rf_Dfn.Private Function GuidLinAyPjNm(PjNm$) As String()
'QLib.Std.MIde_Gen_Rf_Dfn.Private Function StdRfNyPj(PjNm$) As String()
'QLib.Std.MIde_Gen_Rf_InfDta.Private Sub Z_PjRfDrs()
'QLib.Std.MIde_Gen_Rf_InfDta.Sub DmpPjRf(A As VBProject)
'QLib.Std.MIde_Gen_Rf_InfDta.Function PjRfDrs(A As VBProject) As Drs
'QLib.Std.MIde_Gen_Rf_InfDta.Property Get PjRfFny() As String()
'QLib.Std.MIde_Gen_Rf_InfDta.Function PjRfDry(A As VBProject) As Variant()
'QLib.Std.MIde_Gen_Rf_InfDta.Function DrRf(A As Vbide.Reference) As Variant()
'QLib.Std.MIde_Gen_Rf_InfDta.Property Get RfFny() As String()
'QLib.Std.MIde_Gen_Rf_InfDta.Function PjAyRfDrs(A() As VBProject) As Drs
'QLib.Std.MIde_Hit.Function WhMthKd(S) As String()
'QLib.Std.MIde_Hit.Function WhMthMdyPm(A As LinPm) As String()
'QLib.Std.MIde_Hit.Function WhMthMdy(WhStr$) As String()
'QLib.Std.MIde_Hit.Function HitCmp(A As VBComponent, B As WhMd) As Boolean
'QLib.Std.MIde_Identifier.Private Sub Z_NyzStr()
'QLib.Std.MIde_Identifier.Private Sub Z_NsetzStr()
'QLib.Std.MIde_Identifier.Function NsetzStr(S) As Aset
'QLib.Std.MIde_Identifier.Function RplPun$(Str)
'QLib.Std.MIde_Identifier.Function SyeNonNm(Sy$()) As String()
'QLib.Std.MIde_Identifier.Function NyzStr(S) As String()
'QLib.Std.MIde_Identifier.Function RelOf_PubMthNm_To_ModNy_OfPj() As Rel
'QLib.Std.MIde_Identifier.Function RelOfMthNmToCmlOfVbe(Optional WhStr$) As Rel
'QLib.Std.MIde_Identifier.Function RelOfMthNmToCmlzVbe(A As Vbe, Optional WhStr$) As Rel
'QLib.Std.MIde_Identifier.Function RelOf_PubMthNm_To_ModNy_zPj(A As VBProject) As Rel
'QLib.Std.MIde_Identifier.Function RelOf_MthNm_To_MdNy_zPj(A As VBProject) As Rel
'QLib.Std.MIde_Identifier.Function RelOf_MthNm_To_MdNy_OfPj() As Rel
'QLib.Std.MIde_Identifier.Function MthExtNy(MthPjDotMdNm$, PubMthLy$(), PubMthNm_To_PjDotModNy As Dictionary) As String()
'QLib.Std.MIde_Identifier.Property Get VbKwAy() As String()
'QLib.Std.MIde_Identifier.Property Get VbKwAset() As Aset
'QLib.Std.MIde_Lis.Function MdzPjLisDt(A As VBProject, Optional B As WhMd) As Dt
'QLib.Std.MIde_Lis.Sub MdzPjLisBrwDt(A As VBProject, Optional B As WhMd)
'QLib.Std.MIde_Lis.Sub MdzPjLisDmpDt(A As VBProject, Optional B As WhMd)
'QLib.Std.MIde_Lis.Sub LisMd(Optional Patn$, Optional Exl$)
'QLib.Std.MIde_Lis.Sub LisPj()
'QLib.Std.MIde_Lis.Sub LisMth(Optional WhStr$)
'QLib.Std.MIde_Lis.Private Function WhStrzMthPatn$(MthPatn$, Optional PubOnly As Boolean)
'QLib.Std.MIde_Lis.Private Function WhStrzPubOnly$(PubOnly As Boolean)
'QLib.Std.MIde_Lis.Function WhStrzPfx$(MthPfx$, Optional PubOnly As Boolean)
'QLib.Std.MIde_Lis.Function WhStrzSfx$(MthSfx$, Optional PubOnly As Boolean)
'QLib.Std.MIde_Lis.Sub LisMthPfx(Pfx$, Optional PubOnly As Boolean)
'QLib.Std.MIde_Lis.Sub LisMthSfx(Sfx$, Optional PubOnly As Boolean)
'QLib.Std.MIde_Lis.Sub LisMthPatn(Patn$, Optional InclPrv As Boolean)
'QLib.Std.MIde_Loc.Function MthPos(MthNm) As MdPos()
'QLib.Std.MIde_Loc.Function LocLyPatn(Patn$) As String()
'QLib.Std.MIde_Loc.Function LocLyzPjPatn(A As VBProject, Patn$) As String()
'QLib.Std.MIde_Loc.Function CurLocLyPjRe(Re_Or_Patn) As String()
'QLib.Std.MIde_Loc.Function LocLyPjRe(A As VBProject, Re As RegExp) As String()
'QLib.Std.MIde_Loc.Function LocLyMdRe(A As CodeModule, Re As RegExp) As String()
'QLib.Std.MIde_Loc_Jmp.Sub MdJmpLno(A As CodeModule, Lno&)
'QLib.Std.MIde_Loc_Jmp.Function HasMdNm(MdNm$, Optional Inf As Boolean) As Boolean
'QLib.Std.MIde_Loc_Jmp.Sub Jmp(MdNm$)
'QLib.Std.MIde_Loc_Jmp.Function JmpMdNm%(MdNm$)
'QLib.Std.MIde_Loc_Jmp.Sub JmpMdPos(A As MdPos)
'QLib.Std.MIde_Loc_Jmp.Sub LinPosAsg(A As LinPos, OLno&, OC1%, OC2%)
'QLib.Std.MIde_Loc_Jmp.Sub JmpLin(Lno&)
'QLib.Std.MIde_Loc_Jmp.Sub JmpLinPos(A As LinPos)
'QLib.Std.MIde_Loc_Jmp.Sub JmpLno(Lno&)
'QLib.Std.MIde_Loc_Jmp.Function MdPoszLno(A As CodeModule, Lno&) As MdPos
'QLib.Std.MIde_Loc_Jmp.Function EmpPos() As Pos
'QLib.Std.MIde_Loc_Jmp.Sub JmpMdMth(A As CodeModule, MthNm$)
'QLib.Std.MIde_Loc_Jmp.Function MdPosAyzMth(MthNm$) As MdPos()
'QLib.Std.MIde_Loc_Jmp.Sub JmpMth(MthNm$)
'QLib.Std.MIde_Loc_Jmp.Function LinPosStr$(A As LinPos)
'QLib.Std.MIde_Loc_Jmp.Sub JmpCurMd()
'QLib.Std.MIde_Loc_Jmp.Sub JmpMd(A As CodeModule)
'QLib.Std.MIde_Loc_Jmp.Sub JmpPj(A As VBProject)
'QLib.Std.MIde_Loc_Jmp.Sub JmpPjMd(A As VBProject, MdNm$)
'QLib.Std.MIde_Loc_Jmp.Sub MdJmp(A As CodeModule)
'QLib.Std.MIde_Loc_Jmp.Sub JmpClsNm(ClsNm$)
'QLib.Std.MIde_Loc_Jmp.Sub JmpCmp(CmpNm$)
'QLib.Std.MIde_Md.Property Get CurMd() As CodeModule
'QLib.Std.MIde_Md.Private Sub Z_CurMd()
'QLib.Std.MIde_Md.Property Get CurMdDNm$()
'QLib.Std.MIde_Md.Function MdAyCmpAy(A() As CodeModule) As VBComponent()
'QLib.Std.MIde_Md.Function Md(MdDNm) As CodeModule
'QLib.Std.MIde_Md.Function MdAywInTy(A() As CodeModule, WhInTyAy0$) As CodeModule()
'QLib.Std.MIde_Md.Function IsMod(A As CodeModule) As Boolean
'QLib.Std.MIde_Md.Function IsCls(A As CodeModule) As Boolean
'QLib.Std.MIde_Md.Sub ClsMd(A As CodeModule)
'QLib.Std.MIde_Md.Sub CmpMdAB(A As CodeModule, B As CodeModule)
'QLib.Std.MIde_Md.Function MdQNmzMd$(A As CodeModule)
'QLib.Std.MIde_Md.Sub RmvMdLno(A As CodeModule, Lno&)
'QLib.Std.MIde_Md.Function SizMd&(A As CodeModule)
'QLib.Std.MIde_Md.Function MdNm$(A As CodeModule)
'QLib.Std.MIde_Md.Function NUsrTyMd%(A As CodeModule)
'QLib.Std.MIde_Md.Function PjzMd(A As CodeModule) As VBProject
'QLib.Std.MIde_Md.Function SrcRmvMth(Src$(), MthNmSet As Aset) As String()
'QLib.Std.MIde_Md.Function SrcLines$(A As CodeModule)
'QLib.Std.MIde_Md.Function MdDNm$(A As CodeModule)
'QLib.Std.MIde_Md.Function PjNmzMd(A As CodeModule)
'QLib.Std.MIde_Md.Function PjNmzCmp(A As VBComponent)
'QLib.Std.MIde_Md.Function MdFn$(A As CodeModule)
'QLib.Std.MIde_Md.Function MdTy(A As CodeModule) As vbext_ComponentType
'QLib.Std.MIde_Md.Private Property Get ZZMd() As CodeModule
'QLib.Std.MIde_Md.Private Sub ZZ_MdDrs()
'QLib.Std.MIde_Md.Private Sub ZZ_MthLnoMdMth()
'QLib.Std.MIde_Md.Function MdTyNm$(A As CodeModule)
'QLib.Std.MDao_Li_Mis.Function LiMis(A As LiPm, B As LiAct) As LiMis
'QLib.Std.MDao_Li_Mis.Private Function MisTbl(A As LiPm, B As LiAct) As LiMisTbl()
'QLib.Std.MDao_Li_Mis.Private Function MisTblFx(A() As LiFx, B() As LiActFx) As LiMisTbl()
'QLib.Std.MDao_Li_Mis.Private Function MisTblFb(A() As LiFb, B() As LiActFb) As LiMisTbl()
'QLib.Std.MDao_Li_Mis.Private Function MisTblFbOpt(A As LiFb, B() As LiActFb) As LiMisTbl
'QLib.Std.MDao_Li_Mis.Private Function MisTblFxOpt(A As LiFx, B() As LiActFx) As LiMisTbl
'QLib.Std.MIde_Mdy_LinShould.Function ShouldIns(IsUsing As Boolean, OldLin$, NewLin$) As Boolean
'QLib.Std.MIde_Mdy_LinShould.Function ShouldDlt(IsUsing As Boolean, OldLin$, NewLin$) As Boolean
'QLib.Std.MIde_Md_Emp.Function IsEmpMd(A As CodeModule) As Boolean
'QLib.Std.MIde_Md_Emp.Sub RmvEmpMd()
'QLib.Std.MIde_Md_Emp.Property Get EmpMdNy() As String()
'QLib.Std.MIde_Md_Emp.Function EmpMdNyzPj(A As VBProject) As String()
'QLib.Std.MIde_Md_Emp.Private Sub Z_IsEmpMd()
'QLib.Std.MIde_Md_Emp.Function IsEmpSrc(A$()) As Boolean
'QLib.Std.MIde_Md_Emp.Function IsEmpSrcLin(A) As Boolean
'QLib.Std.MIde_Md_Emp.Function EmpMdNyzVbe(A As Vbe) As String()
'QLib.Std.MIde_Md_Emp.Function IsNoMthMd(A As CodeModule) As Boolean
'QLib.Std.MIde_Md_Itr.Function ModItr()
'QLib.Std.MIde_Md_Max.Function MaxLinCntMdNm$()
'QLib.Std.MIde_Md_Max.Function MaxLinCntMdNmzPj$(A As VBProject)
'QLib.Std.MIde_Md_Max.Function MaxLinCntMd(A As VBProject) As CodeModule
'QLib.Std.MIde_Md_Max.Function CvMd(A) As CodeModule
'QLib.Std.MIde_Md_Op_Add_Lines.Function MdInsDcl(A As CodeModule, Dcl$) As CodeModule
'QLib.Std.MIde_Md_Op_Add_Lines.Function MdApdLy(A As CodeModule, Ly$()) As CodeModule
'QLib.Std.MIde_Md_Op_Ren.Sub RenTo(FmCmpNm$, ToNm$)
'QLib.Std.MIde_Md_Op_Ren.Sub Ren(NewCmpNm$)
'QLib.Std.MIde_Md_Op_Ren.Sub RenMd(A As CodeModule, NewNm$)
'QLib.Std.MIde_Md_Op_Ren.Sub MthKeyDrFny()
'QLib.Std.MIde_Md_Op_Rmk.Sub Rmk()
'QLib.Std.MIde_Md_Op_Rmk.Sub UnRmk()
'QLib.Std.MIde_Md_Op_Rmk.Sub RmkAllMd()
'QLib.Std.MIde_Md_Op_Rmk.Function IsRmkedMd(A As CodeModule) As Boolean
'QLib.Std.MIde_Md_Op_Rmk.Sub UnRmkAllMd()
'QLib.Std.MIde_Md_Op_Rmk.Private Function RmkMd(A As CodeModule) As Boolean
'QLib.Std.MIde_Md_Op_Rmk.Private Function IfUnRmkMd(A As CodeModule) As Boolean
'QLib.Std.MIde_Md_Op_Rmv_Lines.Sub ClrMd(A As CodeModule)
'QLib.Std.MIde_Md_Op_Rmv_Lines.Sub MdRmvFTIxAy(A As CodeModule, B() As FTIx)
'QLib.Std.MIde_Md_Op_Rmv_Lines.Function CntSzStrzMd$(A As CodeModule)
'QLib.Std.MIde_Md_Op_Rmv_Lines.Function MdLineszMd(A As CodeModule) As MdLines
'QLib.Std.MIde_Md_Op_Rmv_Lines.Sub MdRpl(A As CodeModule, NewMdLines$)
'QLib.Std.MIde_Md_Op_Rmv_Lines.Sub RmvMdFTIx(A As CodeModule, FTIx As FTIx)
'QLib.Std.MIde_Md_Op_Rmv_Lines.Sub RmvMdFtLinesIxAy(A As CodeModule, B() As FTIx)
'QLib.Std.MIde_Md_Op_Rmv_Lines.Private Sub Z_RmvMdFtLinesIxAy()
'QLib.Std.MIde_Md_Pfx.Sub BrwMdPfx()
'QLib.Std.MIde_Md_Pfx.Function MdPfxAyzPj(A As VBProject) As String()
'QLib.Std.MIde_Md_Pfx.Function MdPfxCntDiczPj(A As VBProject) As Dictionary
'QLib.Std.MIde_Md_Pfx.Function MdPfxCntDic() As Dictionary
'QLib.Std.MIde_Md_Pfx.Function MdPfxAy(MdNy$()) As String()
'QLib.Std.MIde_Md_Pfx.Function MdPfx$(MdNm)
'QLib.Std.MIde_Md_Res.Function ResLyMd(A As CodeModule, ResNm$, Optional ResPfx$ = "ZZRes") As String()
'QLib.Std.MIde_Md_Res.Function ReStrMd$(A As CodeModule, ResNm$)
'QLib.Std.MIde_Mod_ChgToCmp.Sub ChgToCmpz(FmModNm$)
'QLib.Std.MIde_Mod_ChgToCmp.Sub ChgToCmp()
'QLib.Std.MIde_Mth.Property Get MthKeyDrFny() As String()
'QLib.Std.MIde_Mth.Function MthDNySq(A$()) As Variant()
'QLib.Std.MIde_Mth.Function MdLinesAyzMth(A As CodeModule, MthNm) As MdLines()
'QLib.Std.MIde_Mth.Function MdRplMth(Md As CodeModule, MthNm, ByLines) As CodeModule
'QLib.Std.MIde_Mth.Private Sub Z()
'QLib.Std.MIde_Mth.Function MdEns(Md As CodeModule, MthNm$, MthLines$) As CodeModule
'QLib.Std.MIde_Mth.Private Sub Z_MthFTixAyzMth()
'QLib.Std.MIde_Mth_Cml.Function MthCmlAsetOfPj(Optional WhStr$) As Aset
'QLib.Std.MIde_Mth_Cml.Function MthCmlFny(NDryCol%) As String()
'QLib.Std.MIde_Mth_Cml.Function MthCmlWs(Optional Vis As Boolean) As Worksheet
'QLib.Std.MIde_Mth_Cml.Function MthCmlLinWsBase() As Worksheet
'QLib.Std.MIde_Mth_Cml.Sub BrwMthCmlLyOfVbe()
'QLib.Std.MIde_Mth_Cml.Function MthCmlLyOfVbe() As String()
'QLib.Std.MIde_Mth_Cml.Function MthCmlLyzVbe(A As Vbe) As String()
'QLib.Std.MIde_Mth_Cnt.Function NMthzMd%(A As CodeModule, Optional WhStr$)
'QLib.Std.MIde_Mth_Cnt.Function NSrcLinPj&(A As VBProject)
'QLib.Std.MIde_Mth_Cnt.Function NPubMthMd%(A As CodeModule)
'QLib.Std.MIde_Mth_Cnt.Function NPubMthVbe%(A As Vbe)
'QLib.Std.MIde_Mth_Cnt.Property Get NPubMth%()
'QLib.Std.MIde_Mth_Cnt.Function NPubMthPj%(A As VBProject)
'QLib.Std.MIde_Mth_Cnt.Function NMthzSrc%(A$(), Optional WhStr$)
'QLib.Std.MIde_Mth_Dcl.Property Get CurMthLinAyzMd() As String()
'QLib.Std.MIde_Mth_Dcl.Function MthLinAyzSrcNm(A$(), MthNm$) As String()
'QLib.Std.MIde_Mth_Dcl.Private Sub Z_Src_PthMthLinAy()
'QLib.Std.MIde_Mth_Dic.Private Sub ZZ_Pj_MthDic()
'QLib.Std.MIde_Mth_Dic.Function Pj_MthDic(A As VBProject) As Dictionary
'QLib.Std.MIde_Mth_Dic.Private Sub ZZ_MdMthDic()
'QLib.Std.MIde_Mth_Dic.Private Sub Z_MthDiczMd()
'QLib.Std.MIde_Mth_Dic.Private Sub Z_PjMthDic()
'QLib.Std.MIde_Mth_Dic.Private Sub Z_PjMthDic1()
'QLib.Std.MIde_Mth_Dic.Private Sub Z()
'QLib.Std.MIde_Mth_Dic.Private Sub Z_MthDic()
'QLib.Std.MIde_Mth_Dic.Function MthDicPj()
'QLib.Std.MIde_Mth_Dic.Function MthDiczPj(A As VBProject, Optional WiTopRmk As Boolean) As Dictionary
'QLib.Std.MIde_Mth_Dic.Function MthDicMd() As Dictionary
'QLib.Std.MIde_Mth_Dic.Function MthDiczMd(A As CodeModule, Optional WiTopRmk As Boolean) As Dictionary
'QLib.Std.MIde_Mth_Dic.Function MthNmDic(Src$()) As Dictionary 'Key is MthNm.  One PrpNm may have 2 PrpMth: (Get & Set) or (Get & Let)
'QLib.Std.MIde_Mth_Dic.Function MthDic(Src$(), Optional WiTopRmk As Boolean) As Dictionary 'Key is MthDNm, Val is MthLinesWiTopRmk
'QLib.Std.MIde_Mth_Drs.Function MthInfAyzVbe(A As Vbe) As MthInf()
'QLib.Std.MIde_Mth_Drs.Function MthInfAy_Pj(A As VBProject) As MthInf()
'QLib.Std.MIde_Mth_Drs.Function MthInfAy_Md(A As CodeModule) As MthInf()
'QLib.Std.MIde_Mth_Drs.Function MthInfAy_MdzPjSrc(PjNm$, MdNm$, Src$()) As MthInf()
'QLib.Std.MIde_Mth_Drs.Function PjNmMth$(A As CodeModule)
'QLib.Std.MIde_Mth_Drs.Function MthInfMdzPjSrcFm(PjNm$, MdNm$, Src$(), FmIx) As MthInf
'QLib.Std.MIde_Mth_Drs.Function MthDrs(Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthDrszMd(A As CodeModule, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthLinDryzMd(A As CodeModule, Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Drs.Function MthLinDryzSrc(Src$(), Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Drs.Function MthDryzMd(A As CodeModule, Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Drs.Property Get PjfAy() As String()
'QLib.Std.MIde_Mth_Drs.Function MthWb(Optional WhStr$) As Workbook
'QLib.Std.MIde_Mth_Drs.Property Get MthWs() As Worksheet
'QLib.Std.MIde_Mth_Drs.Function MthDrsFb(Fb, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthDrszFxa(Fxa, Optional WhStr$, Optional Xls As Excel.Application) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthWbPjfAy(PjfAy$(), Optional WhStr$) As Workbook
'QLib.Std.MIde_Mth_Drs.Function MthWsPjfAy(PjfAy, Optional WhStr$) As Worksheet
'QLib.Std.MIde_Mth_Drs.Function MthDrszPjf(Pjf, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthDrszPj(A As VBProject, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthDryzPj(A As VBProject, Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Drs.Function MthWszPj(A As VBProject, Optional WhStr$, Optional Vis As Boolean) As Worksheet
'QLib.Std.MIde_Mth_Drs.Private Sub Z_MthWb()
'QLib.Std.MIde_Mth_Drs.Private Sub Z_MthDrszMd()
'QLib.Std.MIde_Mth_Drs.Private Sub Z_MthWbFmt()
'QLib.Std.MIde_Mth_Drs.Function MthDrszPjfAy(PjfAy, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Private Sub Z_MthDRszPjf()
'QLib.Std.MIde_Mth_Drs.Function MthDrsOfVbe(Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthDryzVbe(A As Vbe, Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Drs.Function MthDrszVbe(A As Vbe, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthWszVbe(A As Vbe, Optional WhStr$) As Worksheet
'QLib.Std.MIde_Mth_Drs.Function MthWbFmt(A As Workbook) As Workbook
'QLib.Std.MIde_Mth_Drs.Function MthLinDrsOfPj(Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthLinDrszPj(A As VBProject, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs.Function MthLinDryzPj(A As VBProject, Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Drs.Property Get MthLinFny() As String()
'QLib.Std.MIde_Mth_Drs.Function MthLinDr(MthLin, Optional B As WhMth) As Variant()
'QLib.Std.MIde_Mth_Drs.Function MthDryzSrc(Src$(), Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Drs.Function MthInfSrcFm(Src$(), MthFmIx&) As Variant()
'QLib.Std.MIde_Mth_Drs.Property Get MthFny() As String()
'QLib.Std.MIde_Mth_Drs.Function VbeAyMthDrs(A() As Vbe) As Drs
'QLib.Std.MIde_Mth_Drs.Function VbeAyMthWs(A() As Vbe) As Worksheet
'QLib.Std.MIde_Mth_Drs.Private Property Get ZZVbeAy() As Vbe()
'QLib.Std.MIde_Mth_Drs.Private Sub Z_MthLinDryzPj()
'QLib.Std.MIde_Mth_Drs.Private Sub Z_VbeAyMthWs()
'QLib.Std.MIde_Mth_Drs.Private Sub Z_MthLinDryzVbe()
'QLib.Std.MIde_Mth_Drs.Private Sub ZZ()
'QLib.Std.MIde_Mth_Drs.Private Sub Z()
'QLib.Std.MIde_Mth_Drs_Cache.Function CacheDtevPjf(Pjf) As Date
'QLib.Std.MIde_Mth_Drs_Cache.Function MthDrszPjfzFmCache(Pjf, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Drs_Cache.Sub EnsTofMthCachezPjf(Pjf)
'QLib.Std.MIde_Mth_Drs_Cache.Sub ThwIfDrsGoodToIupDbt(Drs As Drs, Db As Database, T)
'QLib.Std.MIde_Mth_Drs_Cache.Function BexprzFnyWiSqlQuMk(FNyWiSqlQuMk$())
'QLib.Std.MIde_Mth_Drs_Cache.Function SqlQuMk$(A As Dao.DataTypeEnum)
'QLib.Std.MIde_Mth_Drs_Cache.Function SkFnyWiSqlQuMkPfx(A As Database, T) As String()
'QLib.Std.MIde_Mth_Drs_Cache.Sub IupDbt(A As Database, T, Drs As Drs)
'QLib.Std.MIde_Mth_Drs_Cache.Sub InsDbt(A As Database, T, Dry())
'QLib.Std.MIde_Mth_Drs_Cache.Sub UpdDbt(A As Database, T, Dry())
'QLib.Std.MIde_Mth_Dup_Compare.Sub CmpFun(FunNm$, Optional InclEqLines As Boolean)
'QLib.Std.MIde_Mth_Dup_Compare.Function FmtCmpFun(FunNm, Optional InclSam As Boolean) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Private Sub Z_FunCmp()
'QLib.Std.MIde_Mth_Dup_Compare.Private Function FmtCmpDic(A, Optional OIx% = -1, Optional OSam% = -1, Optional InclSam As Boolean) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Private Function FmtCmpDic__1Hdr(OIx%, MthNm$, Cnt%) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Private Function FmtCmpDic__2Sam(InclSam As Boolean, OSam%, DupMthFNyGp, LinesAy$()) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Private Function FmtCmpDic__2Sam1(OSam%, Lines$, DupMthFNyGp, LinesAy$()) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Private Function FmtCmpDic__3Syn(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Private Function FmtCmpDic__4Cmp(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Function MthNmCmpFmt(A, Optional InclSam As Boolean) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Function VbeDupMthCmpLy(A As Vbe, B As WhPjMth, Optional InclSam As Boolean) As String()
'QLib.Std.MIde_Mth_Dup_Compare.Private Sub ZZ_VbeDupMthCmpLy()
'QLib.Std.MIde_Mth_Dup_Compare.Private Sub Z()
'QLib.Std.MIde_Mth_Fb.Function MthFbOfPj$()
'QLib.Std.MIde_Mth_Fb.Function MthFb$()
'QLib.Std.MIde_Mth_Fb.Function MthFbzPj$(A As VBProject)
'QLib.Std.MIde_Mth_Fb.Function MthPthzPj$(A As VBProject)
'QLib.Std.MIde_Mth_Fb.Function EnsMthFb(MthFb$) As Database
'QLib.Std.MIde_Mth_Fb.Function MthDbzPj(A As VBProject) As Database
'QLib.Std.MIde_Mth_Fb.Property Get MthDbOfPj() As Database
'QLib.Std.MIde_Mth_Fb.Sub BrwMthFb()
'QLib.Std.MIde_Mth_Fb.Private Property Get MthSchm() As String()
'QLib.Std.MIde_Mth_Fb_Gen.Sub CrtDistMth()
'QLib.Std.MIde_Mth_Fb_Gen.Sub CrtMdDic()
'QLib.Std.MIde_Mth_Fb_Gen.Sub UpdMthLoc()
'QLib.Std.MIde_Mth_Fb_Gen.Private Sub ZZ()
'QLib.Std.MIde_Mth_Fb_Gen.Private Sub Z()
'QLib.Std.MIde_Mth_Ix.Private Sub Z_MthIxAy()
'QLib.Std.MIde_Mth_Ix.Function MthIxItr(Src$(), Optional WhStr$)
'QLib.Std.MIde_Mth_Ix.Function EndLinIx&(Src$(), EndLinItm$, FmIx)
'QLib.Std.MIde_Mth_Ix.Function MthIx&(Src$(), MthNm$)
'QLib.Std.MIde_Mth_Ix.Function MthIxAy(Src$(), Optional WhStr$) As Long()
'QLib.Std.MIde_Mth_Ix.Function MthIxzSrcNmTy(Src$(), MthNm, ShtMthTy$) As LngRslt
'QLib.Std.MIde_Mth_Ix.Function MthIxAyzNm(Src$(), MthNm) As Long()
'QLib.Std.MIde_Mth_Ix.Function MthIxzFst&(Src$(), MthNm, Optional SrcFmIx& = 0)
'QLib.Std.MIde_Mth_Ix.Function MthToIxAy(Src$(), FmIxAy&()) As Long()
'QLib.Std.MIde_Mth_Ix.Function MthToIx&(Src$(), MthIx)
'QLib.Std.MIde_Mth_Ix.Function FstMthLnoMd&(Md As CodeModule)
'QLib.Std.MIde_Mth_Ix.Function FstMthIx&(Src$())
'QLib.Std.MIde_Mth_Ix.Function MthLnoMdMth&(A As CodeModule, MthNm)
'QLib.Std.MIde_Mth_Ix.Function MthLnoAyMdMth(A As CodeModule, MthNm) As Long()
'QLib.Std.MIde_Mth_Ix.Private Sub Z()
'QLib.Std.MIde_Mth_Ix.Function MthRgAy(Src$()) As MthRg()
'QLib.Std.MIde_Mth_Ix_FT.Function MthFTIxAyzSrcMth(Src$(), MthNm, Optional WiTopRmk As Boolean) As FTIx()
'QLib.Std.MIde_Mth_Ix_FT.Function MthFTIxAyzMth(A As CodeModule, MthNm, Optional WiTopRmk As Boolean) As FTIx()
'QLib.Std.MIde_Mth_Ix_FT.Function MthFTIxAy(Src$(), Optional WiTopRmk As Boolean) As FTIx()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLinDic(Optional WhStr$) As Dictionary
'QLib.Std.MIde_Mth_LinAy_Drs.Private Function MthLinDiczVbe(A As Vbe, Optional WhStr$) As Dictionary
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLinDiczPj(A As VBProject, Optional WhStr$) As Dictionary
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLinDiczSrc(Src$(), Optional WhStr$) As Dictionary
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLinAyzPj(A As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLinAyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLinAyzVbe(V As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLnxAyzMd(A As CodeModule, Optional WhStr$) As Lnx()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLnxAyzSrc(Src$(), Optional WhStr$) As Lnx()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLinAyzMd(A As CodeModule, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthLinAyzSrc(Src$(), Optional WhStr$) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthNmFny() As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Private Function MthQ1LyzMthQLy(MthQLy$()) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthQ1LyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthQLyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthQLyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthQLyzMd(A As CodeModule) As String()
'QLib.Std.MIde_Mth_LinAy_Drs.Function MthQLyzPj(A As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Lines.Private Property Get XX1()
'QLib.Std.MIde_Mth_Lines.Private Property Let XX1(V)
'QLib.Std.MIde_Mth_Lines.Function MthLineszPub$(PubMthNm)
'QLib.Std.MIde_Mth_Lines.Property Get MthLines$()
'QLib.Std.MIde_Mth_Lines.Sub BrwMthLinesAyPj()
'QLib.Std.MIde_Mth_Lines.Function MthLinesAyPj(Optional WiTopRmk As Boolean) As String()
'QLib.Std.MIde_Mth_Lines.Function MthLinesAyzPj(A As VBProject, Optional WiTopRmk As Boolean) As String()
'QLib.Std.MIde_Mth_Lines.Function MthLinesAyzMd(A As CodeModule, Optional WiTopRmk As Boolean) As String()
'QLib.Std.MIde_Mth_Lines.Function MthLinesAyzSrc(Src$(), Optional WiTopRmk As Boolean) As String()
'QLib.Std.MIde_Mth_Lines.Function MthLineszMd$(Md As CodeModule, MthNm, Optional WiTopRmk As Boolean)
'QLib.Std.MIde_Mth_Lines.Function MthLineszMdNmTy$(Md As CodeModule, MthNm, ShtMthTy$, Optional WiTopRmk As Boolean)
'QLib.Std.MIde_Mth_Lines.Function MthLyzMdMth(Md As CodeModule, MthNm, Optional WiTopRmk As Boolean) As String()
'QLib.Std.MIde_Mth_Lines.Function MthLineszSrcFm$(Src$(), MthFmIx, Optional WiTopRmk As Boolean)
'QLib.Std.MIde_Mth_Lines.Function MthLyzSrcFm(Src$(), MthFmIx, Optional WiTopRmk As Boolean) As String()
'QLib.Std.MIde_Mth_Lines.Function MthLineszSrcNm$(Src$(), N, Optional WiTopRmk As Boolean)
'QLib.Std.MIde_Mth_Lines.Function MthLineszSrcNmTy$(Src$(), N, ShtMthTy$, Optional WiTopRmk As Boolean)
'QLib.Std.MIde_Mth_Lines.Function MthLyzSrcNm(Src$(), N, Optional WiTopRmk As Boolean) As String()
'QLib.Std.MIde_Mth_Lines.Function MthLyzSrcNmTy(Src$(), N, ShtMthTy$, Optional WiTopRmk As Boolean) As String()
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function IsPrpLin(Lin) As Boolean
'QLib.Std.MIde_Mth_Lin_Is_Hit.Private Sub Z_IsMthLin()
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function HitConstNm(SrcLin, ConstNm) As Boolean
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function HitConstNmDic(SrcLin, ConstNmSet) As Boolean
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function HitMthLin(MthLin, B As WhMth) As Boolean
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function HitMthNm3(A As MthNm3, B As WhMth) As Boolean
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function HitShtMdy(ShtMdy$, ShtMthMdyAy$()) As Boolean
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function IsMthLinzPubZ(Lin) As Boolean
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function IsOptLin(Lin) As Boolean
'QLib.Std.MIde_Mth_Lin_Is_Hit.Function IsMthLinzPub(Lin) As Boolean
'QLib.Std.MIde_Mth_Lin_RetTy.Private Sub Z_MthRetTy()
'QLib.Std.MIde_Mth_Lin_RetTy.Function MthRetTy$(Lin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfItmNy(A$, ItmNy0) As Variant()
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfMthTy$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Sub ShfMthTyAsg(A, OMthTy, ORst$)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfTermAftAs$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfShtMthMdy$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfShtMthTy$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfShtMthKd$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfMthMdy$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfMthNm3(OLin) As MthNm3
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfKd$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfMthSfx$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfNm$(OLin)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function ShfRmk(A) As String()
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function TakMthMdy$(A)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function TakMthKd$(A)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function TakMthTy$(A)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function RmvMdy$(A)
'QLib.Std.MIde_Mth_Lin_Shf_Tak_Rmv.Function RmvMthTy$(A)
'QLib.Std.MIde_Mth_Lin_Fmt.Function NArg(MthLin) As Byte
'QLib.Std.MIde_Mth_Lin_Fmt.Function ArgNy(MthLin) As String()
'QLib.Std.MIde_Mth_Lin_Fmt.Function ArgAy(Lin) As String()
'QLib.Std.MIde_Mth_Lin_Fmt.Function Arg(ArgStr$) As Arg
'QLib.Std.MIde_Mth_Lin_Fmt.Function ArgNyzArgAy(A() As Arg) As String()
'QLib.Std.MIde_Mth_Lin_Fmt.Private Sub Z()
'QLib.Std.MIde_Mth_Chr.Function IsTyChr(A$) As Boolean
'QLib.Std.MIde_Mth_Chr.Function TyChrzTyNm$(TyNm$)
'QLib.Std.MIde_Mth_Chr.Function TyNmzTyChr$(TyChr$)
'QLib.Std.MIde_Mth_Chr.Function RmvTyChr$(A)
'QLib.Std.MIde_Mth_Chr.Function ShfTyChr$(OLin)
'QLib.Std.MIde_Mth_Chr.Function TyChr$(Lin)
'QLib.Std.MIde_Mth_Chr.Function TakTyChr$(S)
'QLib.Std.MIde_Mth_Nm.Sub AsgDNm(DNm$, O1$, O2$, O3$)
'QLib.Std.MIde_Mth_Nm.Property Get MdQNm$()
'QLib.Std.MIde_Mth_Nm.Property Get CurMthQNm$()
'QLib.Std.MIde_Mth_Nm.Function MthQNm$(A As CodeModule, Lin)
'QLib.Std.MIde_Mth_Nm.Function MthNyzPub(Src$()) As String()
'QLib.Std.MIde_Mth_Nm.Function MthNyzMthLinAy(MthLinAy$()) As String()
'QLib.Std.MIde_Mth_Nm.Function Ens1Dot(S) As StrRslt
'QLib.Std.MIde_Mth_Nm.Function Ens2Dot(S) As StrRslt
'QLib.Std.MIde_Mth_Nm.Function MdMth(MthQNm) As MdMth
'QLib.Std.MIde_Mth_Nm.Function RmvMthMdy$(L)
'QLib.Std.MIde_Mth_Nm.Function MthDNmzMthNm3$(A As MthNm3)
'QLib.Std.MIde_Mth_Nm.Function RmvMthNm3$(Lin)
'QLib.Std.MIde_Mth_Nm.Function RelOf_MthSDNm_To_MdNm_OfVbe(Optional WhStr$) As Rel
'QLib.Std.MIde_Mth_Nm.Function RelOf_MthSDNm_To_MdNm_zVbe(A As Vbe, Optional WhStr$) As Rel
'QLib.Std.MIde_Mth_Nm.Function MthNm3(Lin, Optional B As WhMth) As MthNm3
'QLib.Std.MIde_Mth_Nm.Function MthNm$(Lin, Optional B As WhMth)
'QLib.Std.MIde_Mth_Nm.Function MthNmzMthDNm$(MthDNm)
'QLib.Std.MIde_Mth_Nm.Function MthDNmzLin$(MthLin)
'QLib.Std.MIde_Mth_Nm.Function MthSQNyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm.Function MthSQNyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm.Function MthSQNm$(MthQNm)
'QLib.Std.MIde_Mth_Nm.Function MthTyc$(ShtMthTy$)
'QLib.Std.MIde_Mth_Nm.Function MthMdyc$(ShtMthMdy$)
'QLib.Std.MIde_Mth_Nm.Function MthDNm$(Lin, Optional B As WhMth)
'QLib.Std.MIde_Mth_Nm.Function MthNmzLin$(Lin, Optional B As WhMth)
'QLib.Std.MIde_Mth_Nm.Function PrpNm$(Lin)
'QLib.Std.MIde_Mth_Nm.Function MthNmzDNm$(MthNm)
'QLib.Std.MIde_Mth_Nm.Private Sub Z_MthNm()
'QLib.Std.MIde_Mth_Nm.Function MthMdy$(Lin)
'QLib.Std.MIde_Mth_Nm.Function MthKd$(Lin)
'QLib.Std.MIde_Mth_Nm.Function Rpl$(S, SubStr$, By$, Optional Ith% = 1)
'QLib.Std.MIde_Mth_Nm.Function PoszSubStr(S, SubStr) As Pos
'QLib.Std.MIde_Mth_Nm.Property Get Rel0MthNm2MdNm() As Rel
'QLib.Std.MIde_Mth_Nm.Function ModNyzPubMthNm(PubMthNm) As String()
'QLib.Std.MIde_Mth_Nm.Function MthTy$(Lin)
'QLib.Std.MIde_Mth_Nm.Private Sub Z_MthTy()
'QLib.Std.MIde_Mth_Nm.Private Sub Z_MthKd()
'QLib.Std.MIde_Mth_Nm.Private Sub Z()
'QLib.Std.MIde_Mth_Nm_Get.Sub Z_MthNsetOfVbeWiVerb()
'QLib.Std.MIde_Mth_Nm_Get.Sub Z_DryOf_MthNm_Verb_OfVbe()
'QLib.Std.MIde_Mth_Nm_Get.Function DryOf_MthNm_Verb_OfVbe() As Variant()
'QLib.Std.MIde_Mth_Nm_Get.Sub Z_MthNsetOfVbeWoVerb()
'QLib.Std.MIde_Mth_Nm_Get.Property Get MthNyOfVbeWiVerb() As String()
'QLib.Std.MIde_Mth_Nm_Get.Property Get MthNyOfVbeWoVerb() As String()
'QLib.Std.MIde_Mth_Nm_Get.Function HasVerb(Nm) As Boolean
'QLib.Std.MIde_Mth_Nm_Get.Property Get MthNsetOfVbeWiVerb() As Aset
'QLib.Std.MIde_Mth_Nm_Get.Property Get MthNsetOfVbeWoVerb() As Aset
'QLib.Std.MIde_Mth_Nm_Get.Function MthNsetOfVbe(Optional WhStr$) As Aset
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyzSrcFm(Src$(), FmMthIxAy&()) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthNsetOfPj(Optional WhStr$) As Aset
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyOfPj(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyOfPubVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyzPj(A As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthQNyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthQNyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthQNyzMd(A As CodeModule, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthQNmDryzPj(A As VBProject, Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Nm_Get.Function MthQNmDr(MthQNm) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthQNyzPj(A As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNyzVbezPub(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNyzMthNm(A As Vbe, MthNm$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNyzMdMthNm(Md As CodeModule, MthNm$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Property Get MMthNyOfVbe() As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyzFb(Fb) As String()
'QLib.Std.MIde_Mth_Nm_Get.Private Sub Z_MthNyFb()
'QLib.Std.MIde_Mth_Nm_Get.Private Sub Z_MthNyzSrc()
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyzSrc(Src$(), Optional B As WhMth) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyPubzMd(A As CodeModule, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Private Sub Z()
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyzMd(A As CodeModule, Optional B As WhMth) As String()
'QLib.Std.MIde_Mth_Nm_Get.Private Sub ZZ_MthNyzSrc()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNmSqzPj(A As VBProject) As Variant()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNmWszPj(A As VBProject) As Worksheet
'QLib.Std.MIde_Mth_Nm_Get.Function MthNyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthAsetVbe(Optional WhStr$) As Aset
'QLib.Std.MIde_Mth_Nm_Get.Property Get CurMthNyzMd() As String()
'QLib.Std.MIde_Mth_Nm_Get.Private Sub Z_MthDNy()
'QLib.Std.MIde_Mth_Nm_Get.Private Sub Z_MthDNyzSrc()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNyzMd(A As CodeModule, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Get.Function MthDNyzSrc(Src$(), Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Nm_Dic.Function Md_MthNmDic(A As CodeModule) As Dictionary
'QLib.Std.MIde_Mth_Nm_Dic.Private Sub Z_Src_MthNmDic()
'QLib.Std.MIde_Mth_Nm_Dic.Private Sub Z()
'QLib.Std.MIde_Mth_Nm_Drs.Function MthNmCmlSetVbe(Optional WhStr$) As Aset
'QLib.Std.MIde_Mth_Nm_Drs.Function MthNmDrsVbe(Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Nm_Drs.Function MthNmDrsPj(Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Nm_Drs.Function MthNmDrsMd(Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Nm_Drs.Private Function MthNmDrszMd(M As CodeModule, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Nm_Drs.Private Function MthNmDrszVbe(A As Vbe, Optional WhStr$) As Drs
'QLib.Std.MIde_Mth_Nm_Drs.Function MthNmDrszPj(A As VBProject, Optional WhStr$)
'QLib.Std.MIde_Mth_Nm_Drs.Private Function MthNmDryzMd(M As CodeModule, Optional B As WhMth) As Variant()
'QLib.Std.MIde_Mth_Nm_Drs.Private Function MthNmDryzVbe(A As Vbe, Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Nm_Drs.Private Function MthNmDryzPj(P As VBProject, Optional WhStr$) As Variant()
'QLib.Std.MIde_Mth_Nm_Drs.Private Function MthNmDryzSrc(Src$(), Optional B As WhMth) As Variant()
'QLib.Std.MIde_Mth_Nm_Drs_Dup.Function DupMthDrsPj() As Drs
'QLib.Std.MIde_Mth_Nm_Drs_Dup.Private Function DupMthDrszPj(A As VBProject) As Drs
'QLib.Std.MIde_Mth_Nm_Drs_Dup.Private Function AddColzMthLines(MthNmDrs As Drs) As Drs
'QLib.Std.MIde_Mth_Nm_Drs_Dup.Private Function MthLinesAyzDry_Md_MthNm_ShtMthTy(Dry()) As String()
'QLib.Std.MIde_Mth_Nm_Has.Function HasMthSrc(Src$(), MthNm) As Boolean
'QLib.Std.MIde_Mth_Nm_Has.Function HasMthMd(A As CodeModule, MthNm) As Boolean
'QLib.Std.MIde_Mth_Op.Function TmpMod() As CodeModule
'QLib.Std.MIde_Mth_Op.Sub ClrTmpMod()
'QLib.Std.MIde_Mth_Op.Property Get TmpModNm$()
'QLib.Std.MIde_Mth_Op.Function NowNm$()
'QLib.Std.MIde_Mth_Op.Sub RmvMth(A As CodeModule, MthNmNN)
'QLib.Std.MIde_Mth_Op.Sub RmvMthzNm(A As CodeModule, MthNm, Optional WiTopRmk As Boolean)
'QLib.Std.MIde_Mth_Op.Sub RmvMdMth(Md As CodeModule, MthNm)
'QLib.Std.MIde_Mth_Op.Private Sub Z_RmvMdMth()
'QLib.Std.MIde_Mth_Op.Private Sub Z()
'QLib.Std.MIde_Mth_Op.Sub CpyMdMthToMd(Md As CodeModule, MthNm, ToMd As CodeModule, Optional IsSilent As Boolean)
'QLib.Std.MIde_Mth_Op.Sub MovMthTo(MthNm, ToMdNm$)
'QLib.Std.MIde_Mth_Op.Sub MovMdMthTo(Md As CodeModule, MthNm, ToMd As CodeModule)
'QLib.Std.MIde_Mth_Pfx.Private Sub Z_MthPfx()
'QLib.Std.MIde_Mth_Pfx.Private Sub ZZ_MthPfx()
'QLib.Std.MIde_Mth_Pfx.Function MthPfxAyMd(A As CodeModule) As String()
'QLib.Std.MIde_Mth_Pfx.Function MthPfx$(MthNm)
'QLib.Std.MIde_Mth_Pfx.Private Sub Z()
'QLib.Std.MIde_Mth_Pm_Arg.Function MthPm$(MthLin)
'QLib.Std.MIde_Mth_Pm_Arg.Property Get ArgAsetOfPj() As Aset
'QLib.Std.MIde_Mth_Pm_Arg.Function ArgAsetzPj(A As VBProject) As Aset
'QLib.Std.MIde_Mth_Pm_Arg.Private Sub Z_ArgAsetOfPj()
'QLib.Std.MIde_Mth_Pm_Arg.Function DimItmzArg$(Arg)
'QLib.Std.MIde_Mth_Pm_Arg.Function ArgSfx$(Arg)
'QLib.Std.MIde_Mth_PurePrp.Sub ImPurePrpPjBrw()
'QLib.Std.MIde_Mth_PurePrp.Function ImpPurePrpLyOfPj() As String()
'QLib.Std.MIde_Mth_PurePrp.Function ImPurePrpLyzPj(A As VBProject) As String()
'QLib.Std.MIde_Mth_PurePrp.Function ImPurePrpLyzMd(A As CodeModule) As String()
'QLib.Std.MIde_Mth_PurePrp.Private Sub Z_ImPurePrpLyzSrc()
'QLib.Std.MIde_Mth_PurePrp.Function ImPurePrpLyzSrc(Src$()) As String()
'QLib.Std.MIde_Mth_PurePrp.Sub PurePrpPjBrw()
'QLib.Std.MIde_Mth_PurePrp.Function PurePrpLyPj() As String()
'QLib.Std.MIde_Mth_PurePrp.Function PurePrpLyzPj(A As VBProject) As String()
'QLib.Std.MIde_Mth_PurePrp.Function PurePrpIxAy(Src$()) As Long()
'QLib.Std.MIde_Mth_PurePrp.Function PurePrpLnoAy(A As CodeModule) As Long()
'QLib.Std.MIde_Mth_PurePrp.Function PurePrpLy(A As CodeModule) As String()
'QLib.Std.MIde_Mth_PurePrp.Function PurePrpNy(A As CodeModule) As String()
'QLib.Std.MIde_Mth_PurePrp.Function LetSetPrpNset(MthLinAy$()) As Aset
'QLib.Std.MIde_Mth_PurePrp.Private Function LetSetPrpNm$(Lin)
'QLib.Std.MIde_Mth_PurePrp.Function IsImPurePrpLin(Lin, LetSetPrpNset As Aset) As Boolean
'QLib.Std.MIde_Mth_PurePrp.Function IsPurePrpLin(Lin) As Boolean
'QLib.Std.MIde_Mth_PurePrp.Function HasMthPm(MthLin) As Boolean
'QLib.Std.MIde_Mth_Rmk.Sub UnRmkMth(A As CodeModule, MthNm$)
'QLib.Std.MIde_Mth_Rmk.Sub RmkMth(A As CodeModule, MthNm$)
'QLib.Std.MIde_Mth_Rmk.Private Sub ZZ_RmkMth()
'QLib.Std.MIde_Mth_Rmk.Function NxtSrcIx&(Src$(), Ix&)
'QLib.Std.MIde_Mth_Rmk.Function NxtMdLno&(A As CodeModule, Lno&)
'QLib.Std.MIde_Mth_Rmk.Sub UnRmkMdzFTIxAy(A As CodeModule, B() As FTIx)
'QLib.Std.MIde_Mth_Rmk.Sub UnRmkMdzFTIx(A As CodeModule, B As FTIx)
'QLib.Std.MIde_Mth_Rmk.Sub RmkMdzFTIxAy(A As CodeModule, B() As FTIx)
'QLib.Std.MIde_Mth_Rmk.Sub RmkMdzFTIx(A As CodeModule, B As FTIx)
'QLib.Std.MIde_Mth_Rmk.Function IsRmkedzMdFTIx(A As CodeModule, B As FTIx) As Boolean
'QLib.Std.MIde_Mth_Rmk.Function IsRmkedzSrc(A$()) As Boolean
'QLib.Std.MIde_Mth_Rmk.Function MthCxtFTIx(Src$(), MthFTIx As FTIx) As FTIx
'QLib.Std.MIde_Mth_Rmk.Function MthCxtLy(MthLy$()) As String()
'QLib.Std.MIde_Mth_Rmk.Function MthCxtFTIxAy(Src$(), MthNm$) As FTIx()
'QLib.Std.MIde_Mth_Rmk.Private Sub ZZ_MthCxtFTIxAy  ()
'QLib.Std.MIde_Mth_SubZDash.Function MthLinAyzSubZDashMd(A As CodeModule) As String()
'QLib.Std.MIde_Mth_SubZDash.Function MthNyzSubZDashMd(A As CodeModule) As String()
'QLib.Std.MIde_Mth_SubZDash.Function IsSubZDashMthLin(MthLin) As Boolean
'QLib.Std.MIde_Mth_TopRmk.Private Sub Z_MthFTIxAyzSrcMth()
'QLib.Std.MIde_Mth_TopRmk.Function AyRmvBlankLin(Ay) As String()
'QLib.Std.MIde_Mth_TopRmk.Function MthTopRmkLy(Src$(), MthFmIx) As String()
'QLib.Std.MIde_Mth_TopRmk.Function MthTopRmkIx&(Src$(), MthFmIx)
'QLib.Std.MIde_Mth_TopRmk.Function MthTopRmkLnoMdFm&(Md As CodeModule, MthLno)
'QLib.Std.MIde_Mth_Ty_ConstAy.Property Get PrpTyAy() As String()
'QLib.Std.MIde_Mth_Ty_ConstAy.Property Get MthTyAy() As String()
'QLib.Std.MIde_Mth_Ty_ConstAy.Property Get MthMdyAy() As String()
'QLib.Std.MIde_Mth_Ty_ConstAy.Property Get ShtMthMdyAy() As String()
'QLib.Std.MIde_Mth_Ty_ConstAy.Property Get ShtMthKdAy() As String()
'QLib.Std.MIde_Mth_Ty_ConstAy.Property Get ShtMthTyAy() As String()
'QLib.Std.MIde_Mth_Ty_ConstAy.Property Get MthKdAy() As String()
'QLib.Std.MIde_Mth_Ty_ConstAy.Property Get DclItmAy() As String()
'QLib.Std.MIde_Pj.Sub ThwIfCompileBtn(NEPjNm$)
'QLib.Std.MIde_Pj.Function CvPj(I) As VBProject
'QLib.Std.MIde_Pj.Function IsPjNm(A) As Boolean
'QLib.Std.MIde_Pj.Function Pj(PjNm) As VBProject
'QLib.Std.MIde_Pj.Sub RmvPj(Pj As VBProject)
'QLib.Std.MIde_Pj.Function StrzCurPjf$()
'QLib.Std.MIde_Pj.Function PjfPj$()
'QLib.Std.MIde_Pj.Function PjPth$(A As VBProject)
'QLib.Std.MIde_Pj.Function Pjf$(A As VBProject)
'QLib.Std.MIde_Pj.Function PjFnn$(A As VBProject)
'QLib.Std.MIde_Pj.Function IsUsrLibPj(A As VBProject) As Boolean
'QLib.Std.MIde_Pj.Function MdzPj(A As VBProject, Nm) As CodeModule
'QLib.Std.MIde_Pj.Sub Compile()
'QLib.Std.MIde_Pj.Sub CompilePj(A As VBProject)
'QLib.Std.MIde_Pj.Sub ActPj(A As VBProject)
'QLib.Std.MIde_Pj.Sub SavPj(A As VBProject)
'QLib.Std.MIde_Pj.Private Sub ZZ_SavPj()
'QLib.Std.MIde_Pj.Private Sub Z_PjCompile()
'QLib.Std.MIde_Pj.Private Sub ZZ()
'QLib.Std.MIde_Pj.Function IsProtectzInf(A As VBProject) As Boolean
'QLib.Std.MIde_Pj.Function IsProtect(A As VBProject) As Boolean
'QLib.Std.MIde_Pj.Sub BrwPthPj()
'QLib.Std.MIde_Pj.Function PjzXls(A As Excel.Application, Fxa) As VBProject
'QLib.Std.MIde_Pj.Function FstMd(A As VBProject) As CodeModule
'QLib.Std.MIde_Pj.Function FstMod(A As VBProject) As CodeModule
'QLib.Std.MIde_Pj.Function IsFbaPj(A As VBProject) As Boolean
'QLib.Std.MIde_Pjf.Sub ClsPjf(Pjf)
'QLib.Std.MIde_Pjf.Function VbePjf(Pjf) As Vbe
'QLib.Std.MIde_Pjf.Sub OpnPjf(Pjf)  ' Return either Xls.Application (Xls) or Acs.Application (Function-static)
'QLib.Std.MIde_Pjf.Sub RmvPjzXlsPjf(Xls As Excel.Application, Pjf)
'QLib.Std.MIde_Pjf.Function TmpFxa$(Optional Fdr$, Optional Fnn$)
'QLib.Std.MIde_Pj_Cur.Property Get CurPj() As VBProject
'QLib.Std.MIde_Pj_Cur.Function EnsMd(MdNm$) As CodeModule
'QLib.Std.MIde_Pj_Cur.Function EnsModzPj(A As VBProject, ModNm$) As CodeModule
'QLib.Std.MIde_Pj_Cur.Function HasMd(A As VBProject, MdNm) As Boolean
'QLib.Std.MIde_Pj_Cur.Sub ThwIfNotMod(A As CodeModule, Fun$)
'QLib.Std.MIde_Pj_Cur.Function HasMod(A As VBProject, ModNm) As Boolean
'QLib.Std.MIde_Pj_Cur.Property Get PjNm$()
'QLib.Std.MIde_Pj_Cur.Function PthPj$()
'QLib.Std.MIde_Pj_Cur.Sub BrwPjPth()
'QLib.Std.MIde_Pj_Dte.Function PjDteFb(A) As Date
'QLib.Std.MIde_Pj_Dte.Function PjDtePjf(Pjf) As Date
'QLib.Std.MIde_Pj_Dte.Function AcsPjDte(A As Access.Application)
'QLib.Std.MIde_PrpFun.Function IsPrpFunLin(Lin) As Boolean
'QLib.Std.MIde_PrpFun.Function PrpFunLnoAy(A As CodeModule) As Long()
'QLib.Std.MIde_PrpFun.Sub EnsPjFunzMd(Md As CodeModule, Optional WhatIf As Boolean)
'QLib.Std.MIde_PrpFun.Sub EnsPrpFunzPj(Pj As VBProject, Optional WhatIf As Boolean)
'QLib.Std.MIde_PrpFun.Sub EnsPrpFun()
'QLib.Std.MIde_PrpFun.Private Sub EnsPrpFunMdLno(A As CodeModule, Lno, Optional WhatIf As Boolean)
'QLib.Std.MIde_Src.Private Sub ZZ_SrcDcl()
'QLib.Std.MIde_Src.Private Sub ZZ_FstMthIx()
'QLib.Std.MIde_Src.Function SrczMdNm(MdNm$) As String()
'QLib.Std.MIde_Src.Private Sub ZZ_MthTopRmIx_SrcFm()
'QLib.Std.MIde_Src.Private Property Get ZZSrc() As String()
'QLib.Std.MIde_Src.Private Property Get ZZSrcLin$()
'QLib.Std.MIde_Src.Private Sub Z_MthNyzSrc()
'QLib.Std.MIde_Src.Sub AsgMthDr(MthDr, OMdy$, OTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OTopRmk$)
'QLib.Std.MIde_Src.Private Sub Z_MthLinDryzSrc()
'QLib.Std.MIde_Src.Function MthLinDryzSrc(Src$()) As Variant()
'QLib.Std.MIde_Src.Function MthDNyzPj(A As VBProject, Optional WhStr$) As String()
'QLib.Std.MIde_Src.Function CurSrcLines$()
'QLib.Std.MIde_Src.Function SrcLinesOfPj$()
'QLib.Std.MIde_Src.Function SrcLineszPj$(A As VBProject)
'QLib.Std.MIde_Src.Function SrcLineszMd$(A As CodeModule)
'QLib.Std.MIde_Src.Function SrczMd(A As CodeModule) As String()
'QLib.Std.MIde_Src.Function Src(A As CodeModule) As String()
'QLib.Std.MIde_Src.Function SrcOfPj() As String()
'QLib.Std.MIde_Src.Function SrcOfVbe() As String()
'QLib.Std.MIde_Src.Function SrczPj(A As VBProject) As String()
'QLib.Std.MIde_Src.Function SrczVbe(A As Vbe) As String()
'QLib.Std.MIde_Src.Property Get CurSrc() As String()
'QLib.Std.MIde_Src.Function NMthzSrc%(A$())
'QLib.Std.MIde_Src.Function NUsrTySrc%(A$())
'QLib.Std.MIde_Src_Ret.Function LineszMdFTIx$(A As CodeModule, B As FTIx)
'QLib.Std.MIde_Src_Ret.Function LyzMdFTIx(A As CodeModule, B As FTIx) As String()
'QLib.Std.MIde_Src_Ret.Function LyMdRe(A As CodeModule, B As RegExp) As String()
'QLib.Std.MIde_Src_Ret.Function LyzPjPatn(A As VBProject, Patn$)
'QLib.Std.MIde_Srt.Private Sub SrtzPj(A As VBProject)
'QLib.Std.MIde_Srt.Private Sub ZZ()
'QLib.Std.MIde_Srt.Private Sub ZZ_Dcl_BefAndAft_Srt()
'QLib.Std.MIde_Srt.Private Sub ZZ_SrtMd()
'QLib.Std.MIde_Srt.Private Sub ZZ_SrtedSrcLineszMd()
'QLib.Std.MIde_Srt.Private Sub Z_SrcLinesSrt()
'QLib.Std.MIde_Srt.Function MthNm3zDNm(MthDNm) As MthNm3
'QLib.Std.MIde_Srt.Function MthSrtKey$(MthDNm)
'QLib.Std.MIde_Srt.Function SrtedSrczMd(A As CodeModule) As String()
'QLib.Std.MIde_Srt.Function SrtedMdDicOfPj() As Dictionary
'QLib.Std.MIde_Srt.Function SrtedMdDiczPj(A As VBProject) As Dictionary
'QLib.Std.MIde_Srt.Function SrcSrt(Src$()) As String()
'QLib.Std.MIde_Srt.Function SrcDic(Src$(), Optional WiTopRmk As Boolean) As Dictionary
'QLib.Std.MIde_Srt.Function SrtedSrcDic(Src$()) As Dictionary
'QLib.Std.MIde_Srt.Function SrcLinesSrt$(Src$())
'QLib.Std.MIde_Srt.Function SrtedSrcLinesOfMd$()
'QLib.Std.MIde_Srt.Function SrtedSrcLineszMd$(A As CodeModule)
'QLib.Std.MIde_Srt.Sub BrwSrtRptzMd(A As CodeModule)
'QLib.Std.MIde_Srt.Sub BrwSrtedMdDic()
'QLib.Std.MIde_Srt.Sub RplPj(A As VBProject, MdDic As Dictionary)
'QLib.Std.MIde_Srt.Sub Srt()
'QLib.Std.MIde_Srt.Sub SrtPj()
'QLib.Std.MIde_Srt.Sub SrtzMd(A As CodeModule)
'QLib.Std.MIde_Srt_Rpt.Private Function SrtRpt(Src$()) As String()
'QLib.Std.MIde_Srt_Rpt.Private Sub Z_SrtRpt()
'QLib.Std.MIde_Srt_Rpt.Property Get SrtRptMd() As String()
'QLib.Std.MIde_Srt_Rpt.Function SrtRptzPj(A As VBProject) As String()
'QLib.Std.MIde_Srt_Rpt.Function SrtRptDiczPj(A As VBProject) As Dictionary
'QLib.Std.MIde_Srt_Rpt.Function SrtRptzMd(A As CodeModule) As String()
'QLib.Std.MIde_Srt_Rpt.Function SrtDicMd(A As CodeModule) As Dictionary
'QLib.Std.MIde_Ty_Component.Function ShtCmpTyzMd$(A As CodeModule)
'QLib.Std.MIde_Ty_Component.Function ShtCmpTy$(A As vbext_ComponentType)
'QLib.Std.MIde_Ty_Component.Function CmpTyzMd(Md As CodeModule) As vbext_ComponentType
'QLib.Std.MIde_Ty_Component.Function CmpTy(ShtCmpTy) As vbext_ComponentType
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function MthKdByMthTy$(MthTy)
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function IsMthTy(Str$) As Boolean
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function IsMthMdy(A$) As Boolean
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function MthMdyBySht$(ShtMthMdy)
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function ShtMthMdy$(MthMdy)
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function MthTyBySht$(ShtMthTy)
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function ShtMthTy$(MthTy)
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function ShtMthKdByShtMthTy$(ShtMthTy)
'QLib.Std.MIde_Ty_Mth_ShtTy_Cv.Function ShtMthKd$(MthKd)
'QLib.Std.MIde_Vbe.Function CvVbe(A) As Vbe
'QLib.Std.MIde_Vbe.Sub DmpPjIsSav()
'QLib.Std.MIde_Vbe.Function PjIsSavDryzVbe(A As Vbe) As Variant()
'QLib.Std.MIde_Vbe.Function PjIsSavDrszVbe(A As Vbe) As Drs
'QLib.Std.MIde_Vbe.Function Vbe_Pj(A As Vbe, PjNm$) As VBProject
'QLib.Std.MIde_Vbe.Function PjzPjfVbe(Vbe As Vbe, PjFil) As VBProject
'QLib.Std.MIde_Vbe.Function MdDryzVbe(A As Vbe, Optional WhStr$) As Variant()
'QLib.Std.MIde_Vbe.Function MdDr(A As CodeModule) As Variant()
'QLib.Std.MIde_Vbe.Sub SavVbe(A As Vbe)
'QLib.Std.MIde_Vbe.Function VisWinCntz%(A As Vbe)
'QLib.Std.MIde_Vbe.Sub CompileVbe(A As Vbe)
'QLib.Std.MIde_Vbe.Function MthLinDryzVbe(A As Vbe, Optional WhStr$) As Variant()
'QLib.Std.MIde_Vbe.Function PjAy(A As Vbe, Optional WhStr$, Optional NmPfx$) As VBProject()
'QLib.Std.MIde_Vbe.Function ItrwNmStr(Itr, WhStr$, Optional NmPfx$)
'QLib.Std.MIde_Vbe.Function ItrwNm(Itr, B As WhNm)
'QLib.Std.MIde_Vbe.Function PjItr(A As Vbe, Optional WhStr$, Optional NmPfx$)
'QLib.Std.MIde_Vbe.Property Get PjfAyOfVbe() As String()
'QLib.Std.MIde_Vbe.Function PjfAyzVbe(A As Vbe) As String()
'QLib.Std.MIde_Vbe.Function PjNyOfVbe(Optional WhStr$, Optional NmPfx$) As String()
'QLib.Std.MIde_Vbe.Function PjNyzVbe(A As Vbe, Optional WhStr$, Optional NmPfx$) As String()
'QLib.Std.MIde_Vbe.Function FstQPj(A As Vbe) As VBProject
'QLib.Std.MIde_Vbe.Function MthWbVbe(A As Vbe) As Workbook
'QLib.Std.MIde_Vbe.Function VbeSrtRpt() As String()
'QLib.Std.MIde_Vbe.Function HasVbeBar(A As Vbe, Nm$) As Boolean
'QLib.Std.MIde_Vbe.Function Vbe_HasPj(A As Vbe, PjNm) As Boolean
'QLib.Std.MIde_Vbe.Function HasPjfVbe(A As Vbe, Ffn) As Boolean
'QLib.Std.MIde_Vbe.Function SrtRptVbe(A As Vbe) As String()
'QLib.Std.MIde_Vbe.Private Sub ZZ_VbeFunPfx()
'QLib.Std.MIde_Vbe.Private Sub ZZ_MthNyzVbe()
'QLib.Std.MIde_Vbe.Private Sub ZZ_MthNyzVbeWh()
'QLib.Std.MIde_Vbe.Private Sub ZZ()
'QLib.Std.MIde_Vbe.Private Sub Z()
'QLib.Std.MIde_Vbe_Cur.Property Get CurVbe() As Vbe
'QLib.Std.MIde_Vbe_Cur.Function HasBar(Nm$) As Boolean
'QLib.Std.MIde_Vbe_Cur.Function HasPjf(Pjf) As Boolean
'QLib.Std.MIde_Vbe_Cur.Function PjzPjf(A) As VBProject
'QLib.Std.MIde_Vbe_Cur.Function MdDrszVbe(A As Vbe, Optional WhStr$) As Drs
'QLib.Std.MIde_Vbe_Cur.Function MdTblFny() As String()
'QLib.Std.MIde_Vbe_Cur.Sub SavCurVbe()
'QLib.Std.MIde_Wh.Function WhMthzPfx(WhMthNmPfx$, Optional InclPrv As Boolean) As WhMth
'QLib.Std.MIde_Wh.Function WhMthzSfx(WhMthNmSfx$, Optional InclPrv As Boolean) As WhMth
'QLib.Std.MIde_Wh.Function WhMthzStr(WhStr$) As WhMth
'QLib.Std.MIde_Wh.Function WhMdMth(Optional Md As WhMd, Optional Mth As WhMth) As WhMdMth
'QLib.Std.MIde_Wh.Function WhMdzWhMdMth(A As WhMdMth) As WhMd
'QLib.Std.MIde_Wh.Function WhMthzWhMdMth(A As WhMdMth) As WhMth
'QLib.Std.MIde_Wh.Function WhMth(ShtMdy$(), ShtKd$(), Nm As WhNm) As WhMth
'QLib.Std.MIde_Wh.Function WhPjMth(Optional Pj As WhNm, Optional MdMth As WhMdMth) As WhPjMth
'QLib.Std.MIde_Wh.Function WhNm(Patn$, LikeAy$(), ExlLikAy$()) As WhNm
'QLib.Std.MIde_Wh.Function WhMd(CmpTy() As vbext_ComponentType, Nm As WhNm) As WhMd
'QLib.Std.MIde_Wh.Function WhMdzStr(WhStr$) As WhMd
'QLib.Std.MIde_Ws_Cmp.Function MdzWs(A As Worksheet) As CodeModule
'QLib.Std.MIde_Ws_Cmp.Function CmpzWs(A As Worksheet) As VBComponent
'QLib.Cls.MthCnt.Friend Function Init(MdNm$, NPubSub%, NPubFun%, NPubPrp%, NPrvSub%, NPrvFun%, NPrvPrp%, NFrdSub%, NFrdFun%, NFrdPrp%) As MthCnt
'QLib.Cls.MthCnt.Property Get N%()
'QLib.Cls.MthCnt.Function Lin$(Optional Hdr As eHdr)
'QLib.Cls.MthInf.Property Get ArgAy() As String()
'QLib.Cls.MthInf.Property Let ArgAy(V$())
'QLib.Cls.MthNm3.Friend Function Init(MthMdy, MthTy, Nm) As MthNm3
'QLib.Cls.MthNm3.Function Lin$(Optional Hdr As eHdr)
'QLib.Cls.MthNm3.Property Get DNm$()
'QLib.Cls.MthNm3.Property Get IsEmp() As Boolean:  IsEmp = Nm = "":                              End Property
'QLib.Cls.MthNm3.Property Get ShtMdy$():          ShtMdy = ShtMthMdy(MthMdy):                    End Property
'QLib.Cls.MthNm3.Property Get ShtTy$():            ShtTy = ShtMthTy(MthTy):                      End Property
'QLib.Cls.MthNm3.Property Get ShtKd$():            ShtKd = ShtMthKd(T1(MthTy)):                  End Property
'QLib.Cls.MthNm3.Property Get MthKd$():            MthKd = T1(MthTy):                            End Property
'QLib.Cls.MthNm3.Property Get IsPub() As Boolean:  IsPub = (MthMdy = "") Or (MthMdy = "Public"): End Property
'QLib.Cls.MthNm3.Property Get IsPrv() As Boolean:  IsPrv = MthMdy = "Private":                   End Property
'QLib.Cls.MthNm3.Property Get IsFrd() As Boolean:  IsFrd = MthMdy = "Friend":                    End Property
'QLib.Cls.MthNm3.Property Get IsSub() As Boolean:  IsSub = MthTy = "Sub":                        End Property
'QLib.Cls.MthNm3.Property Get IsPrp() As Boolean:  IsPrp = MthKd = "Property":                   End Property
'QLib.Cls.MthNm3.Property Get IsFun() As Boolean:  IsFun = MthTy = "Function":                   End Property
'QLib.Cls.MthNm3.Property Get MthNmDr() As String()
'QLib.Std.MTp_SqyRslt.Function SqyRsltzEr(Sqy$(), Er$()) As SqyRslt
'QLib.Std.MTp_SqyRslt.Function SqyRslt(SqTp$) As SqyRslt
'QLib.Std.MTp_SqyRslt.Private Function LnxAyzBlk(A() As Blk, BlkTy$) As Lnx()
'QLib.Std.MTp_SqyRslt_1_BlkAy.Function BlkAy(SqTp$) As Blk()
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function BlkAyzGpAy(A() As Gp) As Blk()
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function BlkTyzGp$(A As Gp)
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function GpAy(Ly$()) As Gp()
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function GpAyzRmvRmk(A() As Gp) As Gp()
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function RmvRmkzGp(A As Gp) As Gp
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function Blk(A As Gp) As Blk
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function BlkTy$(Ly$())
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function IsPmLy(A$()) As Boolean
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function IsRmLy(A$()) As Boolean
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function IsSqLy(A$()) As Boolean
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Function IsSwLy(Ly$()) As Boolean
'QLib.Std.MTp_SqyRslt_1_BlkAy.Private Sub ZZ()
'QLib.Std.MTp_SqyRslt_2_PmRsltzLnxAy.Function LnxAyRsltzEr(LnxAy() As Lnx, Er$()) As LnxAyRslt
'QLib.Std.MTp_SqyRslt_2_PmRsltzLnxAy.Private Function PmRsltzEr(Pm As Dictionary, Er$()) As PmRslt
'QLib.Std.MTp_SqyRslt_2_PmRsltzLnxAy.Function PmRsltzLnxAy(A() As Lnx) As PmRslt
'QLib.Std.MTp_SqyRslt_2_PmRsltzLnxAy.Private Function PmLyRslt(A() As Lnx) As LyRslt
'QLib.Std.MTp_SqyRslt_2_PmRsltzLnxAy.Private Function LnxAyRsltzDupKey(A() As Lnx) As LnxAyRslt
'QLib.Std.MTp_SqyRslt_2_PmRsltzLnxAy.Private Function LnxAyRsltzPercentagePfx(A As LnxAyRslt) As LnxAyRslt
'QLib.Std.MTp_SqyRslt_31_SwBrkAy.Private Function SwBrk(A As Lnx) As SwBrk
'QLib.Std.MTp_SqyRslt_31_SwBrkAy.Function SwBrkAy(A() As Lnx) As SwBrk()
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function SwBrkAyRsltzEr(SwBrkAy() As SwBrk, Er$()) As SwBrkAyRslt
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Function SwBrkAyRslt(A() As SwBrk, Pm As Dictionary) As SwBrkAyRslt
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function Msgz$(A As SwBrk, B$)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzDupNm(A() As SwBrk) As String()
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzLeftOvrAftEvl(A() As SwBrk, Sw As Dictionary) As String()
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzNoNm$(A As SwBrk)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzOpStrEr$(A As SwBrk)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzPfx$(A As SwBrk)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzTermCntAndOr$(A As SwBrk)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzTermCntEqNe$(A As SwBrk)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzTermMustBegWithQuestOrAt$(TermAy$(), A As SwBrk)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzTermNotInPm$(TermAy$(), A As SwBrk)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function MsgzTermNotInSw$(TermAy$(), A As SwBrk, SwNm As Dictionary)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function ErzDupNm(A() As SwBrk, O() As SwBrk) As String()
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function ErzFld(A() As SwBrk, O() As SwBrk) As String()
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function ErzFldLin(A As SwBrk, SwNm As Dictionary, Pm As Dictionary, OEr$()) As SwBrk
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function ErzLeftOvr(A() As SwBrk, O() As SwBrk) As String()
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function ErzLin1$(IO As SwBrk)
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function ErzLin(A() As SwBrk, O() As SwBrk) As String()
'QLib.Std.MTp_SqyRslt_32_SwBrkAyRslt.Private Function ErzPfx(A() As Lnx, OEr$()) As Lnx()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function SwRsltzEr(FldSw As Dictionary, StmtSw As Dictionary, Er$()) As SwRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Sub AAMain()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Function SwRsltzLnxAy(A() As Lnx, Pm As Dictionary) As SwRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function SwRsltzSwDic(SwDic As Dictionary, Er$()) As SwRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function EvlBoolTerm(BoolTerm, Sw As Dictionary, BoolTermPm As Dictionary) As BoolRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function EvlSwBrkLin(A As SwBrk, Sw As Dictionary) As DicRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function EvlT1(T1$, Sw As Dictionary, SwTermPm As Dictionary) As StrRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function EvlT1T2(T1$, T2$, EQ_NE$, Sw As Dictionary) As BoolRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function IsEqStrRslt(A As StrRslt, B As StrRslt) As Boolean
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function EvlT2(T2$, Sw As Dictionary, SwTermPm As Dictionary) As StrRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function EvlTerm(SwTerm$, Sw As Dictionary, SwTermPm As Dictionary) As StrRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function EvlTermAy1(SwTermAy$(), AND_OR$, Sw As Dictionary, BoolTermPm As Dictionary) As Boolean()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function EvlTermAy(SwTermAy$(), AND_OR$, Sw As Dictionary, BoolTermPm As Dictionary) As BoolRslt
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Sub A()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Function SwBrkAyFmt(A() As SwBrk) As String()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Property Get OpStrAy() As String()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Sub BrwSwBrkAy(A() As SwBrk)
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Sub Z_SwRslt()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Sub ZZ()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Private Sub Z()
'QLib.Std.MTp_SqyRslt_3_SwRsltzLnxAy.Function CvSwBrk(A) As SwBrk
'QLib.Std.MTp_SqyRslt_41_ErzSqLy.Function ErzSqLy(SqLy$()) As LyRslt
'QLib.Std.MTp_SqyRslt_41_ErzSqLy.Private Function MsgAp_Lin_TyEr(A As Lnx) As String()
'QLib.Std.MTp_SqyRslt_41_ErzSqLy.Private Function MsgMustBeIntoLin$(A As Lnx)
'QLib.Std.MTp_SqyRslt_41_ErzSqLy.Private Function MsgMustBeSelorSelDis$(A As Lnx)
'QLib.Std.MTp_SqyRslt_41_ErzSqLy.Private Function MsgMustNotHasSpcInTbl_NmOfIntoLin$(A As Lnx)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function SqlRsltzEr(Sql$, Er$()) As SqlRslt
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Sub AAMain()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Function SqyRsltzGpAy(SqGpAy() As Gp, PmDic As Dictionary, StmtSwDic As Dictionary, FldSwDic As Dictionary) As SqyRslt
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function SqlRsltzSqLy(SqLy$()) As SqlRslt
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function SqlRsltDrp(SqLy$()) As SqlRslt
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function SqlRsltSel(SqLy$(), ExprDic As Dictionary) As SqlRslt
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function RmvExprLin(SqLy$()) As String()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function SqlRsltUpd(A$(), E As Dictionary) As SqlRslt
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function Fny_WhActive(A$()) As String()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function ExprDic(A) As Dictionary
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function StmtTy(SqLy$()) As eStmtTy
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function StmtSwKey$(SqLy$(), Ty As eStmtTy)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function StmtSwKey_SEL$(SqLy$())
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function StmtSwKey_UPD$(SqLy)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function FndVy(K, E As Dictionary, OVy$(), OQ$)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function FndValPair(K, E As Dictionary, OV1, OV2)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function IsXXX(A$(), XXX$) As Boolean
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Property Get SampExprDic() As Dictionary
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Property Get SampSqLnxAy() As Lnx()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function SqLyRslt(A As Gp) As LyRslt
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PushSqlRslt(A As SqyRslt, B As SqlRslt) As SqyRslt
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function MsgAndLinOp_ShouldBe_BetOrIn(A)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XAnd(A$(), E As Dictionary)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XGp$(L$, E As Dictionary)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XJnOrLeftJn(A$(), E As Dictionary) As String()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopJnOrLeftJn(A$()) As String()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopXXXOpt$(A$(), XXX$)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopXXX$(A$(), XXX$)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopGp$(A$())
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopWh$(A$())
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopAnd(A$()) As String()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopXorYOpt$(A$(), X$, Y$)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopMulXorYOpt(A$(), X$, Y$) As String()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function PopMulXXX(A$(), XXX$) As String()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XSel$(A$, E As Dictionary)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XSelFny(Fny$(), FldSw As Dictionary) As String()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XSet(A As Lnx, E As Dictionary, OEr$())
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XUpd(A As Lnx, E As Dictionary, OEr$())
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XWh$(L$, E As Dictionary)
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XWhBetNbr$(A As Lnx, E As Dictionary, OEr$())
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XWhExpr(A As Lnx, E As Dictionary, OEr$())
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Function XWhInNbrLis$(A As Lnx, E As Dictionary, OEr$())
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Sub Z_SqlRsltSel()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Sub Z_ExprDic()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Sub Z_SqyRslt()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Sub Z_Sel()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Sub Z_StmtSwKey()
'QLib.Std.MTp_SqyRslt_4_SqyRsltzGpAy.Private Sub Z()
'QLib.Std.MTp_SqyRslt_5_ErzBlkAy.Function ErzBlkAy(A() As Blk) As String()
'QLib.Std.MTp_SqyRslt_5_ErzBlkAy.Private Function CvGpAy(A) As Gp()
'QLib.Std.MTp_SqyRslt_5_ErzBlkAy.Private Function ErzExcessPmBlk(A() As Blk) As String()
'QLib.Std.MTp_SqyRslt_5_ErzBlkAy.Private Function ErzExcessSwBlk(A() As Blk) As String()
'QLib.Std.MTp_SqyRslt_5_ErzBlkAy.Function ErzGpAy(GpAy() As Gp, Msg$) As String()
'QLib.Std.MTp_TpBrk.Function TpBrk(Tp$) As TpBrk
'QLib.Std.MTp_TpBrk.Function LnxAyzT1(Ly$(), T1) As Lnx()
'QLib.Std.MTp_TpBrk.Function LnxAyDic(Ly$()) As Dictionary
'QLib.Std.MTp_TpBrk.Function LnxAyDiczT1nn(Ly$(), T1nn) As Dictionary
'QLib.Std.MTp_TpBrk.Function LnxAy(Ly$()) As Lnx()
'QLib.Std.MTp_TpBrk.Function HasMajPfx(Ly$(), MajPfx$) As Boolean
'QLib.Std.MTp_Tp_Gp.Function GpzLy(Ly$()) As Gp
'QLib.Std.MTp_Tp_Gp.Function CvGp(A) As Gp
'QLib.Std.MTp_Tp_Gp.Function LyzGp(A As Gp) As String()
'QLib.Std.MTp_Tp_Gp.Function Gp(A() As Lnx) As Gp
'QLib.Std.MTp_Tp_Gp.Function GpAyzBlkTy(A() As Blk, BlkTy$) As Gp()
'QLib.Std.MTp_Tp_Lin_Cln.Function ClnBrk1(A$(), Ny0) As Variant()
'QLib.Std.MTp_Tp_Lin_Cln.Function ErzT1(Ly$(), T1ss) As String()
'QLib.Std.MTp_Tp_Lin_Cln.Function ClnLin$(Lin)
'QLib.Std.MTp_Tp_Lin_Cln.Function LnxAyzCln(Ly$()) As Lnx()
'QLib.Std.MTp_Tp_Lin_Is.Function IsDDRmkLin(A$) As Boolean
'QLib.Std.MTp_Tp_Lnx.Function CvLnx(A) As Lnx
'QLib.Std.MTp_Tp_Lnx.Function EmpLnx() As Lnx
'QLib.Std.MTp_Tp_Lnx.Function Lnx(Ix, Lin) As Lnx
'QLib.Std.MTp_Tp_Lnx.Sub LnxAsg(A As Lnx, OLin$, OIx%)
'QLib.Std.MTp_Tp_Lnx.Sub BrwLnxAy(A() As Lnx)
'QLib.Std.MTp_Tp_Lnx.Function DupT2AyzLnxAy(A() As Lnx) As String()
'QLib.Std.MTp_Tp_Lnx.Function LyzLnxAy(A() As Lnx) As String()
'QLib.Std.MTp_Tp_Lnx.Function LnxAyeT1Ay(A() As Lnx, T1Ay$()) As Lnx()
'QLib.Std.MTp_Tp_Lnx.Function LyzLnxAyzWithLno(A() As Lnx) As String()
'QLib.Std.MTp_Tp_Lnx.Function ErzLnxAyT1ss(A() As Lnx, T1ss) As String()
'QLib.Std.MTp_Tp_Lnx.Function LnxAywT2(A() As Lnx, T2) As Lnx()
'QLib.Std.MTp_Tp_Lnx.Function LnxAywRmvT1(A() As Lnx, T) As Lnx()
'QLib.Std.MTp_Tp_Lnx.Function LnxRmvT1$(A As Lnx)
'QLib.Std.MTp_Tp_Lnx.Function LnxStr$(A As Lnx)
'QLib.Std.MTp_Tp_Pos.Function TpPos_FmtStr$(A As TpPos)
'QLib.Std.MTp_Tp_RRCC.Function IsEmpRRCC(A As RRCC) As Boolean
'QLib.Std.MTp_Tp_RRCC.Function CvRRCC(A) As RRCC
'QLib.Std.MTp_Tp_RRCC.Function NewRRCC(R1, R2, C1, C2) As RRCC
'QLib.Std.MVb_Align.Function AlignL$(A, W)
'QLib.Std.MVb_Align.Function AlignR$(S, W)
'QLib.Std.MVb_Ay.Sub Asg_ValTo_VarVarible_and_EleOfVariantAy_and_Ap()
'QLib.Std.MVb_Ay.Private Sub WAsg(ParamArray C())
'QLib.Std.MVb_Ay.Sub AsgAp(Ay, ParamArray OAp())
'QLib.Std.MVb_Ay.Sub AyAsgT1AyRestAy(A, OT1Ay$(), ORestAy$())
'QLib.Std.MVb_Ay.Function VcAy(A, Optional Fnn$)
'QLib.Std.MVb_Ay.Function BrwAy(A, Optional Fnn$, Optional UseVc As Boolean)
'QLib.Std.MVb_Ay.Function AyCln(A)
'QLib.Std.MVb_Ay.Function ChkAyDup(A, QMsg$) As String()
'QLib.Std.MVb_Ay.Function AyDupT1(A) As String()
'QLib.Std.MVb_Ay.Function AyEmpChk(A, Msg$) As String()
'QLib.Std.MVb_Ay.Function ChkEqAy(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
'QLib.Std.MVb_Ay.Function AyOfAyAy(AyOfAy)
'QLib.Std.MVb_Ay.Private Sub Z_AyFlat()
'QLib.Std.MVb_Ay.Function AyFlat(AyOfAy)
'QLib.Std.MVb_Ay.Function AyItmCnt%(A, M)
'QLib.Std.MVb_Ay.Function AywLasN(A, N)
'QLib.Std.MVb_Ay.Function LasEle(Ay)
'QLib.Std.MVb_Ay.Function AyMid(A, Fm, Optional L = 0)
'QLib.Std.MVb_Ay.Function AyNPfxStar%(A)
'QLib.Std.MVb_Ay.Function AyNxtNm$(A, Nm$, Optional MaxN% = 99)
'QLib.Std.MVb_Ay.Function ItrzSS(Ssl_or_Sy)
'QLib.Std.MVb_Ay.Function ItrzStr(Str_or_Sy)
'QLib.Std.MVb_Ay.Function LinItr(Lines)
'QLib.Std.MVb_Ay.Function Itr(A)
'QLib.Std.MVb_Ay.Function AyRTrim(A$()) As String()
'QLib.Std.MVb_Ay.Sub ReSumSiN(OAy, N)
'QLib.Std.MVb_Ay.Sub Resz(OAy, U)
'QLib.Std.MVb_Ay.Sub ReSumSiU(OAy, U)
'QLib.Std.MVb_Ay.Function AyReverse(A)
'QLib.Std.MVb_Ay.Function AyReverseI(A)
'QLib.Std.MVb_Ay.Function OyReverse(A)
'QLib.Std.MVb_Ay.Function AyRplMid(Ay, B As FTIx, ByAy)
'QLib.Std.MVb_Ay.Function AyRpl_Star_InEach_Ele(A$(), By) As String()
'QLib.Std.MVb_Ay.Function AyRplT1(A$(), T1$) As String()
'QLib.Std.MVb_Ay.Function AySampLin$(A)
'QLib.Std.MVb_Ay.Function Itm_IsSel(Itm, Ay) As Boolean
'QLib.Std.MVb_Ay.Function SeqCntDicvAy(A) As Dictionary 'The return dic of key=AyEle pointing to 2-Ele-LngAy with Ele-0 as Seq#(0..) and Ele- as Cnt
'QLib.Std.MVb_Ay.Function SqzAyH(Ay) As Variant()
'QLib.Std.MVb_Ay.Function SqzAyV(Ay) As Variant()
'QLib.Std.MVb_Ay.Function AyT1Chd(A, T1) As String()
'QLib.Std.MVb_Ay.Function AyIndent(A, Optional Ident% = 4) As String()
'QLib.Std.MVb_Ay.Function AyTrim(A) As String()
'QLib.Std.MVb_Ay.Function WdtzAy%(A)
'QLib.Std.MVb_Ay.Function AyWrpPad(A, W%) As String() ' Each Itm of Ay[A] is padded to line with WdtzAy(A).  return all padded lines as String()
'QLib.Std.MVb_Ay.Sub WrtAy(A, Ft, Optional OvrWrt As Boolean)
'QLib.Std.MVb_Ay.Function AyZip(A1, A2) As Variant()
'QLib.Std.MVb_Ay.Function AyZip_Ap(A, ParamArray Ap()) As Variant()
'QLib.Std.MVb_Ay.Function AyItmAddAy(Itm, Ay)
'QLib.Std.MVb_Ay.Function SubDrFnySel(Dr(), DrFny$(), SelFF) As Variant()
'QLib.Std.MVb_Ay.Private Sub ZZZ_AyabCzFT()
'QLib.Std.MVb_Ay.Private Sub ZZ_AyAsgAp()
'QLib.Std.MVb_Ay.Private Sub ZZ_ChkEqAy()
'QLib.Std.MVb_Ay.Private Sub ZZ_MaxAy()
'QLib.Std.MVb_Ay.Private Sub ZZ_AyMinus()
'QLib.Std.MVb_Ay.Private Sub ZZ_AyeEmpEleAtEnd()
'QLib.Std.MVb_Ay.Private Sub ZZ_SyzAy()
'QLib.Std.MVb_Ay.Private Sub ZZ_AyTrim()
'QLib.Std.MVb_Ay.Private Sub Z_ChkAyDup()
'QLib.Std.MVb_Ay.Private Sub Z_ChkEqAy()
'QLib.Std.MVb_Ay.Private Sub Z_AyabCzFTIxIx()
'QLib.Std.MVb_Ay.Private Sub Z_HasEleDupEle()
'QLib.Std.MVb_Ay.Private Sub Z_AyInsItm()
'QLib.Std.MVb_Ay.Private Sub Z_AyInsAy()
'QLib.Std.MVb_Ay.Private Sub Z_AyMinus()
'QLib.Std.MVb_Ay.Private Sub Z_SyzAy()
'QLib.Std.MVb_Ay.Private Sub Z_AyTrim()
'QLib.Std.MVb_Ay.Private Sub Z_KKCMiDry()
'QLib.Std.MVb_Ay.Private Sub Z_SubDrFnySel()
'QLib.Std.MVb_Ay.Function CvAy(A) As Variant()
'QLib.Std.MVb_Ay.Function CvAyITM(Itm_or_Ay) As Variant()
'QLib.Std.MVb_Ay.Private Sub Z()
'QLib.Std.MVb_Ay.Private Sub Z_SyPfxSsl()
'QLib.Std.MVb_Ay.Function SyPfxSsl(A) As String()
'QLib.Std.MVb_Ay.Function StrItrzSsl(Ssl)
'QLib.Std.MVb_Ay.Function SySsl(S) As String()
'QLib.Std.MVb_Ay.Function IntSeq(N&, Optional IsFmOne As Boolean) As Integer()
'QLib.Std.MVb_Ay_AB.Function JnAyab(A, B, Optional Sep$) As String()
'QLib.Std.MVb_Ay_AB.Function JnAyabSpc(A, B) As String()
'QLib.Std.MVb_Ay_AB.Function DicAyab(A, B) As Dictionary
'QLib.Std.MVb_Ay_AB.Function FmtAyab(A, B) As String()
'QLib.Std.MVb_Ay_AB.Function LyAyabJnsepForNonEmpB(A, B, Optional Sep$ = " ") As String()
'QLib.Std.MVb_Ay_AB.Sub AsgAyaReSzMax(A, B, OA, OB)
'QLib.Std.MVb_Ay_AB.Sub ReSumSiabMax(OA, OB)
'QLib.Std.MVb_Ay_AB.Sub ThwAyabNE(A, B, Optional N1$ = "Exp", Optional N2$ = "Act")
'QLib.Std.MVb_Ay_BoolAy.Function AndBoolAy(A() As Boolean) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function IsAllFalse(A() As Boolean) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function IsAllTrue(A() As Boolean) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function IsSomTrue(A() As Boolean) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function OrBoolAy(A() As Boolean) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function IsSomFalse(A() As Boolean) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function BoolOp(BoolOpStr) As eBoolOp
'QLib.Std.MVb_Ay_BoolAy.Function IsAndOrStr(A$) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function IsEqNeStr(A$) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function IsVdtBoolOpStr(BoolOpStr$) As Boolean
'QLib.Std.MVb_Ay_BoolAy.Function IfStr$(IfTrue As Boolean, RetStr$)
'QLib.Std.MVb_Ay_BoolAy.Property Get BoolOpSy() As String()
'QLib.Std.MVb_Ay_Emp.Property Get EmpAv() As Variant()
'QLib.Std.MVb_Ay_Emp.Property Get EmpBoolAy() As Boolean()
'QLib.Std.MVb_Ay_Emp.Property Get EmpBytAy() As Byte()
'QLib.Std.MVb_Ay_Emp.Property Get EmpDblAy() As Double()
'QLib.Std.MVb_Ay_Emp.Property Get EmpDicAy() As Dictionary()
'QLib.Std.MVb_Ay_Emp.Property Get EmpDteAy() As Date()
'QLib.Std.MVb_Ay_Emp.Property Get EmpIntAy() As Integer()
'QLib.Std.MVb_Ay_Emp.Property Get EmpLngAy() As Long()
'QLib.Std.MVb_Ay_Emp.Property Get EmpMdAy() As CodeModule
'QLib.Std.MVb_Ay_Emp.Property Get EmpSngAy() As Single()
'QLib.Std.MVb_Ay_Emp.Function EmpSy(Optional Anything) As String()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function AvzAy(Ay) As Variant()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function Av(ParamArray Ap()) As Variant()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function AvzAp(ParamArray Ap()) As Variant()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function SyzApNonBlank(ParamArray Ap()) As String()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function SyzAp(ParamArray Ap()) As String()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function Sy(ParamArray Ap()) As String()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function DteAy(ParamArray Ap()) As Date()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function IntAyzLngAy(LngAy&()) As Integer()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function IntAy(ParamArray Ap()) As Integer()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function LngAy(ParamArray Ap()) As Long()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function SngAy(ParamArray Ap()) As Single()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function SyNoBlank(ParamArray Itm_or_AyAp()) As String()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function IntAyzFmTo(FmInt%, ToInt%) As Integer()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function LngAyzFmTo(FmLng&, ToLng&) As Long()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function SyzAv(Av() As Variant) As String()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function SyzAy(Ay) As String()
'QLib.Std.MVb_Ay_FmTo_Fm_Ap.Function SyzAyNonBlank(Ay) As String()
'QLib.Std.MVb_Ay_FmTo_To_Dic.Function IxDiczAy(Ay) As Dictionary
'QLib.Std.MVb_Ay_FmTo_To_Dic.Function IdCntDiczAy(Ay) As Dictionary
'QLib.Std.MVb_Ay_FmTo_To_Dry.Function DryzAyAddC(Ay, C) As Variant()
'QLib.Std.MVb_Ay_FmTo_To_Dry.Function DryzCAddAy(A, C) As Variant()
'QLib.Std.MVb_Ay_FmTo_To_Dry.Function DryzAyzTyNmVal(Ay) As Variant()
'QLib.Std.MVb_Ay_FmTo_To_Dry.Sub DmpAyzTyNmVal(Ay)
'QLib.Std.MVb_Ay_FmTo_To_Into.Function IntozAy(Into, Ay)
'QLib.Std.MVb_Ay_FmTo_To_Into.Function IntozItrNy(Into, Itr, Ny$())
'QLib.Std.MVb_Ay_FmTo_To_Into.Function IntozItr(Into, Itr)
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyNTerm(Ay, N%) As String()
'QLib.Std.MVb_Ay_Map_Align.Private Function FmtAyNTerm1$(A, W%())
'QLib.Std.MVb_Ay_Map_Align.Private Function WdtAyNTermAy(NTerm%, Ay) As Integer()
'QLib.Std.MVb_Ay_Map_Align.Private Function WdtAyNTermLin(N%, Lin) As Integer()
'QLib.Std.MVb_Ay_Map_Align.Private Function WdtAyab(A%(), B%()) As Integer()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyAtChr(Ay, AtChr$) As String()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyAtDot(A) As String()
'QLib.Std.MVb_Ay_Map_Align.Sub BrwDotLy(DotLy$())
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyDot(DotLy$()) As String()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyDot1(DotLy$()) As String()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyT1(A$()) As String()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyT2(A) As String()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyT3(A$()) As String()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyT4(A$()) As String()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAySamWdt(Ay) As String()
'QLib.Std.MVb_Ay_Map_Align.Function FmtAyR(Ay) As String()
'QLib.Std.MVb_Ay_Map_Align.Private Sub Z_FmtAyT2()
'QLib.Std.MVb_Ay_Map_Align.Private Sub Z_FmtAyT3()
'QLib.Std.MVb_Ay_Map_Align.Private Sub Z()
'QLib.Std.MVb_Ay_Map_Quote.Function AyQuote(A, QuoteStr$) As String()
'QLib.Std.MVb_Ay_Map_Quote.Function AyQuoteDbl(A) As String()
'QLib.Std.MVb_Ay_Map_Quote.Function AyQuoteSng(A) As String()
'QLib.Std.MVb_Ay_Map_Quote.Function AyQuoteSq(A) As String()
'QLib.Std.MVb_Ay_Map_Quote.Function AyQuoteSqIf(Ay) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Private Sub Y(S$, X$)
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmvFstChr(A) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmvFstNonLetter(A) As String() 'Gen:AyXXX
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmvLasChr(A) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmvPfx(A, Pfx) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmvSngQRmk(A$()) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmvSngQuote(A$()) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmvT1(A) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmvTT(A$()) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRpl(Ay, Fm$, By$, Optional Cnt& = 1) As String()
'QLib.Std.MVb_Ay_Map_Rmv.Function AyRmv2Dash(Ay) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakBefDD(A) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakAftDot(A) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakAft(A, Sep$) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakBef(A, Sep$) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakBefDot(A) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakBefOrAll(A, Sep$) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakT1(A) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakT2(A) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakT3(A) As String()
'QLib.Std.MVb_Ay_Map_Tak.Function AyTakBetBkt(Ay) As String()
'QLib.Std.MVb_Ay_Map_Transform.Function AyAddIxPfx(A, Optional BegFm&) As String()
'QLib.Std.MVb_Ay_Map_Transform.Function AyIncEle1(A)
'QLib.Std.MVb_Ay_Map_Transform.Function AyIncEleN(A, N)
'QLib.Std.MVb_Ay_Map_Transform.Function T1Ay(Ay) As String()
'QLib.Std.MVb_Ay_Map_Transform.Function T2Ay(Ay) As String()
'QLib.Std.MVb_Ay_Map_Transform.Function TermAsetzTLinAy(TLinAy$()) As Aset
'QLib.Std.MVb_Ay_Op.Function DashLT1Ay(Ay) As String()
'QLib.Std.MVb_Ay_Op.Function AyEndTrim(A$()) As String()
'QLib.Std.MVb_Ay_Op.Function AyIntersect(A, B)
'QLib.Std.MVb_Ay_Op.Function MinAy(A)
'QLib.Std.MVb_Ay_Op.Function AyMinusAp(Ay, ParamArray Ap())
'QLib.Std.MVb_Ay_Op.Function AyMinus(A, B)
'QLib.Std.MVb_Ay_Op.Function MaxAy(A)
'QLib.Std.MVb_Ay_Op.Function Ny(A) As String()
'QLib.Std.MVb_Ay_Op.Function CvVy(Vy)
'QLib.Std.MVb_Ay_Op.Function CvBytAy(A) As Byte()
'QLib.Std.MVb_Ay_Op.Function CvAv(A) As Variant()
'QLib.Std.MVb_Ay_Op.Function CvSy(A) As String()
'QLib.Std.MVb_Ay_Op.Function SyShow(XX$, Sy$()) As String()
'QLib.Std.MVb_Ay_Op.Private Sub ZZ()
'QLib.Std.MVb_Ay_Op.Private Sub Z()
'QLib.Std.MVb_Ay_Op_Add.Function AyAdd1(A)
'QLib.Std.MVb_Ay_Op_Add.Function SyAdd(A$(), B$()) As String()
'QLib.Std.MVb_Ay_Op_Add.Function AyAdd(A, B)
'QLib.Std.MVb_Ay_Op_Add.Function SyAddSorSyAp(GivenSy$(), ParamArray SorSyAp()) As String()
'QLib.Std.MVb_Ay_Op_Add.Function TyNyzAy(Ay) As String()
'QLib.Std.MVb_Ay_Op_Add.Function SyAddSyAv(Sy$(), SyAv()) As String()
'QLib.Std.MVb_Ay_Op_Add.Function SyAddAp(Sy$(), ParamArray SyAp()) As String()
'QLib.Std.MVb_Ay_Op_Add.Function AyAddAp(Ay, ParamArray Itm_or_AyAp())
'QLib.Std.MVb_Ay_Op_Add.Function AyMap(Ay, MapFun$) As Variant()
'QLib.Std.MVb_Ay_Op_Add.Function DryzAyMap(Ay, Map$) As Variant()
'QLib.Std.MVb_Ay_Op_Add.Function AyAddItm(A, Itm)
'QLib.Std.MVb_Ay_Op_Add.Function AyAddN(A, N)
'QLib.Std.MVb_Ay_Op_Add.Private Sub Z_AyAdd()
'QLib.Std.MVb_Ay_Op_Add.Private Sub ZZ_AyAdd()
'QLib.Std.MVb_Ay_Op_Add.Private Sub ZZ_AyAddPfx()
'QLib.Std.MVb_Ay_Op_Add.Private Sub ZZ_AyAddPfxSfx()
'QLib.Std.MVb_Ay_Op_Add.Function AyTab(A) As String()
'QLib.Std.MVb_Ay_Op_Add.Private Sub ZZ_AyAddSfx()
'QLib.Std.MVb_Ay_Op_Add.Private Sub Z()
'QLib.Std.MVb_Ay_Op_Brk.Function AyabzAyPfx(Ay, Pfx) As AyAB
'QLib.Std.MVb_Ay_Op_Brk.Function AyabzAyEle(Ay, Ele) As AyAB
'QLib.Std.MVb_Ay_Op_Brk.Function AyabCzFT(Ay, FmIx&, ToIx&) As AyABC
'QLib.Std.MVb_Ay_Op_Brk.Function AyabCzFTIx(Ay, B As FTIx) As AyABC
'QLib.Std.MVb_Ay_Op_Cnt.Function CntDryWhGt1zAy(A) As Variant()
'QLib.Std.MVb_Ay_Op_Cnt.Function CntDryzAy(A) As Variant()
'QLib.Std.MVb_Ay_Op_Cnt.Private Sub Z_CntDryzAy()
'QLib.Std.MVb_Ay_Op_Cnt.Function SumSi&(Ay)
'QLib.Std.MVb_Ay_Op_Cnt.Private Sub Z_CntSiLin()
'QLib.Std.MVb_Ay_Op_Cnt.Function CntSiLin$(Ay)
'QLib.Std.MVb_Ay_Op_Cnt.Private Sub Z()
'QLib.Std.MVb_Ay_Op_Has.Function HasObj(Ay, Obj) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function HasDup(Ay) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function HasEle(Ay, Ele) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function HasEleAy(Ay, EleAy) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function HasElezInSomAyzOfAp(ParamArray AyAp()) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function IsSubAy(SubAy, SuperAy) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function IsSuperAy(SuperAy, SubAy) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function ThwNotSuperAy(SuperAy, SubAy) As String()
'QLib.Std.MVb_Ay_Op_Has.Function HasEleAyInSeq(A, B) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function HasEleDupEle(A) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function HasEleNegOne(A) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function HasElePredPXTrue(A, PX$, P) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function HasElePredXPTrue(A, XP$, P) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Function IsAySub(Ay, SubAy) As Boolean
'QLib.Std.MVb_Ay_Op_Has.Private Sub ZZ_HasEleAyInSeq()
'QLib.Std.MVb_Ay_Op_Has.Private Sub ZZ_HasEleDupEle()
'QLib.Std.MVb_Ay_Op_Ins.Function AyInsVVAt(A, V1, V2, Optional At&)
'QLib.Std.MVb_Ay_Op_Ins.Function AyIns(Ay)
'QLib.Std.MVb_Ay_Op_Ins.Function AyInsItm(Ay, Itm, Optional At& = 0)
'QLib.Std.MVb_Ay_Op_Ins.Private Sub Z_AyInsItmAt()
'QLib.Std.MVb_Ay_Op_Ins.Function AyInsAy(A, B)
'QLib.Std.MVb_Ay_Op_Ins.Function AyInsAyAt(A, B, At&)
'QLib.Std.MVb_Ay_Op_Ins.Private Function AyResz(Ay, At&, Optional Cnt = 1)
'QLib.Std.MVb_Ay_Op_Ins.Private Sub Z_AyResz()
'QLib.Std.MVb_Ay_Op_Ins.Private Sub Z()
'QLib.Std.MVb_Ay_Op_Is.Function IsAllEleHasVyzDicKK(A) As Boolean
'QLib.Std.MVb_Ay_Op_Is.Function IsAllEleEqAy(A) As Boolean
'QLib.Std.MVb_Ay_Op_Is.Function IsAllStrAy(A) As Boolean
'QLib.Std.MVb_Ay_Op_Is.Function IsEqSz(A, B) As Boolean
'QLib.Std.MVb_Ay_Op_Is.Function IsEqAy(A, B) As Boolean
'QLib.Std.MVb_Ay_Op_Is.Function IsSamAy(A, B) As Boolean
'QLib.Std.MVb_Ay_Op_Is.Private Sub Z()
'QLib.Std.MVb_Ay_Op_IxAy.Function IxzAy&(Ay, Itm, Optional FmIx& = 0)
'QLib.Std.MVb_Ay_Op_IxAy.Function IxAyU(U&) As Long()
'QLib.Std.MVb_Ay_Op_IxAy.Function IxAy(Ay, SubAy, Optional ThwNotFnd As Boolean) As Long()
'QLib.Std.MVb_Ay_Op_IxAy.Sub AsgItmAyIxay(Ay, IxAy&(), ParamArray OItmAp())
'QLib.Std.MVb_Ay_Op_IxAy.Function IntIxAy(Ay, SubAy) As Integer()
'QLib.Std.MVb_Ay_Op_IxAy.Function IxAyzDup(AyWithDup) As Long()
'QLib.Std.MVb_Ay_Op_Jn.Function JnSpcApNoBlank$(ParamArray Ap())
'QLib.Std.MVb_Ay_Op_Jn.Function JnDollarAp$(ParamArray Ap())
'QLib.Std.MVb_Ay_Op_Jn.Function JnPthSepAp$(ParamArray Ap())
'QLib.Std.MVb_Ay_Op_Jn.Function JnVbarAp$(ParamArray Ap())
'QLib.Std.MVb_Ay_Op_Jn.Function JnVbarApSpc$(ParamArray Ap())
'QLib.Std.MVb_Ay_Op_Jn.Function JnSpcAp$(ParamArray Ap())
'QLib.Std.MVb_Ay_Op_Jn.Function JnSemiColonAp$(ParamArray Ap())
'QLib.Std.MVb_Ay_Op_Jn.Private Sub ZZ()
'QLib.Std.MVb_Ay_Op_Jn.Private Sub Z()
'QLib.Std.MVb_Ay_Op_ReOrder.Function AyReOrd(Ay, SubAy)
'QLib.Std.MVb_Ay_Op_Shf.Private Sub Z_AyShf()
'QLib.Std.MVb_Ay_Op_Shf.Private Sub Z_AyShfItm()
'QLib.Std.MVb_Ay_Op_Shf.Private Sub Z_AyShfItmNy()
'QLib.Std.MVb_Ay_Op_Shf.Function AyShf(OAy)
'QLib.Std.MVb_Ay_Op_Shf.Function AyShfFstNEle(OAy, N)
'QLib.Std.MVb_Ay_Op_Shf.Function AyShfItm$(OAy, Itm)
'QLib.Std.MVb_Ay_Op_Shf.Function AyShfItmEq(A, Itm$) As Variant()
'QLib.Std.MVb_Ay_Op_Shf.Function AyShfItmNy(A$(), ItmNy0) As Variant()
'QLib.Std.MVb_Ay_Op_Shf.Function AyShfQItm$(OAy, QItm)
'QLib.Std.MVb_Ay_Op_Shf.Function AyShfStar(OAy, OItmy$()) As String()
'QLib.Std.MVb_Ay_Op_Shf.Private Sub Z()
'QLib.Std.MVb_Ay_Op_Srt.Function LinesSrt$(A$)
'QLib.Std.MVb_Ay_Op_Srt.Function IsSrtAy(A) As Boolean
'QLib.Std.MVb_Ay_Op_Srt.Function AyQSrt(Ay, Optional IsDes As Boolean)
'QLib.Std.MVb_Ay_Op_Srt.Sub AyQSrtLH(Ay, L&, H&)
'QLib.Std.MVb_Ay_Op_Srt.Function AyQSrtPartition1&(OAy, L&, H&) 'Try mdy
'QLib.Std.MVb_Ay_Op_Srt.Function AyQSrtPartition&(OAy, L&, H&)
'QLib.Std.MVb_Ay_Op_Srt.Private Sub Z_AySrtByAy()
'QLib.Std.MVb_Ay_Op_Srt.Function AySrtByAy(Ay, ByAy)
'QLib.Std.MVb_Ay_Op_Srt.Function AySrt(Ay, Optional Des As Boolean)
'QLib.Std.MVb_Ay_Op_Srt.Private Function AySrt__Ix&(A, V, Des As Boolean)
'QLib.Std.MVb_Ay_Op_Srt.Function IxAyzAySrt(Ay, Optional Des As Boolean) As Long()
'QLib.Std.MVb_Ay_Op_Srt.Private Sub Z_AySrt()
'QLib.Std.MVb_Ay_Op_Srt.Private Function IxAyzAySrt_Ix&(Ix&(), A, V, Des As Boolean)
'QLib.Std.MVb_Ay_Op_Srt.Private Sub Z_IxAyzAySrt()
'QLib.Std.MVb_Ay_Op_Srt.Private Function AySrtInToIxIxAy&(Ix&(), A, V, Des As Boolean)
'QLib.Std.MVb_Ay_Op_Srt.Function DicSrt(A As Dictionary, Optional IsDesc As Boolean) As Dictionary
'QLib.Std.MVb_Ay_Op_Srt.Private Sub ZZ_AySrt()
'QLib.Std.MVb_Ay_Op_Srt.Private Sub ZZ_IxAyzAySrt()
'QLib.Std.MVb_Ay_Op_Srt.Private Sub Z()
'QLib.Std.MVb_Ay_Oy.Function OyAdd(Oy1, Oy2)
'QLib.Std.MVb_Ay_Oy.Sub DoItrMth(Itr, ObjMth$)
'QLib.Std.MVb_Ay_Oy.Sub DoOyMth(Oy, ObjMth$)
'QLib.Std.MVb_Ay_Oy.Function FstOyPEv(Oy, P, V)
'QLib.Std.MVb_Ay_Oy.Function AvOyP(Oy, P) As Variant()
'QLib.Std.MVb_Ay_Oy.Function IntoOyP(Into, Oy, P)
'QLib.Std.MVb_Ay_Oy.Function IntAyOyP(A, P) As Integer()
'QLib.Std.MVb_Ay_Oy.Function SyzOyPrp(A, P) As String()
'QLib.Std.MVb_Ay_Oy.Function OyRmvFstNEle(A, N&)
'QLib.Std.MVb_Ay_Oy.Function OyeNothing(A)
'QLib.Std.MVb_Ay_Oy.Function OywNmPfx(Oy, NmPfx$)
'QLib.Std.MVb_Ay_Oy.Function OywNm(Oy, B As WhNm)
'QLib.Std.MVb_Ay_Oy.Function OywPredXPTrue(A, XP$, P)
'QLib.Std.MVb_Ay_Oy.Function OywPEv(Oy, P, Ev)
'QLib.Std.MVb_Ay_Oy.Function IntAyOywPEvSelP(Oy, P, Ev, SelP) As Integer()
'QLib.Std.MVb_Ay_Oy.Function DryOywPEvSelPP(Oy, P, Ev, SelPP$) As Variant()
'QLib.Std.MVb_Ay_Oy.Function OywPIn(A, P, InAy)
'QLib.Std.MVb_Ay_Oy.Function DryOySelPP(Oy, SelPP$) As Variant()
'QLib.Std.MVb_Ay_Oy.Private Sub ZZ_OyDrs()
'QLib.Std.MVb_Ay_Oy.Private Sub ZZ_OyP_Ay()
'QLib.Std.MVb_Ay_Push.Sub Push(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushAp(O, ParamArray Ap())
'QLib.Std.MVb_Ay_Push.Sub PushAy(O, Ay)
'QLib.Std.MVb_Ay_Push.Sub PushAyNoDup(O, Ay)
'QLib.Std.MVb_Ay_Push.Sub PushDic(O As Dictionary, A As Dictionary)
'QLib.Std.MVb_Ay_Push.Sub PushI(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushIAy(O, MAy)
'QLib.Std.MVb_Ay_Push.Sub PushISomSz(OAy, IAy)
'QLib.Std.MVb_Ay_Push.Sub PushItmAy(O, Itm, Ay)
'QLib.Std.MVb_Ay_Push.Sub PushNoDup(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushNoDupNonBlankStr(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushNoDupAy(O, Ay)
'QLib.Std.MVb_Ay_Push.Sub PushNonBlankStr(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushNonBlankSy(O, Sy$())
'QLib.Std.MVb_Ay_Push.Sub PushNonEmp(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushNonNothing(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushNonZSz(O, Ay)
'QLib.Std.MVb_Ay_Push.Sub PushObjzExlNothing(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushObj(O, M)
'QLib.Std.MVb_Ay_Push.Sub PushObjzItr(O, Itr)
'QLib.Std.MVb_Ay_Push.Sub PushObjAy(O, Oy)
'QLib.Std.MVb_Ay_Push.Sub PushWithSz(O, Ay)
'QLib.Std.MVb_Ay_Push.Function Si&(A)
'QLib.Std.MVb_Ay_Push.Function UB&(A)
'QLib.Std.MVb_Ay_Push.Function Pop(O)
'QLib.Std.MVb_Ay_Push.Function AyReSzU(Ay, U&)
'QLib.Std.MVb_Ay_Run_Do.Sub DoAy(A, FunNm$)
'QLib.Std.MVb_Ay_Run_Do.Sub DoAyabX(Ay, ABX$, A, B)
'QLib.Std.MVb_Ay_Run_Do.Sub DoAyAXB(Ay, AXB$, A, B)
'QLib.Std.MVb_Ay_Run_Do.Sub DoAyPPXP(A, PPXP$, P1, P2, P3)
'QLib.Std.MVb_Ay_Run_Do.Sub DoAyPX(A, PX$, P)
'QLib.Std.MVb_Ay_Run_Do.Sub DoAyXP(A, XP$, P)
'QLib.Std.MVb_Ay_Run_Pred.Function IsAllTruezItrPred_AyPred(A, Pred$) As Boolean
'QLib.Std.MVb_Ay_Run_Pred.Function IsSomeTruezAyPred(A, Pred$) As Boolean
'QLib.Std.MVb_Ay_Run_Pred.Sub AyPredSplitAsg(A, Pred$, OTrueAy, OFalseAy)
'QLib.Std.MVb_Ay_Run_Pred.Function IsAllFalsezAyPred(Ay, Pred$) As Boolean
'QLib.Std.MVb_Ay_Seq.Private Function IntoSeq_FmTo(FmNum, ToNum, OAy)
'QLib.Std.MVb_Ay_Seq.Function CvIntAy(A) As Integer()
'QLib.Std.MVb_Ay_Seq.Function CvLngAy(A) As Long()
'QLib.Std.MVb_Ay_Seq.Function IntSeq_FmTo(FmNum%, ToNum%) As Integer()
'QLib.Std.MVb_Ay_Seq.Function LngSeq_FmTo(FmNum&, ToNum&) As Long()
'QLib.Std.MVb_Ay_Seq.Function IntSeq_0U(U%) As Integer()
'QLib.Std.MVb_Ay_Seq.Function LngSeq_0U(U&) As Long()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyePatn(A, Patn$) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeRe(Ay, Re As RegExp) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeAtCnt(A, Optional At = 0, Optional Cnt = 1)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeDDLin(A) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeDotLin(A) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeEle(A, Ele)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeEleAt(Ay, Optional At = 0, Optional Cnt = 1)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeEleLik(A, Lik$) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeEmpEle(A)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeEmpEleAtEnd(A)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeFmTo(A, FmIx, ToIx)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeFstEle(A)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeFstNEle(A, Optional N = 1)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeFTIx(A, B As FTIx)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeIxSet(Ay, IxSet As Aset)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeIxAy(A, IxAy)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeLasEle(A)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeLasNEle(A, Optional NEle% = 1)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeLik(A, Lik) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeLikAy(A, LikeAy$()) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeLikss(A, Likss$) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeLikssAy(A, LikssAy$()) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeNeg(A)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeNEle(A, Ele, Cnt%)
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeOneTermLin(A) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyePfx(A, ExlPfx$) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Function AyeT1Ay(A, ExlT1Ay0) As String()
'QLib.Std.MVb_Ay_Sub_Exl.Private Sub Z_AyeAtCnt()
'QLib.Std.MVb_Ay_Sub_Exl.Private Sub Z_AyeEmpEleAtEnd()
'QLib.Std.MVb_Ay_Sub_Exl.Private Sub Z_AyeFTIx()
'QLib.Std.MVb_Ay_Sub_Exl.Private Sub Z_AyeFTIx1()
'QLib.Std.MVb_Ay_Sub_Exl.Private Sub Z_AyeIxAy()
'QLib.Std.MVb_Ay_Sub_Exl.Private Sub Z()
'QLib.Std.MVb_Ay_Sub_Exl.Function SyRmvBlank(Ay$()) As String()
'QLib.Std.MVb_Ay_Sub_Fst.Function ShfFstEle(OAy)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEle(Ay)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEleEv(A, V)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEleLik$(A, Lik$)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstElePfx$(PfxAy, Lin$)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEleInAset(Ay, InAset As Aset)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstElePredPX(A, PX$, P)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstElePredXABTrue(Ay, XAB$, A, B)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstElePredXP(A, XP$, P)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEleRmvT1$(Ay, T1Val, Optional IgnCas As Boolean)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEleT1$(Ay, T1Val, Optional IgnCas As Boolean)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEleT2$(A, T2)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEleTT$(A, T1, T2)
'QLib.Std.MVb_Ay_Sub_Fst.Function FstEleRmvTT$(A, T1$, T2$)
'QLib.Std.MVb_Ay_Sub_Wh.Sub AyDupAss(A, Fun$, Optional IgnCas As Boolean)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywIxCnt(Ay, Ix, Cnt)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywBetEle(Ay, FmEle, ToEle)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywDist(A, Optional IgnCas As Boolean)
'QLib.Std.MVb_Ay_Sub_Wh.Private Sub Z_FmtCntDic()
'QLib.Std.MVb_Ay_Sub_Wh.Function SyzDistAy(Ay) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywDistT1(A) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AyAddDicKey(Ay, Dic As Dictionary)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywDup(A, Optional IgnCas As Boolean)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywFmIx(A, FmIx)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywFT(A, FmIx, ToIx)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywFstUEle(Ay, U)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywFstNEle(Ay, N)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywFTIx(A, B As FTIx)
'QLib.Std.MVb_Ay_Sub_Wh.Function IsOutRange(IxAy, U&) As Boolean
'QLib.Std.MVb_Ay_Sub_Wh.Function AywIxAyzMust(Ay, IxAy)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywIxAy(A, IxAy)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywLik(A, Lik) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywLikAy(A, LikeAy$()) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function IsEmpWhNm(A As WhNm) As Boolean
'QLib.Std.MVb_Ay_Sub_Wh.Function AywWhStrPfx(A, WhStr$, Optional NmPfx$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywNmStr(A, WhStr$, Optional NmPfx$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywNm(A, B As WhNm) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AyePfx(A, Pfx$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywObjPred(A, Obj, Pred$)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPatn(A, Patn$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPatnExl(A, Patn$, ExlLikss$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AyPatn_IxAy(A, Patn$) As Long()
'QLib.Std.MVb_Ay_Sub_Wh.Function AyRe_IxAy(A, B As RegExp) As Long()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPfx(A, Pfx$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPred(A, Pred$)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPredFalse(A, Pred$)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPredNot(A, Pred$)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPredXAB(Ay, XAB$, A, B)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPredXABC(Ay, XABC$, A, B, C)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPredXAP(A, PredXAP$, ParamArray Ap())
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPredXP(A, XP$, P)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywPredXPNot(A, XP$, P)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywRe(A, Re As RegExp) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywRmvEle(A, Ele)
'QLib.Std.MVb_Ay_Sub_Wh.Function ItrzAywRmvT1(A, T1$)
'QLib.Std.MVb_Ay_Sub_Wh.Function ItrzSsl(Ssl)
'QLib.Std.MVb_Ay_Sub_Wh.Function ItrzRmvT1(Ay, T1$)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywRmvT1(Ay, T1$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywRmvTT(A, T1$, T2$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywSfx(A, Sfx$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywSingleEle(A)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywSng(A)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywSngEle(A)
'QLib.Std.MVb_Ay_Sub_Wh.Function AywT1(Ay, T1) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywT1InAy(A, Ay$()) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywT1SelRst(A, T1) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywT2EqV(A$(), V) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywTT(A, T1$, T2$) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function AywTTSelRst(A, T1, T2) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Function SywFT(A$(), FmIx, ToIx) As String()
'QLib.Std.MVb_Ay_Sub_Wh.Private Sub ZZ()
'QLib.Std.MVb_Ay_Sub_Wh.Private Sub Z()
'QLib.Std.MVb_Ay_Vbl.Function SyzVbl(Vbl) As String()
'QLib.Std.MVb_Ay_Vbl.Function ItrVbl(Vbl)
'QLib.Std.MVb_Ay_Vbl.Function LineszVbl$(Vbl$)
'QLib.Std.MVb_Ay_Vbl.Function IsVbl(A$) As Boolean
'QLib.Std.MVb_Ay_Vbl.Function IsVblAy(VblAy$()) As Boolean
'QLib.Std.MVb_Ay_Vbl.Function IsVdtVbl(Vbl$) As Boolean
'QLib.Std.MVb_Cml.Private Function MthDotCmlGpAsetzVbe(A As Vbe, Optional WhStr$) As Aset
'QLib.Std.MVb_Cml.Private Sub Z_CmlAset()
'QLib.Std.MVb_Cml.Private Sub Z_ShfCml()
'QLib.Std.MVb_Cml.Function Cml0Ay(Nm) As String()
'QLib.Std.MVb_Cml.Function Cml1Ay(Nm) As String()
'QLib.Std.MVb_Cml.Function CmlAset(Ny$()) As Aset
'QLib.Std.MVb_Cml.Function CmlAy(Nm) As String()
'QLib.Std.MVb_Cml.Function CmlAyzNy(Ny) As String()
'QLib.Std.MVb_Cml.Function CmlGpAy(Nm) As String()
'QLib.Std.MVb_Cml.Function CmlLin$(Nm)
'QLib.Std.MVb_Cml.Function CmlLy(Ny) As String()
'QLib.Std.MVb_Cml.Function CmlQGpAy(Nm) As String()
'QLib.Std.MVb_Cml.Function CmlSetzNy(Ny$()) As Aset
'QLib.Std.MVb_Cml.Function DotCml$(Nm)
'QLib.Std.MVb_Cml.Function DotCmlGp$(Nm) ' = JnQDot . CmpGpAy
'QLib.Std.MVb_Cml.Function DotCmlQGp$(Nm) ' = JnQDot . CmpGp1Ay
'QLib.Std.MVb_Cml.Function FstCml$(S)
'QLib.Std.MVb_Cml.Function FstCmlAy(Ay) As String()
'QLib.Std.MVb_Cml.Function FstCmlx$(S)
'QLib.Std.MVb_Cml.Function FstCmlxAy(Ay) As String()
'QLib.Std.MVb_Cml.Function FstCmlzWithSng$(S)
'QLib.Std.MVb_Cml.Function IsAscCmlChr(A%) As Boolean
'QLib.Std.MVb_Cml.Function IsAscFstCmlChr(A%) As Boolean
'QLib.Std.MVb_Cml.Function IsBRKCml(Cml) As Boolean
'QLib.Std.MVb_Cml.Function IsULCml(Cml) As Boolean
'QLib.Std.MVb_Cml.Function MthDotCmlGpAsetOfVbe(Optional WhStr$) As Aset
'QLib.Std.MVb_Cml.Function MthDotCmlGpAyOfVbe(Optional WhStr$) As String()
'QLib.Std.MVb_Cml.Function MthDotCmlGpAyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MVb_Cml.Function RmvDigSfx$(S)
'QLib.Std.MVb_Cml.Function RmvLDashSfx$(S)
'QLib.Std.MVb_Cml.Function Seg1ErNy() As String()
'QLib.Std.MVb_Cml.Function ShfCml$(OStr$)
'QLib.Std.MVb_Cml.Function ShfCmlAy(S) As String()
'QLib.Std.MVb_Cml.Sub VcMthDotCmlGpAsetOfVbe(Optional WhStr$)
'QLib.Std.MVb_Cml.Sub Z_Cml1Ay()
'QLib.Std.MVb_Cml_Rel.Function CmlRel(Ny$()) As Rel
'QLib.Std.MVb_Const.Property Get SampDb_ShpCst() As Database
'QLib.Std.MVb_Const.Property Get DbEng() As DBEngine
'QLib.Std.MVb_Const.Private Function Db(A) As Dao.Database
'QLib.Std.MVb_Const.Property Get SampCn_DutyDta() As ADODB.Connection
'QLib.Std.MVb_Const.Property Get SampFb$()
'QLib.Std.MVb_Const.Property Get SampDb() As Dao.Database
'QLib.Std.MVb_Const.Property Get SampDb_DutyDta() As Database
'QLib.Std.MVb_Const.Private Sub AAAAA()
'QLib.Std.MVb_Const.Function LinPm(PmStr$) As LinPm
'QLib.Std.MVb_Csv.Function CvCsv$(A)
'QLib.Std.MVb_Csv.Function CsvzDr$(A)
'QLib.Std.MVb_Dft.Function Dft(V, DftV)
'QLib.Std.MVb_Dft.Function DftStr$(Str, Dft)
'QLib.Std.MVb_Dic.Function CvDic(A) As Dictionary
'QLib.Std.MVb_Dic.Function AsetzDicKey(A As Dictionary) As Aset
'QLib.Std.MVb_Dic.Function CvDicAy(A) As Dictionary()
'QLib.Std.MVb_Dic.Function DicAyAdd(A As Dictionary, Dy() As Dictionary) As Dictionary
'QLib.Std.MVb_Dic.Function AddDicKeyPfx(A As Dictionary, Pfx) As Dictionary
'QLib.Std.MVb_Dic.Sub DicAddOrUpd(A As Dictionary, K$, V, Sep$)
'QLib.Std.MVb_Dic.Function DicAllKeyIsNm(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic.Function DicAddKeyPfx(A As Dictionary, KeyPfx$) As Dictionary
'QLib.Std.MVb_Dic.Function DicAyKy(A() As Dictionary) As Variant()
'QLib.Std.MVb_Dic.Function DiczDryOfTwoCol(Dry(), Optional Sep$ = " ") As Dictionary
'QLib.Std.MVb_Dic.Function DicClone(A As Dictionary) As Dictionary
'QLib.Std.MVb_Dic.Function DrDicKy(A As Dictionary, Ky$()) As Variant()
'QLib.Std.MVb_Dic.Function DicFny(InclDicValOptTy As Boolean) As String()
'QLib.Std.MVb_Dic.Function DryzDotAy(DotAy$()) As Variant()
'QLib.Std.MVb_Dic.Function DryzDic(A As Dictionary, Optional InclDicValOptTy As Boolean) As Variant()
'QLib.Std.MVb_Dic.Function DicIntersect(A As Dictionary, B As Dictionary) As Dictionary
'QLib.Std.MVb_Dic.Sub ThwDifDic(A As Dictionary, B As Dictionary, Fun$, Optional N1$ = "A", Optional N2$ = "B")
'QLib.Std.MVb_Dic.Function KeySet(A As Dictionary) As Aset
'QLib.Std.MVb_Dic.Function KeySyzDic(A As Dictionary) As String()
'QLib.Std.MVb_Dic.Function ValOfDicKyJn$(A As Dictionary, Ky, Optional Sep$ = vbCrLf & vbCrLf)
'QLib.Std.MVb_Dic.Function SyzDicKy(Dic As Dictionary, Ky$()) As String()
'QLib.Std.MVb_Dic.Function LineszDic$(A As Dictionary)
'QLib.Std.MVb_Dic.Function FmtDic2(A As Dictionary) As String()
'QLib.Std.MVb_Dic.Function FmtDic2__1(K$, Lines$) As String()
'QLib.Std.MVb_Dic.Function DicMap(A As Dictionary, ValMapFun$) As Dictionary
'QLib.Std.MVb_Dic.Function DicMaxValSz%(A As Dictionary)
'QLib.Std.MVb_Dic.Function DicMge(A As Dictionary, PfxSsl$, ParamArray DicAp()) As Dictionary
'QLib.Std.MVb_Dic.Function DicMinus(A As Dictionary, B As Dictionary) As Dictionary
'QLib.Std.MVb_Dic.Function DicSelIntozAy(A As Dictionary, Ky$()) As Variant()
'QLib.Std.MVb_Dic.Function DicSelIntoSy(A As Dictionary, Ky$()) As String()
'QLib.Std.MVb_Dic.Function SyzDicKey(A As Dictionary) As String()
'QLib.Std.MVb_Dic.Function DiczSwapKV(A As Dictionary) As Dictionary
'QLib.Std.MVb_Dic.Function DicValOpt(A As Dictionary, K)
'QLib.Std.MVb_Dic.Function KeyzLikAyDic_Itm$(Dic As Dictionary, Itm)
'QLib.Std.MVb_Dic.Function KeyzLikssDic_Itm$(A As Dictionary, Itm)
'QLib.Std.MVb_Dic.Private Sub Z_DicMaxValSz()
'QLib.Std.MVb_Dic.Private Sub ZZ()
'QLib.Std.MVb_Dic.Private Sub Z()
'QLib.Std.MVb_Dic.Function WbzNmToLinesDic(A As Dictionary) As Workbook
'QLib.Std.MVb_Dic_Ay.Function DicAdd(A As Dictionary, B As Dictionary) As Dictionary
'QLib.Std.MVb_Dic_Ay.Function DicAyAdd(DicAy() As Dictionary) As Dictionary
'QLib.Std.MVb_Dic_Ay.Function ColDicAyKey(DicAy() As Dictionary, Key) As Variant()
'QLib.Std.MVb_Dic_Ay.Function DicExlKeySet(Dic As Dictionary, ExlKeySet As Aset) As Dictionary
'QLib.Std.MVb_Dic_Cmp.Function FmtCmpDic(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As String()
'QLib.Std.MVb_Dic_Cmp.Function FmtDicCmp(A As DicCmp, Optional ExlSam As Boolean) As String()
'QLib.Std.MVb_Dic_Cmp.Function DicCmp(A As Dictionary, B As Dictionary, Nm1$, Nm2$) As DicCmp
'QLib.Std.MVb_Dic_Cmp.Sub BrwCmpDicAB(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
'QLib.Std.MVb_Dic_Cmp.Function DicSamKV(A As Dictionary, B As Dictionary) As Dictionary
'QLib.Std.MVb_Dic_Cmp.Private Sub AsgADifBDif(A As Dictionary, B As Dictionary,     OADif As Dictionary, OBDif As Dictionary)
'QLib.Std.MVb_Dic_Cmp.Private Function FmtDif(A As Dictionary, B As Dictionary) As String()
'QLib.Std.MVb_Dic_Cmp.Private Function FmtExcess(A As Dictionary, Nm$) As String()
'QLib.Std.MVb_Dic_Cmp.Private Function FmtSam(A As Dictionary) As String()
'QLib.Std.MVb_Dic_Cmp.Private Sub Z_BrwCmpDicAB()
'QLib.Std.MVb_Dic_Cmp.Private Sub Z()
'QLib.Std.MVb_Dic_Def.Function DefDic(Ly$(), KK) As Dictionary
'QLib.Std.MVb_Dic_Fmt.Private Sub Z_BrwDic()
'QLib.Std.MVb_Dic_Fmt.Sub BrwDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional UseVc As Boolean)
'QLib.Std.MVb_Dic_Fmt.Sub DmpDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val")
'QLib.Std.MVb_Dic_Fmt.Function S1S2AyzSyDic(A As Dictionary) As S1S2()
'QLib.Std.MVb_Dic_Fmt.Function FmtDicTit(A As Dictionary, Tit$) As String()
'QLib.Std.MVb_Dic_Fmt.Function FmtDic(A As Dictionary, Optional InclValTy As Boolean, Optional Nm1$ = "Key", Optional Nm2$ = "Val") As String()
'QLib.Std.MVb_Dic_Fmt.Private Function FmtDic2(A As Dictionary) As String()
'QLib.Std.MVb_Dic_Fmt.Function FmtDic1(A As Dictionary, Optional Sep$ = " ") As String()
'QLib.Std.MVb_Dic_Fmt.Function FmtDic3(A As Dictionary) As String()
'QLib.Std.MVb_Dic_Fmt.Private Function FmtDic4(K, Lines) As String()
'QLib.Std.MVb_Dic_GetVal.Function VyzDicKK(Dic As Dictionary, Ky$()) As Variant()
'QLib.Std.MVb_Dic_Has.Function DicHasAllKeyIsNm(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Has.Function DicHasAllValIsStr(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Has.Function DicHasBlankKey(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Has.Function DicHasK(A As Dictionary, K$) As Boolean
'QLib.Std.MVb_Dic_Has.Function DicHasKeyLvs(A As Dictionary, KeyLvs) As Boolean
'QLib.Std.MVb_Dic_Has.Sub DicHasKeyssAss(A As Dictionary, KeySS$)
'QLib.Std.MVb_Dic_Has.Function DicHasKeySsl(A As Dictionary, KeySsl) As Boolean
'QLib.Std.MVb_Dic_Has.Function DicHasKy(A As Dictionary, Ky) As Boolean
'QLib.Std.MVb_Dic_Has.Sub DicHasKyAss(A As Dictionary, Ky)
'QLib.Std.MVb_Dic_Has.Function DicKeysIsAllStr(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Has.Private Sub Z_DicKeysIsAllStr()
'QLib.Std.MVb_Dic_Is.Function IsDiczEmp(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Is.Function IsDiczSy(A) As Boolean
'QLib.Std.MVb_Dic_Is.Function IsDiczLines2(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Is.Function IsDiczLines(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Is.Function IsDiczPrim(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Is.Function IsDiczStr(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Is.Private Function IsDiczLines1(A As Dictionary) As Boolean
'QLib.Std.MVb_Dic_Is.Function DicTy$(A As Dictionary)
'QLib.Std.MVb_Dic_New.Function DiczFt(Ft) As Dictionary
'QLib.Std.MVb_Dic_New.Function NewSyDic(TermLinAy$()) As Dictionary
'QLib.Std.MVb_Dic_New.Sub DicSetKv(O As Dictionary, K, V)
'QLib.Std.MVb_Dic_New.Function FmtDic1(A As Dictionary) As String()
'QLib.Std.MVb_Dic_New.Sub AddDiczNonBlankStr(ODic As Dictionary, K, S$)
'QLib.Std.MVb_Dic_New.Function DiczLines(DicLines, Optional JnSep$ = vbCrLf) As Dictionary
'QLib.Std.MVb_Dic_New.Sub AddDiczApp(OLinesDic As Dictionary, K, StrItm$, Sep$)
'QLib.Std.MVb_Dic_New.Function LyzLinesDicItems(LineszDic As Dictionary) As String()
'QLib.Std.MVb_Dic_New.Function Dic(Ly$(), Optional JnSep$ = vbCrLf) As Dictionary
'QLib.Std.MVb_Dic_New.Function DiczKyVy(Ky, Vy) As Dictionary
'QLib.Std.MVb_Dic_New.Function DiczVbl(Vbl, Optional JnSep$ = vbCrLf) As Dictionary
'QLib.Std.MVb_Dic_Rel.Function CvRel(A) As Rel
'QLib.Std.MVb_Dic_Rel.Property Get EmpRel() As Rel
'QLib.Std.MVb_Dic_Rel.Function IsRel(A) As Boolean
'QLib.Std.MVb_Dic_Rel.Function RelzVbl(RelVbl$) As Rel
'QLib.Std.MVb_Dic_Rel.Function RelzDotLy(DotLy$()) As Rel
'QLib.Std.MVb_Dic_Rel.Function Rel(RelLy$()) As Rel
'QLib.Std.MVb_Dic_Rel.Function RelVbl(Vbl$) As Rel
'QLib.Std.MVb_Dic_Set.Property Get EmpAset() As Aset
'QLib.Std.MVb_Dic_Set.Function CvAset(A) As Aset
'QLib.Std.MVb_Dic_Set.Function IsAset(A) As Boolean
'QLib.Std.MVb_Dic_Set.Function AsetzAp(ParamArray Ap()) As Aset
'QLib.Std.MVb_Dic_Set.Function AsetzItr(Itr) As Aset
'QLib.Std.MVb_Dic_Set.Function AsetzFF(FF) As Aset
'QLib.Std.MVb_Dic_Set.Function AsetzSsl(Ssl) As Aset
'QLib.Std.MVb_Dic_Set.Function AsetzAy(A) As Aset
'QLib.Std.MVb_Dic_SyDic.Sub PushItmzSyDic(A As Dictionary, K, Itm)
'QLib.Std.MVb_Dic_SyDic.Sub ThwNotSyDic(A As Dictionary, Fun$)
'QLib.Std.MVb_Dic_SyDic.Function KeyToLikAyDic_T1LikssLy(TLikssLy$()) As Dictionary
'QLib.Std.MVb_Dic_Wh.Function DicwKK(A As Dictionary, KK) As Dictionary
'QLib.Std.MVb_DotNet_Crypto.Private Sub XXXX()
'QLib.Std.MVb_DotNet_Crypto.Private Sub Z_AsmAy()
'QLib.Std.MVb_DotNet_Crypto.Property Get AsmAy() As Object()
'QLib.Std.MVb_DotNet_Crypto.Sub YY()
'QLib.Std.MVb_DotNet_Crypto.Private Sub XXX()
'QLib.Std.MVb_DotNet_Crypto.Function ToBase64String(rabyt)
'QLib.Std.MVb_DotNet_Crypto.Function ToAscString(rabyt)
'QLib.Std.MVb_DotNet_Crypto.Sub SHA256()
'QLib.Std.MVb_DotNet_Crypto.Sub SHA512()
'QLib.Std.MVb_DotNet_Crypto.Private Sub XXXXX()
'QLib.Std.MVb_DotNet_Crypto.Private Sub ZZ()
'QLib.Std.MVb_DotNet_Crypto.Private Sub Z()
'QLib.Std.MVb_Dta.Function IsEqDry(A(), B()) As Boolean
'QLib.Std.MVb_Dta.Private Sub ZZ()
'QLib.Std.MVb_Dta.Private Sub Z()
'QLib.Std.MVb_Dte.Property Get CurMth() As Byte
'QLib.Std.MVb_Dte.Function NxtMthzM(M As Byte) As Byte
'QLib.Std.MVb_Dte.Function PrvMthzM(M As Byte) As Byte
'QLib.Std.MVb_Dte.Function FstDteOfMth(A As Date) As Date
'QLib.Std.MVb_Dte.Function IsVdtDte(A) As Boolean
'QLib.Std.MVb_Dte.Function LasDteOfMth(A As Date) As Date
'QLib.Std.MVb_Dte.Function NxtMth(A As Date) As Date
'QLib.Std.MVb_Dte.Function PrvDte(A As Date) As Date
'QLib.Std.MVb_Dte.Function YYMM$(A As Date)
'QLib.Std.MVb_Dte.Function FstDtezYYMM(YYMM) As Date
'QLib.Std.MVb_Dte.Function IsVdtYYYYMMDD(A) As Boolean
'QLib.Std.MVb_Dte.Function FstDtezYM(Y As Byte, M As Byte) As Date
'QLib.Std.MVb_Dte.Function LasDtezYM(Y As Byte, M As Byte) As Date
'QLib.Std.MVb_Dte.Function YofNxtMzYM(Y As Byte, M As Byte) As Byte
'QLib.Std.MVb_Dte.Function YofPrvMzYM(Y As Byte, M As Byte) As Byte
'QLib.Std.MVb_Dte.Property Get CurY() As Byte
'QLib.Std.MVb_Dte.Property Get CurYY%()
'QLib.Std.MVb_FmCnt.Function IsEqFTIxAy(A() As FTIx, B() As FTIx) As Boolean
'QLib.Std.MVb_FmCnt.Function FTIxAyIsInOrd(A() As FTIx) As Boolean
'QLib.Std.MVb_FmCnt.Function FTIxAyLinCnt%(A() As FTIx)
'QLib.Std.MVb_FmCnt.Function LyzFTIxAy(A() As FTIx) As String()
'QLib.Std.MVb_FmCnt.Function FTIxIsEq(A As FTIx, B As FTIx) As Boolean
'QLib.Std.MVb_FmCnt.Function FTIxStr$(A As FTIx)
'QLib.Std.MVb_FmCnt.Private Sub ZZ()
'QLib.Std.MVb_FmCnt.Private Sub Z()
'QLib.Std.MVb_Fs_Ffn_AyWh.Function FxAyFfnAy(FfnAy$()) As String()
'QLib.Std.MVb_Fs_Ffn_AyWh.Function FbAyFfnAy(FfnAy$()) As String()
'QLib.Std.MVb_Fs_Ffn_Exist.Sub AsgFfnExistMisAset(OExistFfn As Aset, OMisFfn As Aset, FfnAy$())
'QLib.Std.MVb_Fs_Ffn_Exist.Function FfnExistPair(FfnAy) As SyPair
'QLib.Std.MVb_Fs_Ffn_Exist.Function FfnAywExist(FfnAy) As String()
'QLib.Std.MVb_Fs_Ffn_Exist.Function HasFfn(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Exist.Function ExistFfnAset(FfnAy$()) As Aset
'QLib.Std.MVb_Fs_Ffn_Exist.Function MisFfnAset(FfnAy$()) As Aset
'QLib.Std.MVb_Fs_Ffn_Exist.Function ExistFfnAy(FfnAy$()) As String()
'QLib.Std.MVb_Fs_Ffn_Exist.Function MisFfnAy(FfnAy$()) As String()
'QLib.Std.MVb_Fs_Ffn_Exist.Function IsFfn(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Exist.Function ChkHasFfn(Ffn, Optional FileKind$ = "File") As String()
'QLib.Std.MVb_Fs_Ffn_Exist.Sub ThwNoPth(Pth, Fun$, Optional PthKd$ = "Path")
'QLib.Std.MVb_Fs_Ffn_Exist.Sub ThwNoFfn(Ffn, Fun$, Optional FilKd$)
'QLib.Std.MVb_Fs_Ffn_Exist.Function LyzGpPth(Ffn As Aset) As String()
'QLib.Std.MVb_Fs_Ffn_Exist.Function EnsFfn$(A)
'QLib.Std.MVb_Fs_Ffn_Ext.Function RplExt$(Ffn, NewExt)
'QLib.Std.MVb_Fs_Ffn_Ext.Sub ThwNotExt(S)
'QLib.Std.MVb_Fs_Ffn_Ext.Function IsExt(S) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is.Function IsFxa(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is.Function IsFba(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is.Function IsPjf(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is.Function IsFb(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is.Function IsFx(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is_Sam.Function IsDifFfn(A, B, Optional UseNotEq As Boolean) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is_Sam.Function IsEqFfn(A, B) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is_Sam.Function IsSamFfn(A, B) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is_Sam.Function IsSamSzFfn(A, B) As Boolean
'QLib.Std.MVb_Fs_Ffn_Is_Sam.Function MsgSamFfn(A, B, Si&, Tim$, Optional Msg$) As String()
'QLib.Std.MVb_Fs_Ffn_Is_Sam.Private Sub Z_FfnBlk()
'QLib.Std.MVb_Fs_Ffn_Is_Sam.Function FnoBlk$(Fno%, IBlk)
'QLib.Std.MVb_Fs_Ffn_Is_Sam.Function FfnBlk$(Ffn, IBlk)
'QLib.Std.MVb_Fs_Ffn_Kind.Function TblKd$(Ffn)
'QLib.Std.MVb_Fs_Ffn_Kind.Function FfnKd$(Ffn)
'QLib.Std.MVb_Fs_Ffn_MisEr.Sub ThwMisFfnAy(FfnAy$(), Fun$, Optional FilKind$ = "File")
'QLib.Std.MVb_Fs_Ffn_MisEr.Function MsgzMisFfn(Ffn, Optional FilKind$ = "File") As String()
'QLib.Std.MVb_Fs_Ffn_MisEr.Function MsgzMisFfnAset(MisFfn As Aset, Optional FilKind$ = "file") As String()
'QLib.Std.MVb_Fs_Ffn_MisEr.Function MsgzMisFfnAy(FfnAy$(), Optional FilKind$ = "File") As String()
'QLib.Std.MVb_Fs_Ffn_MisEr.Function ChkMisFfn(Ffn$, Optional FilKind$ = "File") As String()
'QLib.Std.MVb_Fs_Ffn_MisEr.Function ChkMisFfnAy(FfnAy$(), Optional FilKind$ = "File") As String()
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Sub CpyPthzClr(FmPth, ToPth$)
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Sub CpyFilzUp(Ffn)
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Sub CpyFilzToNxtzAy(FfnAy$())
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Function CpyFilzToNxt$(Ffn)
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Sub CpyFilzIfDif(FfnAy_or_Ffn, ToPth$, Optional UseEq As Boolean)
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Sub CpyFilzToFfn(FmFfn, ToFfn$, Optional OvrWrt As Boolean)
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Function CpyFilzToPth$(FfnAy_or_Ffn, ToPth$, Optional OvrWrt As Boolean)
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Sub CpyFilzIfDifzSng(Ffn, ToPth$, Optional UseEq As Boolean)
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Function IsNxtFfn(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Function IsDigStr(S) As Boolean
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Function FfnzNxtFfn$(NxtFfn)
'QLib.Std.MVb_Fs_Ffn_Op_Cpy.Function NxtFfn$(Ffn)
'QLib.Std.MVb_Fs_Ffn_Op_Dlt.Sub DltFfnyAyIf(FfnAy)
'QLib.Std.MVb_Fs_Ffn_Op_Dlt.Sub DltFfn(A)
'QLib.Std.MVb_Fs_Ffn_Op_Dlt.Sub DltFfnIf(Ffn)
'QLib.Std.MVb_Fs_Ffn_Op_Dlt.Function DltFfnIfPrompt(Ffn, Msg$) As Boolean 'Return true if error
'QLib.Std.MVb_Fs_Ffn_Op_Dlt.Function DltFfnDone(Ffn) As Boolean
'QLib.Std.MVb_Fs_Ffn_Op_Mov.Sub MovFilUp(Pth)
'QLib.Std.MVb_Fs_Ffn_Op_Mov.Sub MovFfn(Ffn$, ToPth$)
'QLib.Std.MVb_Fs_Ffn_Prp.Function FfnDte(Ffn) As Date
'QLib.Std.MVb_Fs_Ffn_Prp.Function FfnSz&(Ffn)
'QLib.Std.MVb_Fs_Ffn_Prp.Function FfnFdr$(Ffn)
'QLib.Std.MVb_Fs_Ffn_Prp.Function TimFfn(Ffn) As Date
'QLib.Std.MVb_Fs_Ffn_Prp.Function SzDotDTimFfn$(A)
'QLib.Std.MVb_Fs_Ffn_Prp.Sub AsgTimFfnSz(A$, OTim As Date, OSz&)
'QLib.Std.MVb_Fs_Ffn_Prp.Function FfnTimStr$(A)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function FdrFfn$(A)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function CutPth$(Ffn)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function Fn$(Ffn)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function FfnUp$(Ffn)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function Fnn$(Ffn)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function RmvExt$(A)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function Ext$(A)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function FfnPth$(Ffn)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function PthUp$(Pth, NUp%)
'QLib.Std.MVb_Fs_Ffn_SubPart.Function Pth$(Ffn)
'QLib.Std.MVb_Fs_Ffn_SubPart_Add.Function FfnAddTimSfx$(Ffn)
'QLib.Std.MVb_Fs_Ffn_SubPart_Add.Function FfnAddFnPfx$(A$, Pfx$)
'QLib.Std.MVb_Fs_Ffn_SubPart_Add.Function FfnAddFnSfx$(Ffn, Sfx$)
'QLib.Std.MVb_Fs_Ft.Sub BrwFt(Ft, Optional UseVc As Boolean)
'QLib.Std.MVb_Fs_Ft.Function FtLines$(A)
'QLib.Std.MVb_Fs_Ft.Function FtLy(A) As String()
'QLib.Std.MVb_Fs_Ft.Function FnoRnd128%(Ffn)
'QLib.Std.MVb_Fs_Ft.Function FnoRnd%(Ffn, RecLen%)
'QLib.Std.MVb_Fs_Ft.Function FnoApp%(A)
'QLib.Std.MVb_Fs_Ft.Function FnoInp%(Ft)
'QLib.Std.MVb_Fs_Ft.Function FnoOup%(Ft)
'QLib.Std.MVb_Fs_Ft.Sub RmvFst4LinesFt(Ft)
'QLib.Std.MVb_Fs_Inst.Function FfnInst$(Ffn)
'QLib.Std.MVb_Fs_Inst.Function PthInst$(Pth)
'QLib.Std.MVb_Fs_Inst.Function CrtPthzInst$(Pth)
'QLib.Std.MVb_Fs_Inst.Function IsInstFfn(Ffn) As Boolean
'QLib.Std.MVb_Fs_Inst.Function IsInstFdr(Fdr$) As Boolean
'QLib.Std.MVb_Fs_Pth.Private Function AddFdrzOne$(Pth, Fdr)
'QLib.Std.MVb_Fs_Pth.Function AddFdrEns$(Pth, ParamArray FdrAp())
'QLib.Std.MVb_Fs_Pth.Private Function AddFdrAv$(Pth, FdrAv())
'QLib.Std.MVb_Fs_Pth.Function AddFdr$(Pth, ParamArray FdrAp())
'QLib.Std.MVb_Fs_Pth.Function MsgzFfnAlreadyLoaded(Ffn$, FilKind$, LTimStr$) As String()
'QLib.Std.MVb_Fs_Pth.Function IsEmpPth(Pth) As Boolean
'QLib.Std.MVb_Fs_Pth.Function PthAddPfx$(Pth, Pfx)
'QLib.Std.MVb_Fs_Pth.Function HitFilAtr(A As VbFileAttribute, Wh As VbFileAttribute) As Boolean
'QLib.Std.MVb_Fs_Pth.Function FdrzFfn$(Ffn)
'QLib.Std.MVb_Fs_Pth.Function Fdr$(Pth)
'QLib.Std.MVb_Fs_Pth.Sub ThwNotFdr(A)
'QLib.Std.MVb_Fs_Pth_Exist.Function PthEns$(Pth)
'QLib.Std.MVb_Fs_Pth_Exist.Function PthEnsAll$(A$)
'QLib.Std.MVb_Fs_Pth_Exist.Function IsPth(Pth) As Boolean
'QLib.Std.MVb_Fs_Pth_Exist.Function HasFdr(Pth, Fdr) As Boolean
'QLib.Std.MVb_Fs_Pth_Exist.Sub ThwNotPth(Pth)
'QLib.Std.MVb_Fs_Pth_Exist.Function HasFilPth(Pth) As Boolean
'QLib.Std.MVb_Fs_Pth_Exist.Function HasPth(Pth) As Boolean
'QLib.Std.MVb_Fs_Pth_Exist.Function HasSubFdr(Pth) As Boolean
'QLib.Std.MVb_Fs_Pth_Exist.Sub ThwNotHasPth(Pth, Fun$)
'QLib.Std.MVb_Fs_Pth_Mbr.Function DirzPth$(Pth)
'QLib.Std.MVb_Fs_Pth_Mbr.Function FdrAyz(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Function EntAy(Pth) As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Function FdrAy(Pth) As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Function FfnItr(Pth)
'QLib.Std.MVb_Fs_Pth_Mbr.Function SubPthAy(Pth) As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Function SubPthAyz(Pth, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Sub AsgEnt(OFdrAy$(), OFnAy$(), Pth$)
'QLib.Std.MVb_Fs_Pth_Mbr.Function FnnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Function FnAy(Pth, Optional Spec$ = "*.*") As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Function FxAy(Pth) As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Function FfnAy(Pth, Optional Spec$ = "*.*") As String()
'QLib.Std.MVb_Fs_Pth_Mbr.Private Sub Z_SubPthAy()
'QLib.Std.MVb_Fs_Pth_Mbr.Private Sub ZZ_FxAy()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Function EmpPthAyR(Pth) As String()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Function EntAyR(Pth, Optional FilSpec$ = "*.*") As String()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Private Sub EntAyR1(Pth)
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Private Sub Z_FfnAyR()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Function FfnAyR(Pth, Optional Spec$ = "*.*") As String()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Private Sub FfnAyR1(Pth)
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Private Sub ZZ_EntAyR()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Private Sub Z_EmpPthAyR()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Private Sub Z_EntAy()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Private Sub Z_RmvEmpPthR()
'QLib.Std.MVb_Fs_Pth_Mbr_R_FfnAy.Private Sub Z()
'QLib.Std.MVb_Fs_Pth_Mbr_R_SubPthAy.Function SubPthAyR(Pth) As String()
'QLib.Std.MVb_Fs_Pth_Mbr_R_SubPthAy.Private Sub SubPthAyRz(Pth)
'QLib.Std.MVb_Fs_Pth_Op_Brw.Sub BrwPthVC(Pth)
'QLib.Std.MVb_Fs_Pth_Op_Brw.Sub BrwPth(Pth)
'QLib.Std.MVb_Fs_Pth_Op_Clr.Sub ClrPth(Pth)
'QLib.Std.MVb_Fs_Pth_Op_Clr.Private Sub Z_ClrPthFil()
'QLib.Std.MVb_Fs_Pth_Op_Clr.Sub ClrPthFil(Pth)
'QLib.Std.MVb_Fs_Pth_Op_Ren.Sub RenPthAddPfx(Pth, Pfx)
'QLib.Std.MVb_Fs_Pth_Op_Ren.Sub RenPth(Pth, NewPth)
'QLib.Std.MVb_Fs_Pth_Op_Rmv.Private Sub Z_RmvEmpPthR()
'QLib.Std.MVb_Fs_Pth_Op_Rmv.Sub RmvEmpPthR(Pth)
'QLib.Std.MVb_Fs_Pth_Op_Rmv.Sub RmvEmpSubDir(Pth)
'QLib.Std.MVb_Fs_Pth_Op_Rmv.Sub RmvPthIfEmp(Pth)
'QLib.Std.MVb_Fs_Pth_Op_Rmv.Private Sub ZZ_RmvEmpSubDir()
'QLib.Std.MVb_Fs_Pth_Rel.Function ParPth$(Pth)
'QLib.Std.MVb_Fs_Pth_Rel.Function ParFdr$(Pth)
'QLib.Std.MVb_Fs_Pth_Rel.Function ParPthN$(Pth, UpN%)
'QLib.Std.MVb_Fs_Pth_Rel_Sibling.Function HasSiblingFdr(Pth, Fdr) As Boolean
'QLib.Std.MVb_Fs_Pth_Rel_Sibling.Function SiblingPth$(Pth, SiblingFdr)
'QLib.Std.MVb_Fs_Pth_Sfx.Function HasPthSfx(A) As Boolean
'QLib.Std.MVb_Fs_Pth_Sfx.Function PthEnsSfx$(A)
'QLib.Std.MVb_Fs_Pth_Sfx.Function PthRmvSfx$(Pth)
'QLib.Std.MVb_Fs_Sel.Function FfnSel$(Ffn, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
'QLib.Std.MVb_Fs_Sel.Function PthSel$(Pth, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
'QLib.Std.MVb_Fs_Sel.Private Sub Z_PthSel()
'QLib.Std.MVb_Fs_Sel.Private Sub Z()
'QLib.Std.MVb_Fs_Tmp.Function TmpCmd$(Optional Fdr$, Optional Fnn$)
'QLib.Std.MVb_Fs_Tmp.Function TmpFcsv$(Optional Fdr$, Optional Fnn$)
'QLib.Std.MVb_Fs_Tmp.Function TmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
'QLib.Std.MVb_Fs_Tmp.Function TmpFt$(Optional Fdr$, Optional Fnn$)
'QLib.Std.MVb_Fs_Tmp.Function TmpFx$(Optional Fdr$, Optional Fnn$)
'QLib.Std.MVb_Fs_Tmp.Function TmpFxm$(Optional Fdr$, Optional Fnn0$)
'QLib.Std.MVb_Fs_Tmp.Property Get TmpRoot$()
'QLib.Std.MVb_Fs_Tmp.Property Get TmpHom$()
'QLib.Std.MVb_Fs_Tmp.Sub BrwTmpHom()
'QLib.Std.MVb_Fs_Tmp.Function TmpNmzWithSfx$(Optional Pfx$ = "N")
'QLib.Std.MVb_Fs_Tmp.Function TmpNm$(Optional Pfx$ = "N")
'QLib.Std.MVb_Fs_Tmp.Function TmpFdr$(Fdr$)
'QLib.Std.MVb_Fs_Tmp.Property Get TmpPth$()
'QLib.Std.MVb_Fs_Tmp.Sub TmpBrwPth()
'QLib.Std.MVb_Fs_Tmp.Property Get TmpPthFix$()
'QLib.Std.MVb_Fs_Tmp.Property Get TmpPthHom$()
'QLib.Std.MVb_FTIx.Function FTIx_HasU(A As FTIx, U&) As Boolean
'QLib.Std.MVb_FTIx.Sub AssBet(Fun$, V, FmV, ToV)
'QLib.Std.MVb_FTIx.Function FTIxLinCnt%(A As FTIx)
'QLib.Std.MVb_FTIx.Function EmpFTIx() As FTIx
'QLib.Std.MVb_FTIx.Function FTIxzIxCnt(FmIx, Cnt) As FTIx
'QLib.Std.MVb_FTIx.Function FTIx(FmIx, ToIx) As FTIx
'QLib.Std.MVb_FTIx.Function CvFTIx(A) As FTIx
'QLib.Std.MVb_Fun.Sub Swap(OA, OB)
'QLib.Std.MVb_Fun.Sub Asg(Fm, OTo)
'QLib.Std.MVb_Fun.Function InStrWiIthSubStr&(S, SubStr, Optional Ith% = 1)
'QLib.Std.MVb_Fun.Function InStrN&(S, SubStr, Optional N% = 1)
'QLib.Std.MVb_Fun.Function CvNothing(A)
'QLib.Std.MVb_Fun.Private Sub Z_InStrN()
'QLib.Std.MVb_Fun.Function Max(ParamArray Ap())
'QLib.Std.MVb_Fun.Function MaxVbTy(A As VbVarType, B As VbVarType) As VbVarType
'QLib.Std.MVb_Fun.Function CanCvLng(A) As Boolean
'QLib.Std.MVb_Fun.Function Min(ParamArray A())
'QLib.Std.MVb_Fun.Sub SndKeys(A$)
'QLib.Std.MVb_Fun.Function NDig%(N&)
'QLib.Std.MVb_Fun.Sub Vc(A, Optional Fnn$)
'QLib.Std.MVb_Fun.Sub Brw(A, Optional Fnn$, Optional UseVc As Boolean)
'QLib.Std.MVb_Fun.Function Mch(Re As RegExp, S) As MatchCollection
'QLib.Std.MVb_Fun.Function RegExp(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
'QLib.Std.MVb_Fun.Private Sub Z()
'QLib.Std.MVb_Fun.Function SumLngAy@(A&())
'QLib.Std.MVb_Fun.Private Sub Z_SumLngAy()
'QLib.Std.MVb_Fun.Function AscCntAy(S) As Long()
'QLib.Std.MVb_Fun.Function RemBlkSz&(N&, BlkSz%)
'QLib.Std.MVb_Fun.Function NBlk&(N&, BlkSz%)
'QLib.Std.MVb_Hit.Function HitPfxAy(A, PfxAy) As Boolean
'QLib.Std.MVb_Hit.Function HitPfxAp(A, ParamArray PfxAp()) As Boolean
'QLib.Std.MVb_Hit.Function HitPfxSpc(A, Pfx) As Boolean
'QLib.Std.MVb_Hit.Function HitAyElePfx(Ay, ElePfx) As Boolean
'QLib.Std.MVb_Hit.Function Has2T(Lin, T1, T2) As Boolean
'QLib.Std.MVb_Hit.Function Has3T(Lin, T1, T2, T3) As Boolean
'QLib.Std.MVb_Hit.Function Has1T(Lin, T1) As Boolean
'QLib.Std.MVb_Hit.Function HasT2(Lin, T2) As Boolean
'QLib.Std.MVb_Hit.Function HitLikss(S, Likss) As Boolean
'QLib.Std.MVb_Hit.Function HitLikAy(S, LikeAy$()) As Boolean
'QLib.Std.MVb_Hit.Function HitAv(A, Av()) As Boolean
'QLib.Std.MVb_Hit.Function HitAp(V, ParamArray Ap()) As Boolean
'QLib.Std.MVb_Hit.Function HitNmStr(V, WhStr$, Optional NmPfx$) As Boolean
'QLib.Std.MVb_Hit.Function HitNm(V, B As WhNm) As Boolean
'QLib.Std.MVb_Hit.Function HitAy(V, Ay) As Boolean
'QLib.Std.MVb_Hit.Private Sub Z_HitPatn()
'QLib.Std.MVb_Hit.Function HitPatn(A, Patn) As Boolean
'QLib.Std.MVb_Hit.Private Sub Z()
'QLib.Std.MVb_Ide_Win.Property Get CdWinAy() As Vbide.Window()
'QLib.Std.MVb_Ide_Win.Sub ClrWinzImm()
'QLib.Std.MVb_Ide_Win.Sub ClsWinzWin(W As Vbide.Window)
'QLib.Std.MVb_Ide_Win.Sub ClsWin()
'QLib.Std.MVb_Ide_Win.Sub ClsWinzExlWin(A As Vbide.Window)
'QLib.Std.MVb_Ide_Win.Sub ClsWinzExl(ParamArray ExlWinAp())
'QLib.Std.MVb_Ide_Win.Sub SetWinVisOpt(A, Vis As Boolean)
'QLib.Std.MVb_Ide_Win.Sub ClsWinOpt(A)
'QLib.Std.MVb_Ide_Win.Sub ShwWinOpt(A)
'QLib.Std.MVb_Ide_Win.Function IsWin(A) As Boolean
'QLib.Std.MVb_Ide_Win.Sub ClsWinzExlMd(ExlMdNm$)
'QLib.Std.MVb_Ide_Win.Sub ClsWinzImm()
'QLib.Std.MVb_Ide_Win.Sub ClsWinzExlImm()
'QLib.Std.MVb_Ide_Win.Property Get CurCdWin() As Vbide.Window
'QLib.Std.MVb_Ide_Win.Private Property Get CurVbe() As Vbe
'QLib.Std.MVb_Ide_Win.Property Get CurWin() As Vbide.Window
'QLib.Std.MVb_Ide_Win.Function CvWinAy(A) As Vbide.Window()
'QLib.Std.MVb_Ide_Win.Property Get EmpWinAy() As Vbide.Window()
'QLib.Std.MVb_Ide_Win.Property Get WinzImm() As Vbide.Window
'QLib.Std.MVb_Ide_Win.Property Get WinzLcl() As Vbide.Window
'QLib.Std.MVb_Ide_Win.Function WinzMdNm(MdNm) As Vbide.Window
'QLib.Std.MVb_Ide_Win.Function WinzMd(A As CodeModule) As Vbide.Window
'QLib.Std.MVb_Ide_Win.Private Function MdzPj(A As VBProject, Nm) As CodeModule
'QLib.Std.MVb_Ide_Win.Sub ShwDbg()
'QLib.Std.MVb_Ide_Win.Sub JmpNxtStmt()
'QLib.Std.MVb_Ide_Win.Property Get VisWinCnt&()
'QLib.Std.MVb_Ide_Win.Function CvWin(A) As Vbide.Window
'QLib.Std.MVb_Ide_Win.Function CvWinOpt(A) As Vbide.Window
'QLib.Std.MVb_Ide_Win.Sub ClrWin(A As Vbide.Window)
'QLib.Std.MVb_Ide_Win.Property Get WinCnt&()
'QLib.Std.MVb_Ide_Win.Function MdNmCdWin$(CdWin As Vbide.Window)
'QLib.Std.MVb_Ide_Win.Property Get WinNy() As String()
'QLib.Std.MVb_Ide_Win.Function FstWinTy(A As vbext_WindowType) As Vbide.Window
'QLib.Std.MVb_Ide_Win.Function WinAyWinTy(T As vbext_WindowType) As Vbide.Window()
'QLib.Std.MVb_Ide_Win.Function SetVisWin(A As Vbide.Window) As Vbide.Window
'QLib.Std.MVb_Ide_Win.Private Sub Z_Md()
'QLib.Std.MVb_Ide_Win.Private Sub ZZ()
'QLib.Std.MVb_Ide_Win.Private Sub Z()
'QLib.Std.MVb_Ide_Win.Function CdPnezCmpNm(CmpNm) As CodePane
'QLib.Std.MVb_Ide_Win.Function WinzCmpNm(CmpNm) As Vbide.Window
'QLib.Std.MVb_Ide_Win.Sub ShwCmp(CmpNm)
'QLib.Std.MVb_Ide_Win.Sub ClsWinzExlCmpOoImm(MdNm$)
'QLib.Std.MVb_Ide_Win.Function WinAyAv(Av()) As Vbide.Window()
'QLib.Std.MVb_Ide_Win.Sub ClsWinzExlWinAp(ParamArray WinAp())
'QLib.Std.MVb_Ide_Win.Sub ShwWin(A As Vbide.Window)
'QLib.Std.MVb_Ide_Win.Property Get VisWinAy() As Vbide.Window()
'QLib.Std.MVb_Is.Function IsBet(V, A, B) As Boolean
'QLib.Std.MVb_Is.Function IsEmp(A) As Boolean
'QLib.Std.MVb_Is.Function IsNBet(V, A, B) As Boolean
'QLib.Std.MVb_Is.Function IsSqBktQuoted(A) As Boolean
'QLib.Std.MVb_Is.Function Limit(V, A, B)
'QLib.Std.MVb_Is_Asc.Function IsAscDig(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Property Get AscAyzNonPrt() As Integer()
'QLib.Std.MVb_Is_Asc.Function IsAscPrintablezStrI(S, I) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscNonPrt(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscPrintable(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscDigit(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscFstNmChr(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscLDash(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscLCas(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscLetterDig(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscLetter(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscNmChr(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscPun(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscPun1(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscPun2(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscPun3(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscPun4(A%) As Boolean
'QLib.Std.MVb_Is_Asc.Function IsAscUCas(A%) As Boolean
'QLib.Std.MVb_Is_Var.Function IsAv(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsAyDic(A As Dictionary) As Boolean
'QLib.Std.MVb_Is_Var.Function IsAyOfAy(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsBool(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsByt(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsBytAy(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsDic(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsDigit(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsDte(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsEq(A, B) As Boolean
'QLib.Std.MVb_Is_Var.Function IsEqDic(A As Dictionary, B As Dictionary) As Boolean
'QLib.Std.MVb_Is_Var.Function IsEqTy(A, B) As Boolean
'QLib.Std.MVb_Is_Var.Function IsInt(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsIntAy(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsItr(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsLetter(A$) As Boolean
'QLib.Std.MVb_Is_Var.Function IsLines(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsLinesAy(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsLng(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsLngAy(V) As Boolean
'QLib.Std.MVb_Is_Var.Function IsNe(A, B) As Boolean
'QLib.Std.MVb_Is_Var.Function IsNoLinMd(A As CodeModule) As Boolean
'QLib.Std.MVb_Is_Var.Function IsNonBlankStr(V) As Boolean
'QLib.Std.MVb_Is_Var.Function IsNothing(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsObjAy(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsPrim(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsPun(A$) As Boolean
'QLib.Std.MVb_Is_Var.Function IsQuoted(A, Q1$, Optional ByVal Q2$) As Boolean
'QLib.Std.MVb_Is_Var.Function IsSngQRmk(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsSngQuoted(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsSomething(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsNeedQuote(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsStr(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsStrAy(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsEmpSy(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsSy(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsTgl(A) As Boolean
'QLib.Std.MVb_Is_Var.Function IsVbTyNum(A As VbVarType) As Boolean
'QLib.Std.MVb_Is_Var.Function IsVdtLyDicStr(LyDicStr$) As Boolean
'QLib.Std.MVb_Is_Var.Function IsWhiteChr(A) As Boolean
'QLib.Std.MVb_Is_Var.Private Sub ZIsSy()
'QLib.Std.MVb_Is_Var.Private Sub ZZ_IsStrAy()
'QLib.Std.MVb_Is_Var.Private Sub Z_IsVdtLyDicStr()
'QLib.Std.MVb_Is_Var.Private Sub Z()
'QLib.Std.MVb_Is_Var.Function IsAllBlankSy(A$()) As Boolean
'QLib.Std.MVb_Itp.Function IntozItp(OInto, Itr, P) As String()
'QLib.Std.MVb_Itp.Function SyzItp(Itr, P) As String()
'QLib.Std.MVb_Itr.Function ObjVyzItr(Itr) As Variant()
'QLib.Std.MVb_Itr.Function AvzItr(Itr) As Variant()
'QLib.Std.MVb_Itr.Function ItrAddSfx(Itr, Sfx$) As String()
'QLib.Std.MVb_Itr.Function ItrAddPfx(A, Pfx$) As String()
'QLib.Std.MVb_Itr.Function ItrClnAy(A)
'QLib.Std.MVb_Itr.Function NItrPrpTrue(A, BoolPrpNm)
'QLib.Std.MVb_Itr.Sub DoItrFun(A, DoFun$)
'QLib.Std.MVb_Itr.Sub DoItrFunPX(Itr, PX$, P)
'QLib.Std.MVb_Itr.Sub DoItrFunXP(A, XP$, P)
'QLib.Std.MVb_Itr.Function FstItr(A)
'QLib.Std.MVb_Itr.Function FstItrPredXP(A, XP$, P)
'QLib.Std.MVb_Itr.Function FstItrNm(Itr, Nm) ' Return first element in Itr-A with name eq Nm
'QLib.Std.MVb_Itr.Function FstItrPEv(Itr, P, Ev) 'Return first element in Itr-A with its Prp-P eq to V
'QLib.Std.MVb_Itr.Function FstItn(Itr, Nm) 'Return first element in Itr with its PrpNm=Nm being true
'QLib.Std.MVb_Itr.Function FstItrTrueP(Itr, TruePrp) 'Return first element in Itr with its Prp-P being true
'QLib.Std.MVb_Itr.Function HasItrTrueP(Itr, TruePrp) As Boolean
'QLib.Std.MVb_Itr.Function HasItn(Itr, Nm) As Boolean
'QLib.Std.MVb_Itr.Function HasItrPEv(A, P, Ev) As Boolean
'QLib.Std.MVb_Itr.Function HasItrTruePrp(A, P) As Boolean
'QLib.Std.MVb_Itr.Function IsEqNmItr(A, B)
'QLib.Std.MVb_Itr.Function AvzItrMap(Itr, Map$) As Variant()
'QLib.Std.MVb_Itr.Function IntozAyMap(OInto, Ay, Map$)
'QLib.Std.MVb_Itr.Function IntozAvzItrMap(OInto, Itr, Map$)
'QLib.Std.MVb_Itr.Function SyzAvzItrMap(Itr, Map$) As String()
'QLib.Std.MVb_Itr.Function MaxItrPrp(A, P)
'QLib.Std.MVb_Itr.Function NyOy(A) As String()
'QLib.Std.MVb_Itr.Function ItnPEv(Itr, WhPrp, Ev) As String()
'QLib.Std.MVb_Itr.Function SyPrp(Itr, P) As String()
'QLib.Std.MVb_Itr.Function NyzOy(Oy) As String()
'QLib.Std.MVb_Itr.Function Itn(Itr) As String()
'QLib.Std.MVb_Itr.Function IsAllFalsezItrPred(Itr, Pred$) As Boolean
'QLib.Std.MVb_Itr.Function IsAllTruezItrPred(Itr, Pred$) As Boolean
'QLib.Std.MVb_Itr.Function IsSomFalsezItrPred(Itr, Pred$) As Boolean
'QLib.Std.MVb_Itr.Function IsSomTruezItrPred(Itr, Pred$) As Boolean
'QLib.Std.MVb_Itr.Function SyzItrPrp(Itr, P) As String()
'QLib.Std.MVb_Itr.Function AvzItrPrp(Itr, P) As Variant()
'QLib.Std.MVb_Itr.Function IntozItrPrpTrue(Into, Itr, P)
'QLib.Std.MVb_Itr.Function IntozItrPEv(Into, Itr, P, Ev)
'QLib.Std.MVb_Itr.Function IntozItrPrp(Into, Itr, P)
'QLib.Std.MVb_Itr.Function AvItrValue(A) As Variant()
'QLib.Std.MVb_Itr.Function ItrPrp_WhTrue_Into(A, P, Into)
'QLib.Std.MVb_Itr.Function ItrwPrpEqval(A, Prp, EqVal)
'QLib.Std.MVb_Itr.Function ItrwPrpTrue(A, P)
'QLib.Std.MVb_Itr.Function ItrwNm(A, B As WhNm)
'QLib.Std.MVb_Itr.Private Sub ZZ()
'QLib.Std.MVb_Itr.Private Sub Z()
'QLib.Std.MVb_Itr.Function NItrPEv&(Itr, P, Ev)
'QLib.Std.MVb_Itr_Is.Function IsItrzLines(Itr) As Boolean
'QLib.Std.MVb_Itr_Is.Function IsItrzStr(Itr) As Boolean
'QLib.Std.MVb_Itr_Is.Function IsItrzPrim(Itr) As Boolean
'QLib.Std.MVb_Itr_Is.Function IsItrzNm(Itr) As Boolean
'QLib.Std.MVb_Itr_Is.Function IsItrzSy(Itr) As Boolean
'QLib.Std.MVb_JnSplit_Jn.Function Jn$(A, Optional Sep$ = "")
'QLib.Std.MVb_JnSplit_Jn.Function QuoteBktJnComma$(Ay)
'QLib.Std.MVb_JnSplit_Jn.Function JnComma$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnCommaCrLf$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnAnd$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnCommaSpc$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnCrLf$(Ay)
'QLib.Std.MVb_JnSplit_Jn.Function JnDblCrLf$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnDotAp$(ParamArray Ap())
'QLib.Std.MVb_JnSplit_Jn.Function JnQDot$(Ay) 'JnQDot = QuoteDot . JnDot
'QLib.Std.MVb_JnSplit_Jn.Function JnDot$(Ay)
'QLib.Std.MVb_JnSplit_Jn.Function JnDollar$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnDblDollar$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnPthSep$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnQDblComma$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnQDblSpc$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnQSngComma$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnQSngSpc$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnQSqCommaSpc$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnQSqBktSpc$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnSemi$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnOr$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnSpc$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnTab$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnTerm$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnVbar$(A)
'QLib.Std.MVb_JnSplit_Jn.Function JnVbarSpc$(A)
'QLib.Std.MVb_JnSplit_Split.Function SplitComma(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitCommaSpc(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitCrLf(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitTab(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitDot(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitColon(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitSemi(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitSpc(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitSsl(A) As String()
'QLib.Std.MVb_JnSplit_Split.Function SplitVbar(A) As String()
'QLib.Std.MVb_Lg_Lgr.Sub LgrBrw()
'QLib.Std.MVb_Lg_Lgr.Property Get LgrFilNo%()
'QLib.Std.MVb_Lg_Lgr.Property Get LgrFt$()
'QLib.Std.MVb_Lg_Lgr.Sub LgrLg(Msg$)
'QLib.Std.MVb_Lg_Lgr.Property Get LgrPth$()
'QLib.Std.MVb_Lin.Function HasDDRmk(A) As Boolean
'QLib.Std.MVb_Lin.Function IsSngTermLin(A) As Boolean
'QLib.Std.MVb_Lin.Function IsDDLin(A) As Boolean
'QLib.Std.MVb_Lin.Function IsDotLin(A) As Boolean
'QLib.Std.MVb_Lin.Function HasLinT1Ay(Lin, T1Ay$()) As Boolean
'QLib.Std.MVb_Lin.Function PfxLinAp(A, ParamArray PfxAp())
'QLib.Std.MVb_Lin_Lines.Function CntSzStrzLines$(Lines)
'QLib.Std.MVb_Lin_Lines.Function CntSzStr$(Cnt&, Si&)
'QLib.Std.MVb_Lin_Lines.Private Sub Z_LinesWrp()
'QLib.Std.MVb_Lin_Lines.Function LinesWrp$(Lines, Optional Wdt% = 80)
'QLib.Std.MVb_Lin_Lines.Sub Z_LyWrp()
'QLib.Std.MVb_Lin_Lines.Function LyWrp(Ly$(), Optional Wdt% = 80) As String()
'QLib.Std.MVb_Lin_Lines.Private Function LyzLinWrp(Lin, W%) As String()
'QLib.Std.MVb_Lin_Lines.Private Sub ZZ_TrimCrLfAtEnd()
'QLib.Std.MVb_Lin_Lines.Private Sub ZZ_LasNLines()
'QLib.Std.MVb_Lin_Lines.Function FstLin$(Lines)
'QLib.Std.MVb_Lin_Lines.Function LinesRmvBlankLinAtEnd$(Lines)
'QLib.Std.MVb_Lin_Lines.Function LinesApp$(A, L)
'QLib.Std.MVb_Lin_Lines.Function SplitCrLfAy(LinesAy) As String()
'QLib.Std.MVb_Lin_Lines.Sub LinesAsgBrk(A$, Ny0, ParamArray OLyAp())
'QLib.Std.MVb_Lin_Lines.Private Sub Z_TrimCrLfAtEnd()
'QLib.Std.MVb_Lin_Lines.Function LasNLines$(Lines, N%)
'QLib.Std.MVb_Lin_Lines.Function LinCnt&(Lines)
'QLib.Std.MVb_Lin_Lines.Function HSqLines(Lines) As Variant()
'QLib.Std.MVb_Lin_Lines.Function VSqLines(Lines) As Variant()
'QLib.Std.MVb_Lin_Lines.Function TrimR$(S)
'QLib.Std.MVb_Lin_Lines.Function RLenOfCrLf%(S)
'QLib.Std.MVb_Lin_Lines.Function AscAt%(S, Pos)
'QLib.Std.MVb_Lin_Lines.Function IsCrLf(Asc%)
'QLib.Std.MVb_Lin_Lines.Function TrimCrLfAtEnd$(S)
'QLib.Std.MVb_Lin_Lines.Function LasLinLines$(Lines)
'QLib.Std.MVb_Lin_Lines.Function LinesAlignL$(Lines, W%)
'QLib.Std.MVb_Lin_Lines.Function NLines&(Lines$)
'QLib.Std.MVb_Lin_Scl.Sub AsgSclNN(Scl$, NN$, ParamArray OAp())
'QLib.Std.MVb_Lin_Scl.Function ChkSclNN(A$, Ny0) As String()
'QLib.Std.MVb_Lin_Scl.Function SclItm_V(A$, Ny$())
'QLib.Std.MVb_Lin_Scl.Function ShfScl$(OStr$)
'QLib.Std.MVb_Lin_Scl.Private Sub ZZ()
'QLib.Std.MVb_Lin_Scl.Private Sub Z()
'QLib.Std.MVb_Lin_Shf.Function ShfDotSeg$(OLin$)
'QLib.Std.MVb_Lin_Shf.Function ShfBef(OLin$, Sep$)
'QLib.Std.MVb_Lin_Shf.Function ShfBktStr$(OLin$)
'QLib.Std.MVb_Lin_Shf.Function RmvChr$(S, ChrLis$) ' Rmv fst chr if it is in ChrLis
'QLib.Std.MVb_Lin_Shf.Function TakChr$(S, ChrLis$) ' Ret fst chr if it is in ChrLis
'QLib.Std.MVb_Lin_Shf.Function RmvChrzSfx$(S, ChrLis$)
'QLib.Std.MVb_Lin_Shf.Function ShfChr$(OLin, ChrList$)
'QLib.Std.MVb_Lin_Shf.Function ShfPfx(OLin, Pfx) As Boolean
'QLib.Std.MVb_Lin_Shf.Function ShfPfxSpc(OLin, Pfx) As Boolean
'QLib.Std.MVb_Lin_Shf.Private Sub Z_ShfBktStr()
'QLib.Std.MVb_Lin_Shf.Private Property Get Z_ShfPfx()
'QLib.Std.MVb_Lin_Shf.Private Sub Z()
'QLib.Std.MVb_Lin_Term.Function RmvTermAy$(Lin, Ay$())
'QLib.Std.MVb_Lin_Term.Function TLin$(TermAy$())
'QLib.Std.MVb_Lin_Term.Function TLinzAp$(ParamArray TermAp())
'QLib.Std.MVb_Lin_Term.Function JnTermAp$(ParamArray Ap())
'QLib.Std.MVb_Lin_Term.Function JnTermAy$(TermAy$())
'QLib.Std.MVb_Lin_Term.Function TermAyzTT(TT) As String()
'QLib.Std.MVb_Lin_Term.Function LinzTermAy$(TermAy)
'QLib.Std.MVb_Lin_Term.Function TermAset(Lin) As Aset
'QLib.Std.MVb_Lin_Term.Function TermItr(NN)
'QLib.Std.MVb_Lin_Term.Function CvNy(Ny0) As String()
'QLib.Std.MVb_Lin_Term.Function TermAyzNN(NN) As String()
'QLib.Std.MVb_Lin_Term.Function TermAy(Lin) As String()
'QLib.Std.MVb_Lin_Term.Function ShfTermX(OLin, X$) As Boolean
'QLib.Std.MVb_Lin_Term.Function ShfT1$(OLin)
'QLib.Std.MVb_Lin_Term.Function ShfTermDot$(OLin)
'QLib.Std.MVb_Lin_Term.Private Sub Z_ShfT1()
'QLib.Std.MVb_Lin_Term.Private Sub Z()
'QLib.Std.MVb_Lin_Term_Asg.Sub AsgN2tRst(Lin, OT1, OT2, ORst)
'QLib.Std.MVb_Lin_Term_Asg.Sub AsgN3tRst(Lin, OT1, OT2, OT3, ORst)
'QLib.Std.MVb_Lin_Term_Asg.Sub AsgN4t(Lin, O1, O2, O3, O4)
'QLib.Std.MVb_Lin_Term_Asg.Sub AsgN4tRst(Lin, O1, O2, O3, O4, ORst)
'QLib.Std.MVb_Lin_Term_Asg.Sub AsgTRst(Lin, OT1, ORst)
'QLib.Std.MVb_Lin_Term_Asg.Sub AsgN2t(Lin, O1, O2)
'QLib.Std.MVb_Lin_Term_Asg.Sub AsgT1FldLikAy(OT1$, OFldLikAy$(), Lin)
'QLib.Std.MVb_Lin_Term_FstNTerm.Function Fst2Term(Lin) As String()
'QLib.Std.MVb_Lin_Term_FstNTerm.Function Fst3Term(Lin) As String()
'QLib.Std.MVb_Lin_Term_FstNTerm.Function Fst4Term(Lin) As String()
'QLib.Std.MVb_Lin_Term_FstNTerm.Function FstNTerm(Lin, N%) As String()
'QLib.Std.MVb_Lin_Term_NTermRst.Function SyzTRst(Lin) As String()
'QLib.Std.MVb_Lin_Term_NTermRst.Function SyzN2tRst(Lin) As String()
'QLib.Std.MVb_Lin_Term_NTermRst.Function SyzN3TRst(Lin) As String()
'QLib.Std.MVb_Lin_Term_NTermRst.Function SyzN4tRst(Lin) As String()
'QLib.Std.MVb_Lin_Term_NTermRst.Function SyzNTermRst(Lin, N%) As String()
'QLib.Std.MVb_Lin_Term_NTermRst.Private Sub Z_SyzNTermRst()
'QLib.Std.MVb_Lin_Term_NTermRst.Private Sub Z()
'QLib.Std.MVb_Lin_Term_TermN.Function T1zLin$(A)
'QLib.Std.MVb_Lin_Term_TermN.Function T1$(Lin)
'QLib.Std.MVb_Lin_Term_TermN.Function T2zLin$(A)
'QLib.Std.MVb_Lin_Term_TermN.Function T2$(A)
'QLib.Std.MVb_Lin_Term_TermN.Function T3$(A)
'QLib.Std.MVb_Lin_Term_TermN.Function TermN$(Lin, N%)
'QLib.Std.MVb_Lin_Term_TermN.Private Sub Z_TermN()
'QLib.Std.MVb_Lin_Term_TermN.Private Sub Z()
'QLib.Std.MVb_Lin_Vbl.Function DryzTLinAy(TLinAy$()) As Variant()
'QLib.Std.MVb_Lin_Vbl.Function DryzVblLy(A$()) As Variant()
'QLib.Std.MVb_Lin_Vbl.Private Sub Z_DryzVblLy()
'QLib.Std.MVb_Lin_Vbl.Private Sub ZZ_DryzVblLy()
'QLib.Std.MVb_Lin_Vbl.Private Sub Z()
'QLib.Std.MVb_Obj.Function IsEqObj(A, B) As Boolean
'QLib.Std.MVb_Obj.Function IntozOy(OInto, Oy)
'QLib.Std.MVb_Obj.Function LngAyzOyPrp(Oy, Prp) As Long()
'QLib.Std.MVb_Obj.Function IntozOyPrp(OInto, Oy, Prp)
'QLib.Std.MVb_Obj.Function ObjAddAy(Obj, Oy)
'QLib.Std.MVb_Obj.Function ObjNm$(A)
'QLib.Std.MVb_Obj.Function DrzObjPrpNy(Obj, PrpNy$()) As Variant()
'QLib.Std.MVb_Obj.Function LyzObjPrpNy(Obj, B$()) As String()
'QLib.Std.MVb_Obj.Function LyzObjPP(Obj, PP) As String()
'QLib.Std.MVb_Obj.Function DrzObjPP(Obj, PP$) As Variant()
'QLib.Std.MVb_Obj.Function ObjPrp(A, PrpPth, Optional Thw As eThwOpt)
'QLib.Std.MVb_Obj.Function ObjStr$(Obj)
'QLib.Std.MVb_Obj.Private Sub ZZZ_ObjPrp()
'QLib.Std.MVb_Obj.Function Prp(Obj, P)
'QLib.Std.MVb_PfxSfx.Function AddPfx$(A$, Pfx$)
'QLib.Std.MVb_PfxSfx.Function AddPfxSfx$(A$, Pfx$, Sfx$)
'QLib.Std.MVb_PfxSfx.Function AddSfx$(A$, Sfx$)
'QLib.Std.MVb_PfxSfx.Function AddPfxSpc_IfNonBlank$(A)
'QLib.Std.MVb_PfxSfx.Function AyAddPfx(A, Pfx) As String()
'QLib.Std.MVb_PfxSfx.Function AyAddPfxSfx(A, Pfx, Sfx) As String()
'QLib.Std.MVb_PfxSfx.Function AyAddSfx(A, Sfx) As String()
'QLib.Std.MVb_PfxSfx.Function AyIsAllEleHitPfx(A, Pfx$) As Boolean
'QLib.Std.MVb_PfxSfx.Function AyAddCommaSpcSfxExlLas(A) As String()
'QLib.Std.MVb_PfxSfx.Function TakSfxChr$(A, SfxChrLis$, Optional IsCasSen As Boolean)
'QLib.Std.MVb_PfxSfx.Function HasSfxChrLis(A, SfxChrLis$, Optional IsCasSen As Boolean) As Boolean
'QLib.Std.MVb_PfxSfx.Function HasPfx(A, Pfx, Optional IsCasSen As Boolean) As Boolean
'QLib.Std.MVb_PfxSfx.Function HasSfx(A, Sfx, Optional IsCasSen As Boolean) As Boolean
'QLib.Std.MVb_PfxSfx.Function StrEq(A, B, Optional IsCasSen) As Boolean
'QLib.Std.MVb_PfxSfx.Function HasSfxApIgnCas(A, ParamArray SfxAp()) As Boolean
'QLib.Std.MVb_PfxSfx.Function HasSfxAp(A, ParamArray SfxAp()) As Boolean
'QLib.Std.MVb_PfxSfx.Function HasSfxAv(A, SfxAv(), Optional IsCasSen As Boolean) As Boolean
'QLib.Std.MVb_PfxSfx.Function SyIsAllEleHitPfx(A$(), Pfx$) As Boolean
'QLib.Std.MVb_Re.Private Sub ZZ_ReMatch()
'QLib.Std.MVb_Re.Private Sub ZZ_ReRpl()
'QLib.Std.MVb_Rnd.Function AsetNRndStr(N&) As Aset
'QLib.Std.MVb_Rnd.Function AsetNRndLng(N&) As Aset
'QLib.Std.MVb_Rnd.Function AsetNRndInt(N&) As Aset
'QLib.Std.MVb_Rslt.Function LngRslt(Lng) As LngRslt: LngRslt.Som = True: LngRslt.Lng = Lng: End Function
'QLib.Std.MVb_Rslt.Function LyRslt(Er$(), Ly$()) As LyRslt: LyRslt.Er = Er: LyRslt.Ly = Ly: End Function
'QLib.Std.MVb_Rslt.Function StrRslt(Str) As StrRslt: StrRslt.Str = Str: StrRslt.Som = True: End Function
'QLib.Std.MVb_Rslt.Function BoolRslt(Bool As Boolean) As BoolRslt: BoolRslt.Bool = Bool: BoolRslt.Som = True: End Function
'QLib.Std.MVb_Rslt.Function DicRslt(Dic As Dictionary) As DicRslt: Set DicRslt.Dic = Dic: DicRslt.Som = True: End Function
'QLib.Std.MVb_Rslt.Function TrueRslt() As BoolRslt: TrueRslt = BoolRslt(True): End Function
'QLib.Std.MVb_Rslt.Function FalseRslt() As BoolRslt: FalseRslt = BoolRslt(False): End Function
'QLib.Std.MVb_Run.Function Pipe(Pm, MthNN)
'QLib.Std.MVb_Run.Function RunAvzIgnEr(MthNm$, Av())
'QLib.Std.MVb_Run.Function RunAv(MthNm$, Av())
'QLib.Std.MVb_RunFil.Function WaitOpt(TimOutSec%, ChkIntervalDeciSec%, KeepFcmd As Boolean) As WaitOpt
'QLib.Std.MVb_RunFil.Property Get DftWaitOpt() As WaitOpt
'QLib.Std.MVb_RunFil.Sub KillProcessId(ProcessId&)
'QLib.Std.MVb_RunFil.Sub RunFcmd(Fcmd$, ParamArray PmAp())
'QLib.Std.MVb_RunFil.Private Function RunFcmdWaitOpt(Fcmd$, A As WaitOpt, ParamArray PmAp()) As Boolean
'QLib.Std.MVb_RunFil.Private Function RunFcmdWaitOptAv(Fcmd$, A As WaitOpt, ParamArray PmAp()) As Boolean
'QLib.Std.MVb_RunFil.Function RunFcmdWait(Fcmd$, ParamArray PmAp()) As Boolean
'QLib.Std.MVb_RunFil.Function RunFcmdAv&(Fcmd, PmAv())
'QLib.Std.MVb_RunFil.Private Sub ZZ_RunCmd()
'QLib.Std.MVb_RunFil.Function WaitFfnzFcmd$(Ffn)
'QLib.Std.MVb_RunFil.Function Wait(Optional Sec% = 1) As Boolean
'QLib.Std.MVb_RunFil.Function WaitFfn(Ffn, Optional ChkIntervalDeciSec% = 10, Optional TimOutSec% = 60) As Boolean
'QLib.Std.MVb_RunFil.Function WaitDeci(Optional DeciSec% = 10) As Boolean
'QLib.Std.MVb_RunFil.Function NxtDeciSec(DeciSec%) As Date
'QLib.Std.MVb_Run_Cd.Sub RunCdLy(CdLy$())
'QLib.Std.MVb_Run_Cd.Sub RunCd(CdLines$)
'QLib.Std.MVb_Run_Cd.Private Function RunCdMd() As CodeModule
'QLib.Std.MVb_Run_Cd.Private Sub AddMthzCd(MthNm$, CdLines$)
'QLib.Std.MVb_Run_Cd.Private Function MthLines$(MthNm$, CdLines$)
'QLib.Std.MVb_Run_Cd.Private Property Get ZZCdLines$()
'QLib.Std.MVb_Run_Cd.Sub TimFun(FunNN)
'QLib.Std.MVb_Run_Cd.Private Sub ZZ_TimFun()
'QLib.Std.MVb_Run_Cd.Private Sub ZZA()
'QLib.Std.MVb_Run_Cd.Private Sub ZZB()
'QLib.Std.MVb_S1S2.Function SwapS1S2Ay(A() As S1S2) As S1S2()
'QLib.Std.MVb_S1S2.Private Property Get ZZS1S2Ay1() As S1S2()
'QLib.Std.MVb_S1S2.Function S1S2AyAyab(A, B, Optional NoTrim As Boolean) As S1S2()
'QLib.Std.MVb_S1S2.Function CvS1S2(A) As S1S2
'QLib.Std.MVb_S1S2.Function S1S2Ay(ParamArray S1S2Ap()) As S1S2()
'QLib.Std.MVb_S1S2.Function S1S2(S1, S2, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_S1S2.Sub AsgS1S2(A As S1S2, O1, O2)
'QLib.Std.MVb_S1S2.Function S1S2Clone(A As S1S2) As S1S2
'QLib.Std.MVb_S1S2.Function S1S2Lin$(A As S1S2, Optional Sep$ = " ", Optional W1%)
'QLib.Std.MVb_S1S2.Function JnS1S2Ay(A() As S1S2, Optional Sep$ = "") As String()
'QLib.Std.MVb_S1S2.Sub BrwS1S2Ay(A() As S1S2)
'QLib.Std.MVb_S1S2.Function S1S2AyzDic(A As Dictionary) As S1S2()
'QLib.Std.MVb_S1S2.Function DiczS1S2Ay(A() As S1S2, Optional Sep$ = " ") As Dictionary
'QLib.Std.MVb_S1S2.Function Sy1zS1S2Ay(A() As S1S2) As String()
'QLib.Std.MVb_S1S2.Function Sy2zS1S2Ay(A() As S1S2) As String()
'QLib.Std.MVb_S1S2.Function SqzS1S2Ay(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Variant()
'QLib.Std.MVb_S1S2.Function S1S2AyzColonVbl(ColonVbl) As S1S2()
'QLib.Std.MVb_S1S2.Function S1S2AyzAySep(Ay, Sep$, Optional NoTrim As Boolean) As S1S2()
'QLib.Std.MVb_S1S2.Private Sub Z_S1S2AyzDic()
'QLib.Std.MVb_S1S2.Private Sub Z()
'QLib.Std.MVb_S1S2_Fmt.Function FmtS1S2Ay(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As String()
'QLib.Std.MVb_S1S2_Fmt.Function Sy1zS1S2Ay(A() As S1S2) As String()
'QLib.Std.MVb_S1S2_Fmt.Private Function LyzS1S2(A As S1S2, W1%, W2%) As String()
'QLib.Std.MVb_S1S2_Fmt.Function LyAlignLWdt(A, W%) As String()
'QLib.Std.MVb_S1S2_Fmt.Private Function CvS1S2(A) As S1S2
'QLib.Std.MVb_S1S2_Fmt.Private Function LyzS1S2Ay(A() As S1S2, W1%, W2%, S1S2AyHasLines As Boolean, SepLin$) As String()
'QLib.Std.MVb_S1S2_Fmt.Private Function LinzS1S2$(A As S1S2, W1%, W2%)
'QLib.Std.MVb_S1S2_Fmt.Function LinDrWdtAy$(Dr, WdtzAy%())
'QLib.Std.MVb_S1S2_Fmt.Function Sy2zS1S2Ay(A() As S1S2) As String()
'QLib.Std.MVb_S1S2_Fmt.Private Function HasLines(A() As S1S2) As Boolean
'QLib.Std.MVb_S1S2_Fmt.Private Sub Z_FmtS1S2Ay()
'QLib.Std.MVb_S1S2_Fmt.Private Sub Z()
'QLib.Std.MVb_Seed.Function Expand$(Seed$, Ny0)
'QLib.Std.MVb_Seed.Private Sub Z_Expand()
'QLib.Std.MVb_Seed.Private Sub Z()
'QLib.Std.MVb_Stop_Xls.Private Sub AAAAAA()
'QLib.Std.MVb_Stop_Xls.Sub StopXls()
'QLib.Std.MVb_Str.Function AddLib$(V, Lbl$)
'QLib.Std.MVb_Str.Function IsEqStr(A, B, Optional IgnoreCase As Boolean) As Boolean
'QLib.Std.MVb_Str.Function Pad0$(N, NDig%)
'QLib.Std.MVb_Str.Sub BrwStr(A, Optional Fnn$, Optional UseVc As Boolean)
'QLib.Std.MVb_Str.Sub VcStr(A, Optional Fnn$)
'QLib.Std.MVb_Str.Function StrDft$(A, B)
'QLib.Std.MVb_Str.Function Dup$(S, N)
'QLib.Std.MVb_Str.Function HasStrSfxAy(A, SfxAy$()) As Boolean
'QLib.Std.MVb_Str.Function HasStrPfxAy(A, PfxAy$()) As Boolean
'QLib.Std.MVb_Str.Sub EdtStr(S, Ft)
'QLib.Std.MVb_Str.Function WrtStr$(Str, Ft, Optional OvrWrt As Boolean)
'QLib.Std.MVb_Str.Private Sub ZZ()
'QLib.Std.MVb_Str_Appd.Function ApdCrLf$(A)
'QLib.Std.MVb_Str_Appd.Function PpdSpc$(A)
'QLib.Std.MVb_Str_Appd.Function Apd$(A, Sfx, Optional Sep$ = "")
'QLib.Std.MVb_Str_Appd.Function Ppd$(A, Pfx, Optional Sep$ = "")
'QLib.Std.MVb_Str_Bkt.Private Sub Z_AsgBktPos()
'QLib.Std.MVb_Str_Bkt.Private Sub Z_Brk_Bkt()
'QLib.Std.MVb_Str_Bkt.Sub AsgBktPos(A, OpnBkt$, OFmPos%, OToPos%)
'QLib.Std.MVb_Str_Bkt.Function ClsBkt$(OpnBkt$)
'QLib.Std.MVb_Str_Bkt.Function BrkBkt(A, Optional OpnBkt$ = vbOpnBkt) As String()
'QLib.Std.MVb_Str_Bkt.Function BetBktMust$(S, Fun$, Optional OpnBkt$ = vbOpnBkt)
'QLib.Std.MVb_Str_Bkt.Function BetBkt$(A, Optional OpnBkt$ = vbOpnBkt)
'QLib.Std.MVb_Str_Bkt.Function AftBkt$(Lin, Optional OpnBkt$ = vbOpnBkt)
'QLib.Std.MVb_Str_Bkt.Function BefBkt$(Lin, Optional OpnBkt$ = vbOpnBkt)
'QLib.Std.MVb_Str_Bkt.Private Sub Z()
'QLib.Std.MVb_Str_Box.Function BoxLyLines(Lines$) As String()
'QLib.Std.MVb_Str_Box.Function BoxLyAy(Ay) As String()
'QLib.Std.MVb_Str_Brk.Sub AsgBrk1Dot(S, OA, OB, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Brk.Sub AsgBrkDot(S, OA, OB, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Brk.Function Brk1Dot(S, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Function Brk2Dot(S, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Function BrkDot(S, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Function Brk(A, Sep, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Function Brk1(A, Sep, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Sub AsgBrk1(A, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Brk.Function Brk1Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Function Brk2(A, Sep, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Function Brk2__(A, P&, Sep, NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Sub AsgBrk2(A, Sep$, O1, O2, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Brk.Function Brk2Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Sub AsgBrk(A, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Brk.Private Function BrkAtSep(A, P&, Sep, NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Function Brk1At(A, P&, Sep, NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Sub AsgBrkAt(A, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Brk.Function BrkBoth(A, Sep, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Sub AsgBrkQuote(QuoteStr, O1$, O2$)
'QLib.Std.MVb_Str_Brk.Function BrkRev(A, Sep, Optional NoTrim As Boolean) As S1S2
'QLib.Std.MVb_Str_Brk.Sub AsgBrk1At(A, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Brk.Private Sub ZZ_Brk1Rev()
'QLib.Std.MVb_Str_Brk.Private Sub Z_Brk1Rev()
'QLib.Std.MVb_Str_Brk.Private Sub Z()
'QLib.Std.MVb_Str_Cmp.Sub CmpLines(A, B, Optional N1$ = "A", Optional N2$ = "B")
'QLib.Std.MVb_Str_Cmp.Function CmpLinesFmt(A, B, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$) As String()
'QLib.Std.MVb_Str_Cmp.Private Function LyAll(A$(), Nm$) As String()
'QLib.Std.MVb_Str_Cmp.Private Function LyzCmpStr(A$, B$, Ix&) As String()
'QLib.Std.MVb_Str_Cmp.Private Function LyRest(A$(), B$(), MinU&, Nm1$, Nm2$) As String()
'QLib.Std.MVb_Str_Cmp.Sub CmpStr(A, B, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$)
'QLib.Std.MVb_Str_Cmp.Function CmpStrFmt(A, B, Optional N1$ = "A", Optional N2$ = "B", Optional Hdr$) As String()
'QLib.Std.MVb_Str_Cmp.Private Function WDifAt&(A, B)
'QLib.Std.MVb_Str_Cmp.Private Function DifAtIx&(A$(), B$())
'QLib.Std.MVb_Str_Cmp.Function Len_LblAy(L&) As String()
'QLib.Std.MVb_Str_Cmp.Private Function Len_LblLin1$(L&)
'QLib.Std.MVb_Str_Cmp.Private Function Len_LblLin2$(L&)
'QLib.Std.MVb_Str_Cmp.Private Sub Z_FmtCmpLines()
'QLib.Std.MVb_Str_Cmp.Private Sub ZZ()
'QLib.Std.MVb_Str_Cmp.Private Sub Z()
'QLib.Std.MVb_Str_Ens.Function SfxEns$(S, Sfx$)
'QLib.Std.MVb_Str_Ens.Function SfxDotEns$(S)
'QLib.Std.MVb_Str_Ens.Function SfxSemiEns$(S)
'QLib.Std.MVb_Str_Esc.Function Esc$(A, Fm$, ToStr$)
'QLib.Std.MVb_Str_Esc.Function EscBackSlash$(A)
'QLib.Std.MVb_Str_Esc.Function EscCr$(A)
'QLib.Std.MVb_Str_Esc.Function EscCrLf$(A)
'QLib.Std.MVb_Str_Esc.Function EscKey$(A)
'QLib.Std.MVb_Str_Esc.Function EscLf$(A)
'QLib.Std.MVb_Str_Esc.Function EscSpc$(A)
'QLib.Std.MVb_Str_Esc.Function EscSqBkt$(A)
'QLib.Std.MVb_Str_Esc.Function EscTab$(A)
'QLib.Std.MVb_Str_Esc.Function EscUnCr$(A)
'QLib.Std.MVb_Str_Esc.Function EscUnSpc$(A)
'QLib.Std.MVb_Str_Esc.Function EscUnTab(A)
'QLib.Std.MVb_Str_Esc.Function UnEscBackSlash$(A)
'QLib.Std.MVb_Str_Esc.Function UnEscCr$(A)
'QLib.Std.MVb_Str_Esc.Function UnEscCrLf$(A)
'QLib.Std.MVb_Str_Esc.Function UnEscLf$(A)
'QLib.Std.MVb_Str_Esc.Function UnEscSpc$(A)
'QLib.Std.MVb_Str_Esc.Function UnEscSqBkt$(A)
'QLib.Std.MVb_Str_Esc.Function UnEscTab(A)
'QLib.Std.MVb_Str_Expand.Function Expand$(QVbl$, ExpandByTLin$)
'QLib.Std.MVb_Str_Expand.Private Sub Z_Expand()
'QLib.Std.MVb_Str_Filter.Function Filter(S, Src$()) As String()
'QLib.Std.MVb_Str_Filter.Private Function Hit(S, Src) As Boolean
'QLib.Std.MVb_Str_Fmt.Function FmtQQ$(QQVbl$, ParamArray Ap())
'QLib.Std.MVb_Str_Fmt.Function FmtQQAv$(QQVbl, Av())
'QLib.Std.MVb_Str_Fmt.Function SpcSepStr$(A)
'QLib.Std.MVb_Str_Fmt.Function SpcSepStrRev$(A)
'QLib.Std.MVb_Str_Fmt.Private Sub ZZ_FmtQQAv()
'QLib.Std.MVb_Str_Fmt.Function LblTabFmtAySepSS(Lbl$, Ay) As String()
'QLib.Std.MVb_Str_Has.Function HasDot(Str) As Boolean
'QLib.Std.MVb_Str_Has.Function HasSubStr(Str, SubStr, Optional IgnCas As Boolean) As Boolean
'QLib.Std.MVb_Str_Has.Function HasCrLf(A) As Boolean
'QLib.Std.MVb_Str_Has.Function HasHyphen(A) As Boolean
'QLib.Std.MVb_Str_Has.Function HasPound(A) As Boolean
'QLib.Std.MVb_Str_Has.Function HasSpc(A) As Boolean
'QLib.Std.MVb_Str_Has.Function HasSqBkt(A) As Boolean
'QLib.Std.MVb_Str_Has.Function HasChrList(A, ChrList$) As Boolean
'QLib.Std.MVb_Str_Has.Function HasSubStrAy(A, SubStrAy$()) As Boolean
'QLib.Std.MVb_Str_Has.Function HasTT(L, T1, T2) As Boolean
'QLib.Std.MVb_Str_Has.Function HasT1(L, T) As Boolean
'QLib.Std.MVb_Str_Has.Function HasVbar(A$) As Boolean
'QLib.Std.MVb_Str_Likss.Function StrLikss(A, Likss) As Boolean
'QLib.Std.MVb_Str_Likss.Function StrLikAy(A, LikeAy$()) As Boolean
'QLib.Std.MVb_Str_Likss.Function StrLikssAy(A, LikssAy) As Boolean
'QLib.Std.MVb_Str_Likss.Private Sub Z_T1zT1LikTLinAy()
'QLib.Std.MVb_Str_Likss.Function T1zT1LikTLinAy$(T1LikTLinAy$(), Nm)
'QLib.Std.MVb_Str_Likss.Private Sub Z()
'QLib.Std.MVb_Str_Macro.Function NyzMacro(A, Optional ExlBkt As Boolean, Optional OpnBkt$ = vbOpnBigBkt) As String()
'QLib.Std.MVb_Str_Macro.Function FmtMacro(MacroStr$, ParamArray Ap()) As String()
'QLib.Std.MVb_Str_Macro.Function FmtMacroAv(MacroStr$, Av()) As String()
'QLib.Std.MVb_Str_Macro.Function FmtMacroDic$(MacroStr$, Dic As Dictionary)
'QLib.Std.MVb_Str_Nm.Function IsNm(S) As Boolean
'QLib.Std.MVb_Str_Nm.Function IsNmChr(A$) As Boolean
'QLib.Std.MVb_Str_Nm.Function WhNmzStr(WhStr$, Optional NmPfx$) As WhNm
'QLib.Std.MVb_Str_Nm.Function ChrQuote$(S, Chr$)
'QLib.Std.MVb_Str_Nm.Function SpcQuote$(S)
'QLib.Std.MVb_Str_Nm.Function DblQuote$(S)
'QLib.Std.MVb_Str_Nm.Function SngQuote$(S)
'QLib.Std.MVb_Str_Nm.Function HitRe(Str, Re As RegExp) As Boolean
'QLib.Std.MVb_Str_Nm.Function NmSfx$(A)
'QLib.Std.MVb_Str_Nm.Function NxtSeqNm$(Nm, Optional NDig% = 3) 'Nm-Nm can be XXX or XXX_nn
'QLib.Std.MVb_Str_Quote.Function BrkQuote(QuoteStr) As S1S2
'QLib.Std.MVb_Str_Quote.Function QuoteBkt$(A)
'QLib.Std.MVb_Str_Quote.Function QuoteDot$(S)
'QLib.Std.MVb_Str_Quote.Function Quote$(A, QuoteStr$)
'QLib.Std.MVb_Str_Quote.Function QuoteDblVb$(A)
'QLib.Std.MVb_Str_Quote.Function QuoteDbl$(A)
'QLib.Std.MVb_Str_Quote.Function QuoteSng$(A)
'QLib.Std.MVb_Str_Quote.Function QuoteSq$(A)
'QLib.Std.MVb_Str_Quote.Function QuoteSqIf$(S)
'QLib.Std.MVb_Str_Quote.Function QuoteSqAv(Av()) As String()
'QLib.Std.MVb_Str_Rmv.Function RmvDotComma$(A)
'QLib.Std.MVb_Str_Rmv.Function Rmv2Dash$(A)
'QLib.Std.MVb_Str_Rmv.Function Rmv3Dash$(A)
'QLib.Std.MVb_Str_Rmv.Function Rmv3T$(A$)
'QLib.Std.MVb_Str_Rmv.Function RmvAft$(A, Sep$)
'QLib.Std.MVb_Str_Rmv.Function RmvDDRmk$(A$)
'QLib.Std.MVb_Str_Rmv.Function RmzlSpc$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvFstChr$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvFstLasChr$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvFstNChr$(A, Optional N% = 1)
'QLib.Std.MVb_Str_Rmv.Function RmvFstNonLetter$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvLasChr$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvLasNChr$(A, N%)
'QLib.Std.MVb_Str_Rmv.Function RmvNm$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvSqBkt$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvPfx$(A, Pfx)
'QLib.Std.MVb_Str_Rmv.Function RmvPfxAy$(A, PfxAy)
'QLib.Std.MVb_Str_Rmv.Function RmvPfxSpc$(A, Pfx)
'QLib.Std.MVb_Str_Rmv.Function RmvPfxAySpc$(A, PfxAy)
'QLib.Std.MVb_Str_Rmv.Function RmvBkt$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvSfxzBkt$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvSfx$(A, Sfx)
'QLib.Std.MVb_Str_Rmv.Function RmvSngQuote$(A)
'QLib.Std.MVb_Str_Rmv.Function RmvT1$(S)
'QLib.Std.MVb_Str_Rmv.Function RmvTT$(A)
'QLib.Std.MVb_Str_Rmv.Private Sub Z_RmvT1()
'QLib.Std.MVb_Str_Rmv.Private Sub Z_RmvNm()
'QLib.Std.MVb_Str_Rmv.Private Sub Z_RmvPfx()
'QLib.Std.MVb_Str_Rmv.Private Sub Z_RmvPfxAy()
'QLib.Std.MVb_Str_Rmv.Private Sub Z()
'QLib.Std.MVb_Str_Rpl.Private Sub ZZ_RplBet()
'QLib.Std.MVb_Str_Rpl.Private Sub ZZ_RplPfx()
'QLib.Std.MVb_Str_Rpl.Function RmvCr$(A)
'QLib.Std.MVb_Str_Rpl.Function RplCr$(A)
'QLib.Std.MVb_Str_Rpl.Function RplLf$(A)
'QLib.Std.MVb_Str_Rpl.Function RplVbl$(Vbl)
'QLib.Std.MVb_Str_Rpl.Function RplVbar$(Vbl)
'QLib.Std.MVb_Str_Rpl.Function RplBet$(A, By$, S1$, S2$)
'QLib.Std.MVb_Str_Rpl.Function RplDblSpc$(A)
'QLib.Std.MVb_Str_Rpl.Function RplFstChr$(A, By$)
'QLib.Std.MVb_Str_Rpl.Function RplPfx(A, FmPfx, ToPfx)
'QLib.Std.MVb_Str_Rpl.Private Sub Z_RplPfx()
'QLib.Std.MVb_Str_Rpl.Function RplPun$(A)
'QLib.Std.MVb_Str_Rpl.Function RplQ$(A, By)
'QLib.Std.MVb_Str_Rpl.Private Sub Z_RplBet()
'QLib.Std.MVb_Str_Rpl.Private Sub Z()
'QLib.Std.MVb_Str_SubStr.Function LasChr$(A)
'QLib.Std.MVb_Str_SubStr.Function SndChr$(A)
'QLib.Std.MVb_Str_SubStr.Function FstAsc%(A)
'QLib.Std.MVb_Str_SubStr.Function SndAsc%(A)
'QLib.Std.MVb_Str_SubStr.Function FstChr$(A)
'QLib.Std.MVb_Str_SubStr.Function FstTwoChr$(A)
'QLib.Std.MVb_Str_SubStr.Function SubStrCnt&(S, SubStr)
'QLib.Std.MVb_Str_SubStr.Function PoszSubStr(A, SubStr$) As Pos
'QLib.Std.MVb_Str_SubStr.Private Sub Z_SubStrCnt()
'QLib.Std.MVb_Str_SubStr.Function DotCnt&(S)
'QLib.Std.MVb_Str_Tak.Function StrBefDot$(A)
'QLib.Std.MVb_Str_Tak.Function StrAft$(S, Sep, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Tak.Function StrAftAt$(A, At&, S)
'QLib.Std.MVb_Str_Tak.Function StrAftDotOrAll$(A)
'QLib.Std.MVb_Str_Tak.Function StrAftDot$(A)
'QLib.Std.MVb_Str_Tak.Function StrAftMust$(A, Sep, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Tak.Function StrAftOrAll$(S, Sep, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Tak.Function StrAftOrAllRev$(A, S)
'QLib.Std.MVb_Str_Tak.Function StrAftRev$(S, Sep, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Tak.Function StrBef$(S, Sep, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Tak.Function StrBefAt(A, At&)
'QLib.Std.MVb_Str_Tak.Function StrBefDD$(A)
'QLib.Std.MVb_Str_Tak.Function StrBefDDD$(A)
'QLib.Std.MVb_Str_Tak.Function StrBefMust$(S, Sep$, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Tak.Function StrBefOrAll$(S, Sep, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Tak.Function StrBefOrAllRev$(A, S)
'QLib.Std.MVb_Str_Tak.Function StrBefRev$(A, Sep, Optional NoTrim As Boolean)
'QLib.Std.MVb_Str_Tak.Function TakP123(A, S1, S2) As String()
'QLib.Std.MVb_Str_Tak.Sub TakP123Asg(A, S1, S2, O1, O2, O3)
'QLib.Std.MVb_Str_Tak.Private Sub Z_Tak_BefFstLas()
'QLib.Std.MVb_Str_Tak.Function StrBetFstLas$(S, Fst, Las)
'QLib.Std.MVb_Str_Tak.Function StrBet$(S, S1, S2, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
'QLib.Std.MVb_Str_Tak.Private Sub Z_Tak_BetBkt()
'QLib.Std.MVb_Str_Tak.Function TakNm$(A)
'QLib.Std.MVb_Str_Tak.Function TakPfx$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx else return ""
'QLib.Std.MVb_Str_Tak.Function PfxAyFstSpc$(PfxAy$(), Lin) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
'QLib.Std.MVb_Str_Tak.Function PfxLinAy$(A, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
'QLib.Std.MVb_Str_Tak.Function SfxLinAy$(A, SfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
'QLib.Std.MVb_Str_Tak.Function TermLinAy$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
'QLib.Std.MVb_Str_Tak.Function TakPfxS$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx+" " else return ""
'QLib.Std.MVb_Str_Tak.Function TakT1$(A)
'QLib.Std.MVb_Str_Tak.Private Sub Z_AftBkt()
'QLib.Std.MVb_Str_Tak.Private Sub Z_StrBet()
'QLib.Std.MVb_Str_Tak.Private Sub ZZ_Tak_BetBkt()
'QLib.Std.MVb_Str_Tak.Private Sub Z()
'QLib.Std.MVb_Str_Tak.Function StrBefRevOrAll$(S, Sep$)
'QLib.Std.MVb_Str_Trim.Function TrimWhite$(A)
'QLib.Std.MVb_Str_Trim.Function TrimWhiteL$(A)
'QLib.Std.MVb_Str_Trim.Function TrimWhiteR$(S)
'QLib.Std.MVb_Str_UnderLin.Function UnderLin$(A$)
'QLib.Std.MVb_Str_UnderLin.Function UnderLinDbl$(A)
'QLib.Std.MVb_Str_UnderLin.Function PushMsgUnderLinDbl(O$(), M$)
'QLib.Std.MVb_Str_UnderLin.Function PushUnderLin(O$())
'QLib.Std.MVb_Str_UnderLin.Function PushUnderLinDbl(O$())
'QLib.Std.MVb_Str_UnderLin.Function LinesUnderLin$(A)
'QLib.Std.MVb_Str_UnderLin.Function PushMsgUnderLin(O$(), M$)
'QLib.Std.MVb_SyPair.Function SyPair(A, B) As SyPair
'QLib.Std.MVb_TermDefinition.Function DefzCml() As String()
'QLib.Std.MVb_Thw.Sub ThwIfNEgEle(Ay, Fun$)
'QLib.Std.MVb_Thw.Sub ThwIfNESz(A, B, Fun$)
'QLib.Std.MVb_Thw.Sub ThwIfNE(A, B, Optional ANm$ = "A", Optional BNm$ = "B")
'QLib.Std.MVb_Thw.Private Sub ThwIfNEAy(AyA, AyB, ANm$, BNm$)
'QLib.Std.MVb_Thw.Sub ThwDifTy(A, B, Optional ANm$ = "A", Optional BNm$ = "B")
'QLib.Std.MVb_Thw.Sub ThwDifSz(A, B, Optional ANm$ = "A", Optional BNm$ = "B")
'QLib.Std.MVb_Thw.Sub ThwNotExistFfn(Ffn$, Fun$, Optional FilKd$ = "file")
'QLib.Std.MVb_Thw.Sub ThwEqObjNav(A, B, Fun$, Msg$, Nav())
'QLib.Std.MVb_Thw.Sub ThwAyNotSrt(Ay, Fun$)
'QLib.Std.MVb_Thw.Sub ThwOpt(Thw As eThwOpt, Fun$, Msg$, ParamArray Nap())
'QLib.Std.MVb_Thw.Sub Thw(Fun$, Msg$, ParamArray Nap())
'QLib.Std.MVb_Thw.Sub ThwNav(Fun$, Msg$, Nav())
'QLib.Std.MVb_Thw.Sub Ass(A As Boolean)
'QLib.Std.MVb_Thw.Sub ThwNothing(A, Fun$)
'QLib.Std.MVb_Thw.Sub ThwNotAy(A, Fun$)
'QLib.Std.MVb_Thw.Sub ThwIfNEver(Fun$, Optional Msg$ = "Program should not reach here")
'QLib.Std.MVb_Thw.Sub Halt(Optional Fun$)
'QLib.Std.MVb_Thw.Sub Done()
'QLib.Std.MVb_Thw.Sub ThwPgmEr(Er$(), Fun$)
'QLib.Std.MVb_Thw.Function NavAddNNAv(Nav(), NN$, Av()) As Variant()
'QLib.Std.MVb_Thw.Function NavAddNmV(Nav(), Nm$, V) As Variant()
'QLib.Std.MVb_Thw.Sub ThwErMsg(Er$(), Fun$, Msg$, ParamArray Nap())
'QLib.Std.MVb_Thw.Sub ThwEr(Er$(), Fun$)
'QLib.Std.MVb_Thw.Sub ThwLoopingTooMuch(Fun$)
'QLib.Std.MVb_Thw.Sub ThwPmEr(PmVal, Fun$, Optional MsgWhyPmEr$ = "Invalid value")
'QLib.Std.MVb_Thw.Sub D(Optional A)
'QLib.Std.MVb_Thw.Sub Dmp(A)
'QLib.Std.MVb_Thw.Sub DmpTy(A)
'QLib.Std.MVb_Thw.Sub DmpAyWithIx(Ay)
'QLib.Std.MVb_Thw.Sub DmpAy(Ay)
'QLib.Std.MVb_Thw.Sub InfLin(Fun$, Msg$, ParamArray Nap())
'QLib.Std.MVb_Thw.Sub InfNav(Fun$, Msg$, Nav())
'QLib.Std.MVb_Thw.Sub Inf(Fun$, Msg$, ParamArray Nap())
'QLib.Std.MVb_Thw.Sub WarnLin(Fun$, Msg$, ParamArray Nap())
'QLib.Std.MVb_Thw.Sub Warn(Fun$, Msg$, ParamArray Nap())
'QLib.Std.MVb_Thw.Private Sub Z_InfObjPP()
'QLib.Std.MVb_Thw.Private Sub ZZ()
'QLib.Std.MVb_Thw.Sub StopEr(Er$())
'QLib.Std.MVb_Thw.Sub ThwEqObj(A, B, Fun$, Optional Msg$ = "Two given object cannot be same")
'QLib.Std.MVb_Thw_Msg.Function VblzLines$(Lines)
'QLib.Std.MVb_Thw_Msg.Function LinzFunMsg$(Fun$, Msg$)
'QLib.Std.MVb_Thw_Msg.Function LyzFunMsgNav(Fun$, Msg$, Nav()) As String()
'QLib.Std.MVb_Thw_Msg.Function LyzFunMsgNap(Fun$, Msg$, ParamArray Nap()) As String()
'QLib.Std.MVb_Thw_Msg.Function LyzFunMsgObjPP(Fun$, Msg$, Obj, PP) As String()
'QLib.Std.MVb_Thw_Msg.Function LyzFunMsgNyAv(Fun$, Msg$, Ny$(), Av()) As String()
'QLib.Std.MVb_Thw_Msg.Sub InfObjPP(Fun$, Msg$, Obj, PP)
'QLib.Std.MVb_Thw_Msg.Function LyzNv(Nm$, V, Optional Sep$ = ": ") As String()
'QLib.Std.MVb_Thw_Msg.Function LyzNvzStr$(Nm$, V)
'QLib.Std.MVb_Thw_Msg.Function LyzMsgNap(Msg$, ParamArray Nap()) As String()
'QLib.Std.MVb_Thw_Msg.Function LyzMsg(Msg$) As String()
'QLib.Std.MVb_Thw_Msg.Function LyzMsgNav(Msg$, Nav()) As String()
'QLib.Std.MVb_Thw_Msg.Function LyzNNAp(NN$, ParamArray Ap()) As String()
'QLib.Std.MVb_Thw_Msg.Function LyzNNAv(NN$, Av()) As String()
'QLib.Std.MVb_Thw_Msg.Function LinzFunMsgNav$(Fun$, Msg$, Nav())
'QLib.Std.MVb_Thw_Msg.Sub DmpNNAp(NN, ParamArray Ap())
'QLib.Std.MVb_Thw_Msg.Function LyzNyAv(Ny$(), Av(), Optional Sep$ = ": ") As String()
'QLib.Std.MVb_Thw_Msg.Function LinzLyzMsgNav$(Msg$, Nav())
'QLib.Std.MVb_Thw_Msg.Function LinzNav$(Nav())
'QLib.Std.MVb_Thw_Msg.Function LinzNyAv$(Ny$(), Av())
'QLib.Std.MVb_Thw_Msg.Sub AsgNyAv(Nav(), ONy$(), OAv())
'QLib.Std.MVb_Thw_Msg.Function LyzNav(Nav()) As String()
'QLib.Std.MVb_Thw_Msg.Function SclzNyAv$(Ny$(), Av())
'QLib.Std.MVb_Thw_Msg.Function Box(S) As String()
'QLib.Std.MVb_Thw_Msg.Private Function LyzFunMsg(Fun$, Msg$) As String()
'QLib.Std.MVb_Tim.Sub TimBeg(Optional Msg$ = "Time")
'QLib.Std.MVb_Tim.Sub TimEnd(Optional Halt As Boolean)
'QLib.Std.MVb_Tim.Sub Stamp(S$)
'QLib.Std.MVb_Tim_DteTimStr.Function IsDteTimStr(Str) As Boolean
'QLib.Std.MVb_Tim_DteTimStr.Function IsHHMMSS(HHMMSS$) As Boolean
'QLib.Std.MVb_Tim_DteTimStr.Function IsYYYYDashMMDashMM(A$) As Boolean
'QLib.Std.MVb_Tim_DteTimStr.Function TimStr$(A As Date)
'QLib.Std.MVb_Tim_DteTimStr.Function DteTimStr$(A As Date)
'QLib.Std.MVb_Tim_DteTimStr.Property Get NowDteTimStr$()
'QLib.Std.MVb_Tim_DteTimStr.Property Get NowStr$()
'QLib.Std.MVb_Tst.Sub StopNE()
'QLib.Std.MVb_Tst.Sub C(Optional A, Optional E)
'QLib.Std.MVb_Tst.Sub BrwTstPth(Fun$, Cas$)
'QLib.Std.MVb_Tst.Private Function TstPth$(Fun$, Cas$)
'QLib.Std.MVb_Tst.Property Get TstHom$()
'QLib.Std.MVb_Tst.Sub BrwTstHom()
'QLib.Std.MVb_Tst.Sub ShwTstOk(Fun$, Cas$)
'QLib.Std.MVb_Tst.Function TstTxt$(Fun$, Cas$, Itm$, Optional IsEdt As Boolean)
'QLib.Std.MVb_Tst.Sub EdtTstTxt(Fun$, Cas$, Itm$)
'QLib.Std.MVb_Tst.Private Function TstFt$(Fun$, Cas$, Itm$)
'QLib.Std.MVb_Tst_Dic.Private Sub Can_A_AyDic_To_Be_Pushed()
'QLib.Std.MVb_Tst_TstHomClr.Sub ClsTstHom() ' Rmv-Empty-Pth Rmk-Pth-As-At
'QLib.Std.MVb_Tst_TstHomClr.Private Sub Ren_PthPj_AsAt()
'QLib.Std.MVb_Tst_TstHomClr.Private Sub Ren_MdPth_AsAt()
'QLib.Std.MVb_Tst_TstHomClr.Private Sub Ren_MthPth_AsAt()
'QLib.Std.MVb_Tst_TstHomClr.Private Sub Ren_CasPth_AsAt()
'QLib.Std.MVb_Tst_TstHomClr.Private Property Get CasPthAy() As String()
'QLib.Std.MVb_Tst_TstHomClr.Private Sub Ren(PthAy)
'QLib.Std.MVb_UI.Function Cfm(Msg$) As Boolean
'QLib.Std.MVb_UI.Function CfmYes(Msg$) As Boolean
'QLib.Std.MVb_UI.Sub PromptCnl(Optional Msg = "Should cancel and check")
'QLib.Std.MVb_Val.Function LineszVal$(V)
'QLib.Std.MVb_Val.Function StrCellzVal$(V, Optional ShwZer As Boolean, Optional MaxWdt%)
'QLib.Std.MVb_Val.Function LyzVal(V) As String()
'QLib.Std.MVb_Wrd_Cnt.Private Sub Z_WrdCntDic()
'QLib.Std.MVb_Wrd_Cnt.Function WrdCntDic(S) As Dictionary
'QLib.Std.MVb_Wrd_Cnt.Function WrdAset(S) As Aset
'QLib.Std.MVb_Wrd_Cnt.Function CvMch(A) As IMatch
'QLib.Std.MVb_Wrd_Cnt.Function FstWrdAsetOfPjSrc() As Aset
'QLib.Std.MVb_Wrd_Cnt.Function FstWrd$(S)
'QLib.Std.MVb_Wrd_Cnt.Function WrdMch(S) As MatchCollection
'QLib.Std.MVb_Wrd_Cnt.Function WrdAy(S) As String()
'QLib.Std.MVb_Wrd_Pos.Function WrdLblLinPos$(WrdPos%(), OFmNo&)
'QLib.Std.MVb_Wrd_Pos.Function WrdLblLin$(Lin, OFmNo&)
'QLib.Std.MVb_Wrd_Pos.Function WrdPosAy(Lin) As Integer()
'QLib.Std.MVb_Wrd_Pos.Function WrdLblLinPairLno(Lin, Lno&, LnoWdt, OFmNo&) As String()
'QLib.Std.MVb_Wrd_Pos.Function WrdLblLinPair(Lin, OFmNo&) As String()
'QLib.Std.MVb_Wrd_Pos.Function WrdLblLy(Ly$(), OFmNo&) As String()
'QLib.Std.MVb_Wrd_Pos.Private Sub Z_WrdLblLin()
'QLib.Std.MVb_Wrd_Pos.Private Sub Z_WrdPosAy()
'QLib.Std.MVb_Wrd_Pos.Private Sub Z_WrdLblLy()
'QLib.Std.MVb_X.Sub X(S$)
'QLib.Std.MVb_X.Sub X0(S$)
'QLib.Std.MVb_X.Sub X1(S$)
'QLib.Std.MXls_AddIn.Function AddinsDrs(A As Excel.Application) As Drs
'QLib.Std.MXls_AddIn.Sub DmpAddinsXls()
'QLib.Std.MXls_AddIn.Sub DmpAddins(A As Excel.Application)
'QLib.Std.MXls_AddIn.Function AddinsWs(A As Excel.Application) As Worksheet
'QLib.Std.MXls_AddIn.Function Addin(A As Excel.Application, FxaNm) As Excel.Addin
'QLib.Std.MXls_Cell_Clr.Sub ClrCellBelow(Cell As Range)
'QLib.Std.MXls_Cell_Clr.Function RgzBelowCell(Cell As Range)
'QLib.Std.MXls_Colr.Property Get ColrLy() As String()
'QLib.Std.MXls_Colr.Property Get ColrSq() As Variant()
'QLib.Std.MXls_Colr.Function ColrStr$(A&)
'QLib.Std.MXls_Colr.Function Colr&(ColrNm)
'QLib.Std.MXls_Colr.Property Get ColrWb() As Workbook
'QLib.Std.MXls_Colr.Private Sub SetColr_ToDo()
'QLib.Std.MXls_Colr.Sub FSharpBldKnownColor()
'QLib.Std.MXls_Dao.Function CvCn(A) As ADODB.Connection
'QLib.Std.MXls_Dao.Sub RplOleWcFb(Wc As WorkbookConnection, Fb)
'QLib.Std.MXls_Dao.Sub RplLozFbzFbt(Lo As ListObject, Fb As Database, T)
'QLib.Std.MXls_Dao.Function WbzFb(Fb, Optional Vis As Boolean) As Workbook
'QLib.Std.MXls_Dao.Function WbzTT(Db As Database, TT, Optional UseWc As Boolean) As Workbook
'QLib.Std.MXls_Dao.Function SetWsn(Ws As Worksheet, Nm$) As Worksheet
'QLib.Std.MXls_Dao.Sub AddWszT(Wb As Workbook, A As Database, T, Optional Wsn0$)
'QLib.Std.MXls_Dao.Function WbzOupTbl(Db As Database) As Workbook
'QLib.Std.MXls_Dao.Function WbzT(Db As Database, T, Optional Wsn$ = "Data", Optional LoNm$, Optional Vis As Boolean) As Workbook
'QLib.Std.MXls_Dao.Function AtAddDbt(At As Range, Db As Database, T, Optional LoNm$) As Range
'QLib.Std.MXls_Dao.Sub PutDbtWs(A As Database, T, Ws As Worksheet)
'QLib.Std.MXls_Dao.Sub PutDbtAt(A As Database, T, At As Range, Optional LoNm$)
'QLib.Std.MXls_Dao.Sub SetQtFbt(Qt As QueryTable, Fb$, T)
'QLib.Std.MXls_Dao.Sub PutFbtAt(Fb$, T, At As Range, Optional LoNm0$)
'QLib.Std.MXls_Dao.Sub FxzTT(Fx$, Db As Database, TT)
'QLib.Std.MXls_Dao.Function WszWbT(Wb As Workbook, Db As Database, T, Optional Wsn0$) As Worksheet
'QLib.Std.MXls_Dao.Function WszT(Db As Database, T, Optional Wsn$) As Worksheet
'QLib.Std.MXls_Fm_Dta.Function RgzDrs(A As Drs, At As Range) As Range
'QLib.Std.MXls_Fm_Dta.Function LozDrs(Drs As Drs, At As Range, Optional LoNm$) As ListObject
'QLib.Std.MXls_Fm_Dta.Function WSumSi(Ay, Optional Wsn$ = "Sheet1") As Worksheet
'QLib.Std.MXls_Fm_Dta.Function WszDrs(Drs As Drs, Optional Wsn$ = "Sheet1", Optional Vis As Boolean) As Worksheet
'QLib.Std.MXls_Fm_Dta.Function RgzAyV(Ay, At As Range) As Range
'QLib.Std.MXls_Fm_Dta.Function RgzAyH(Ay, At As Range) As Range
'QLib.Std.MXls_Fm_Dta.Function RgzDry(Dry(), At As Range) As Range
'QLib.Std.MXls_Fm_Dta.Function WszDry(Dry(), Optional Wsn$ = "Sheet1") As Worksheet
'QLib.Std.MXls_Fm_Dta.Function WbzDs(A As Ds) As Workbook
'QLib.Std.MXls_Fm_Dta.Function WszDs(A As Ds) As Worksheet
'QLib.Std.MXls_Fm_Dta.Function RgzDt(A As Dt, At As Range, Optional DtIx%)
'QLib.Std.MXls_Fm_Dta.Function LozDt(A As Dt, At As Range) As ListObject
'QLib.Std.MXls_Fm_Dta.Function WszWbDt(Wb As Workbook, Dt As Dt) As Worksheet
'QLib.Std.MXls_Fm_Dta.Function RgzSq(Sq, At As Range) As Range
'QLib.Std.MXls_Fm_Dta.Private Sub ZZ_WszDs()
'QLib.Std.MXls_Fm_Dta.Private Sub ZZ()
'QLib.Std.MXls_Fm_Dta.Private Sub Z()
'QLib.Std.MXls_Fill.Sub FillSeqH(HBar As Range)
'QLib.Std.MXls_Fill.Sub FillSeqV(Vbar As Range)
'QLib.Std.MXls_Fill.Sub FillWsNy(At As Range)
'QLib.Std.MXls_Fun.Sub PutAyColAt(A, At As Range)
'QLib.Std.MXls_Fun.Sub PutAyRgzLc(A, Lo As ListObject, ColNm$)
'QLib.Std.MXls_Fun.Sub PutAyRowAt(Ay, At As Range)
'QLib.Std.MXls_Fun.Function AyabWs(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2", Optional LoNm$ = "Ayab") As Worksheet
'QLib.Std.MXls_Fun.Function NewWsDic(Dic As Dictionary, Optional InclDicValOptTy As Boolean) As Worksheet
'QLib.Std.MXls_Fun.Function NewWsVisDic(A As Dictionary, Optional InclDicValOptTy As Boolean) As Worksheet
'QLib.Std.MXls_Fun.Function NewWsDt(A As Dt, Optional Vis As Boolean) As Worksheet
'QLib.Std.MXls_Fun.Function NyFml(A$) As String()
'QLib.Std.MXls_Fun.Sub SetLcTotLnk(A As ListColumn)
'QLib.Std.MXls_Fun.Function LyWs(Ly$(), Vis As Boolean) As Worksheet
'QLib.Std.MXls_Fun.Property Get MaxWsCol&()
'QLib.Std.MXls_Fun.Property Get MaxWsRow&()
'QLib.Std.MXls_Fun.Function SqHBar(N%) As Variant()
'QLib.Std.MXls_Fun.Function SqVbar(N%) As Variant()
'QLib.Std.MXls_Fun.Function N_ZerFill$(N, NDig%)
'QLib.Std.MXls_Fun.Function WszS1S2Ay(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
'QLib.Std.MXls_Fun.Private Sub Z_AyabWs()
'QLib.Std.MXls_Fun.Private Sub Z_WbFbOupTbl()
'QLib.Std.MXls_Fx.Function BrwFx(Fx)
'QLib.Std.MXls_Fx.Sub CrtFx(A)
'QLib.Std.MXls_Fx.Function FxEns$(Fx)
'QLib.Std.MXls_Fx.Function FstWsn$(Fx)
'QLib.Std.MXls_Fx.Function FxOleCnStr$(A)
'QLib.Std.MXls_Fx.Sub OpnFx(A)
'QLib.Std.MXls_Fx.Sub FxRmvWsIfHas(A, Wsn)
'QLib.Std.MXls_Fx.Function DrsFxq(A, Sql) As Drs
'QLib.Std.MXls_Fx.Sub RunFxq(Fx, Sql)
'QLib.Std.MXls_Fx.Function TmpDbFx(Fx$) As Database
'QLib.Std.MXls_Fx.Function TmpDbzFxww(Fx$, WW) As Database
'QLib.Std.MXls_Fx.Function WbzFx(Fx) As Workbook
'QLib.Std.MXls_Fx.Function WszFxw(Fx, Optional Wsn$ = "Data") As Worksheet
'QLib.Std.MXls_Fx.Function ArsFxwf(A, W, F) As ADODB.Recordset
'QLib.Std.MXls_Fx.Function WsCdNyzFx(Fx) As String()
'QLib.Std.MXls_Fx.Function DtzFxw(Fx, Optional Wsn0$) As Dt
'QLib.Std.MXls_Fx.Function IntAyFxwf(Fx, W, F) As Integer()
'QLib.Std.MXls_Fx.Function WszFxwSy(A, W, Optional F = 0) As String()
'QLib.Std.MXls_Fx.Private Sub ZZ_WsNyzFx()
'QLib.Std.MXls_Fx.Private Sub Z_FstWsn()
'QLib.Std.MXls_Fx.Private Sub Z_TmpDbFx()
'QLib.Std.MXls_Fx.Private Sub Z_WsNyzFx()
'QLib.Std.MXls_Fx.Private Sub ZZ()
'QLib.Std.MXls_Fx.Private Sub Z()
'QLib.Std.MXls_GoWsLnk.Private Sub CrtGoLnkzCell(Cell As Range, Wsn$)
'QLib.Std.MXls_GoWsLnk.Private Function CvCellWsnItm(A) As CellWsnItm
'QLib.Std.MXls_GoWsLnk.Private Function CellWsnItmAy(FstGoCell) As CellWsnItm()
'QLib.Std.MXls_GoWsLnk.Private Function CellWsnItm(Cell As Range, Wsn$) As CellWsnItm
'QLib.Std.MXls_GoWsLnk.Sub CrtGoLnk(FstGoCell As Range)
'QLib.Std.MXls_GoWsLnk.Private Function IsOkToFill(A As Range) As Boolean
'QLib.Std.MXls_GoWsLnk.Sub FillGoWs(FstGoCell As Range)
'QLib.Std.MXls_Lo_Fmt.Function FmtLo(Lo As ListObject, Lof$()) As ListObject
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtAli(L)
'QLib.Std.MXls_Lo_Fmt.Private Function HAlign(S) As XlHAlign
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtBdr(L)
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtBet(L)
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtCor(L)
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtFml(L)
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtFmt(L)
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtLbl(L)
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtLvl(L)
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtTot(L)
'QLib.Std.MXls_Lo_Fmt.Private Function WTotCalczStr(S$) As XlTotalsCalculation
'QLib.Std.MXls_Lo_Fmt.Private Sub WFmtWdt(L)
'QLib.Std.MXls_Lo_Fmt.Private Sub Z_FmtLo()
'QLib.Std.MXls_Lo_Fmt.Private Sub Z_WFmtBdr()
'QLib.Std.MXls_Lo_Fmt.Private Sub ZZ()
'QLib.Std.MXls_Lo_Fmt.Private Function WRg(F) As Range
'QLib.Std.MXls_Lo_Fmt.Private Function WCol(F) As Range
'QLib.Std.MXls_Lo_Fmt.Private Function WLy(T1$) As String()
'QLib.Std.MXls_Lo_Fmt.Private Function WItr(T1$)
'QLib.Std.MXls_Lo_Fmt.Private Function WHdrCell(C) As Range
'QLib.Std.MXls_Lo_Fmt_Tit.Sub SetLoTit(A As ListObject, TitLy$())
'QLib.Std.MXls_Lo_Fmt_Tit.Private Sub MgeTitRg(TitRg As Range)
'QLib.Std.MXls_Lo_Fmt_Tit.Private Sub MgeTitRgH(TitRg As Range)
'QLib.Std.MXls_Lo_Fmt_Tit.Private Sub MgeTitRgV(A As Range)
'QLib.Std.MXls_Lo_Fmt_Tit.Private Function TitAt(Lo As ListObject, NTitRow%) As Range
'QLib.Std.MXls_Lo_Fmt_Tit.Private Function TitSq(TitLy$(), LoFny$()) As Variant()
'QLib.Std.MXls_Lo_Fmt_Tit.Private Sub Z_TitSq()
'QLib.Std.MXls_Lo_Fmt_Tit.Private Sub Z()
'QLib.Std.MXls_Lo_Get_Lo.Function LoAy(A As Workbook) As ListObject()
'QLib.Std.MXls_Lo_Get_Lo.Function LozWs(A As Worksheet, LoNm$) As ListObject 'Return LoOpt
'QLib.Std.MXls_Lo_Get_Lo.Function FstLo(A As Worksheet) As ListObject 'Return LoOpt
'QLib.Std.MXls_Lo_Get_Prp.Sub AddFml(Lo As ListObject, ColNm$, Fml$)
'QLib.Std.MXls_Lo_Get_Prp.Function LoNm$(T)
'QLib.Std.MXls_Lo_Get_Prp.Function CvLo(A) As ListObject
'QLib.Std.MXls_Lo_Get_Prp.Function LoAllCol(A As ListObject) As Range
'QLib.Std.MXls_Lo_Get_Prp.Function LoAllEntCol(A As ListObject) As Range
'QLib.Std.MXls_Lo_Get_Prp.Sub AutoFitLo(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Function BdrLoAround(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Sub BrwLo(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Function RgzLoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
'QLib.Std.MXls_Lo_Get_Prp.Function RgzLc(A As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
'QLib.Std.MXls_Lo_Get_Prp.Function LozWsDta(A As Worksheet, Optional LoNm$) As ListObject
'QLib.Std.MXls_Lo_Get_Prp.Sub DltLo(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Function DrszLo(A As ListObject) As Drs
'QLib.Std.MXls_Lo_Get_Prp.Function DryLo(A As ListObject) As Variant()
'QLib.Std.MXls_Lo_Get_Prp.Function DryRgColAy(Rg As Range, ColIxAy) As Variant()
'QLib.Std.MXls_Lo_Get_Prp.Function DryRgzLoCC(Lo As ListObject, CC) As Variant() ' Return as many column as columns in [CC] from Lo
'QLib.Std.MXls_Lo_Get_Prp.Function DtaAdrzLo$(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Function EntColzLo(A As ListObject, C) As Range
'QLib.Std.MXls_Lo_Get_Prp.Function FbtStrLo$(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Function FnyzLo(A As ListObject) As String()
'QLib.Std.MXls_Lo_Get_Prp.Function HasLoC(Lo As ListObject, C) As Boolean
'QLib.Std.MXls_Lo_Get_Prp.Function HasLoFny(A As ListObject, Fny$()) As Boolean
'QLib.Std.MXls_Lo_Get_Prp.Function LoHasNoDta(A As ListObject) As Boolean
'QLib.Std.MXls_Lo_Get_Prp.Function LoHdrCell(A As ListObject, FldNm) As Range
'QLib.Std.MXls_Lo_Get_Prp.Sub LoKeepFstCol(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Sub LoKeepFstRow(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Function LoNCol%(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Function LoPc(A As ListObject) As PivotCache
'QLib.Std.MXls_Lo_Get_Prp.Function LoQt(A As ListObject) As QueryTable
'QLib.Std.MXls_Lo_Get_Prp.Function R1Lo&(A As ListObject, Optional InclHdr As Boolean)
'QLib.Std.MXls_Lo_Get_Prp.Function R2Lo&(A As ListObject, Optional InclTot As Boolean)
'QLib.Std.MXls_Lo_Get_Prp.Function SqzLo(A As ListObject)
'QLib.Std.MXls_Lo_Get_Prp.Function WsLo(A As ListObject) As Worksheet
'QLib.Std.MXls_Lo_Get_Prp.Function WbLo(A As ListObject) As Workbook
'QLib.Std.MXls_Lo_Get_Prp.Function LoWs(A As ListObject) As Worksheet
'QLib.Std.MXls_Lo_Get_Prp.Function LoWsCno%(A As ListObject, Col)
'QLib.Std.MXls_Lo_Get_Prp.Function LoNmzTblNm$(TblNm)
'QLib.Std.MXls_Lo_Get_Prp.Private Sub ZZ_LoKeepFstCol()
'QLib.Std.MXls_Lo_Get_Prp.Private Sub Z_AutoFitLo()
'QLib.Std.MXls_Lo_Get_Prp.Private Sub Z_BrwLo()
'QLib.Std.MXls_Lo_Get_Prp.Private Sub Z_NewPtLoAtRDCP()
'QLib.Std.MXls_Lo_Get_Prp.Private Sub ZZ()
'QLib.Std.MXls_Lo_Get_Prp.Private Sub Z()
'QLib.Std.MXls_Lo_LofVbl.Function LofVblzQt$(A As QueryTable)
'QLib.Std.MXls_Lo_LofVbl.Property Get LofVblzT$(A As Database, T)
'QLib.Std.MXls_Lo_LofVbl.Property Let LofVblzT(A As Database, T, V$)
'QLib.Std.MXls_Lo_LofVbl.Function LofVblzLo$(A As ListObject)
'QLib.Std.MXls_Lo_LofVbl.Property Get LofVblzFbt$(Fb, T)
'QLib.Std.MXls_Lo_LofVbl.Property Let LofVblzFbt(Fb, T, LofVblzVbl$)
'QLib.Std.MXls_Lo_LofVbl.Function LofVblzFbtStr$(FbtStr$)
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgVal_FmlNotBegWithEq$(Lno&, Fml$)
'QLib.Std.MXls_Lof_ErzLof.Property Get LofT1Ny() As String()
'QLib.Std.MXls_Lof_ErzLof.Function ErzLof(Lof$(), Fny$()) As String() 'Error-of-ListObj-Formatter:Er.z.Lo.f
'QLib.Std.MXls_Lof_ErzLof.Function FnywLikssAy(Fny$(), LikssAy$()) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Sub Init(Lof$(), Fny$())
'QLib.Std.MXls_Lof_ErzLof.Private Property Get WAli_LeftRightCenter() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get WAny_Tot() As Boolean
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErAli() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBdr() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErBdr1(X$) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBdrDup() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBdrExcessFld() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBdrExcessLin() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBdrFld() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBet() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBetDup() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBetFny() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErBetTermCnt() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErCor() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErCorDup(IO$()) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErCorFld(IO$()) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErCorVal(IO$()) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErCorVal1$(L)
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErFld() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErFldss() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErFldSngzDup() As String() 'It is for [SngFldLin] only.  That means T2 of LofLin is field name.  Return error msg for any FldNm is dup.
'QLib.Std.MXls_Lof_ErzLof.Private Function ErFldSngzDup__DupFld_is_fnd(DupFld, LnxAy() As Lnx, T1) As String() '[DupFld] is found within [LnxAy].  All [LnxAy] has [T1]
'QLib.Std.MXls_Lof_ErzLof.Private Function ErFldSngzDup__WithinT1(T1) As String() 'Within T1 any fld is dup?
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErFml() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErFml__InsideFmlHasInvalidFld() As String()
'QLib.Std.MXls_Lof_ErzLof.Function ErFnyzFml(Fld$, Fml$, Fny$()) As String() 'Return Subset-Fny (quote by []) in [Fml] which is error. It is error if any-FmlFny not in [Fny] or =[Fld]
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErFmt() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErLbl() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErLoNm() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErLvl() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErMisFnyzFmti(Fmti) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErTit() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErTot() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErTot_1() '(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As Er
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErVal() As String() 'W-Error-of-LofLinVal:W means working-value. which is using the some Module-Lvl-variables and it is private. Val here means the LofValFld of LofLin
'QLib.Std.MXls_Lof_ErzLof.Private Function ErValOfFml() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErValOfNotBet() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErValOfNotBetz(T1, FmNumVal, ToNumval) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErValOfNotInLis() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function ErValOfNotNum() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Property Get ErWdt() As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function WLnxAyzT1(T1) As Lnx()
'QLib.Std.MXls_Lof_ErzLof.Private Function WMsgzBetTermCnt$(L, NTerm%)
'QLib.Std.MXls_Lof_ErzLof.Private Function WMsgzDupNy(DupNy$(), LnoStrAy$()) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Function WMsgzFny(Fny$(), Lin_Ty$) As String()
'QLib.Std.MXls_Lof_ErzLof.Private Sub Z_ErBet()
'QLib.Std.MXls_Lof_ErzLof.Private Sub Z_ErFldSngzDup()
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Val_NotNum(Lno&, T1$, Val$) As String():                                             MsgOf_Val_NotNum = FmtMacro(M_Val_NotNum, Lno, T1, Val):                                          End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Val_NotBet(Lno&, T1$, Val$, FmNo) As String():                                       MsgOf_Val_NotBet = FmtMacro(M_Val_NotBet, Lno, T1, Val, FmNo):                                    End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Val_NotInLis(Lno&, T1$, ErVal$, VdtValNm$) As String():                              MsgOf_Val_NotInLis = FmtMacro(M_Val_NotInLis, Lno, T1, ErVal, VdtValNm):                          End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Val_FmlFld(Lno&, Fml$, ErFny$(), VdtFny$()) As String():                             MsgOf_Val_FmlFld = FmtMacro(M_Val_FmlFld, Lno, Fml, ErFny, VdtFny):                               End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Val_FmlNotBegEq(Lno&, Fml$) As String():                                             MsgOf_Val_FmlNotBegEq = FmtMacro(M_Val_FmlNotBegEq, Lno, Fml):                                    End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Fld_NotInFny(Lno&, T1$, F) As String():                                              MsgOf_Fld_NotInFny = FmtMacro(M_Fld_NotInFny, Lno, T1, F):                                        End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Fld_Dup(Lno&, T1, F, AlreadyInLno&) As String():                                     MsgOf_Fld_Dup = FmtMacro(M_Fld_Dup, Lno, T1, F, AlreadyInLno):                                    End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Fldss_NotSel(Lno&, T1$, Fldss$) As String():                                         MsgOf_Fldss_NotSel = FmtMacro(M_Fldss_NotSel, Lno, T1, Fldss):                                    End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Fldss_DupSel(Lno&, T1$) As String():                                                 MsgOf_Fldss_DupSel = FmtMacro(M_Fldss_DupSel, Lno, T1):                                           End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_LoNm(Lno&, Val$) As String():                                                        MsgOf_LoNm = FmtMacro(M_LoNm, Lno, Val):                                                          End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_LoNm_Mis() As String():                                                              MsgOf_LoNm_Mis = FmtMacro(M_LoNm_Mis):                                                            End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_LoNm_Dup(Lno&, AlreadyInLno&) As String():                                           MsgOf_LoNm_Dup = FmtMacro(M_LoNm_Dup, Lno, AlreadyInLno):                                         End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Tot_DupSel(Lno&, TotKd$, Fldss$, SelFld$, AlreadyInLno&, AlreadyTotKd$) As String(): MsgOf_Tot_DupSel = FmtMacro(M_Tot_DupSel, Lno, TotKd, Fldss, SelFld, AlreadyInLno, AlreadyTotKd): End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Bet_N3Fld(Lno&) As String():                                                          MsgOf_Bet_N3Fld = FmtMacro(M_Bet_N3Fld, Lno):                                                       End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Bet_EqFmTo(Lno&) As String():                                                        MsgOf_Bet_EqFmTo = FmtMacro(M_Bet_EqFmTo, Lno):                                                   End Function
'QLib.Std.MXls_Lof_ErzLof.Private Function MsgOf_Bet_FldSeq(Lno&) As String():                                                        MsgOf_Bet_FldSeq = FmtMacro(M_Bet_FldSeq, Lno):                                                   End Function
'QLib.Std.MXls_Lo_Minx.Sub MinxLo(A As ListObject)
'QLib.Std.MXls_Lo_Minx.Private Sub MinxLozWs(A As Worksheet)
'QLib.Std.MXls_Lo_Minx.Function MinxLozWszWb(A As Workbook) As Workbook
'QLib.Std.MXls_Lo_Minx.Sub MinxLozWszFx(A)
'QLib.Std.MXls_Lo_Samp.Property Get SampLoVis() As ListObject
'QLib.Std.MXls_Lo_Samp.Property Get SampLo() As ListObject
'QLib.Std.MXls_Lo_Samp.Property Get SampLof() As String()
'QLib.Std.MXls_Lo_Samp.Property Get SampLofTp() As String()
'QLib.Std.MXls_Lo_Samp.Property Get SampDrzAToJ() As Variant()
'QLib.Std.MXls_Lo_Samp.Property Get SampSq1() As Variant()
'QLib.Std.MXls_Lo_Samp.Property Get SampSqWithHdr() As Variant()
'QLib.Std.MXls_Lo_Samp.Property Get SampWs() As Worksheet
'QLib.Std.MXls_Lo_Set.Function WszLo(A As ListObject) As Worksheet
'QLib.Std.MXls_Lo_Set.Function LoSetNm(A As ListObject, LoNm$) As ListObject
'QLib.Std.MXls_New.Function NewA1(Optional Wsn$, Optional Vis As Boolean) As Range
'QLib.Std.MXls_New.Function NewWb(Optional Wsn$) As Workbook
'QLib.Std.MXls_New.Function NewWs(Optional Wsn$) As Worksheet
'QLib.Std.MXls_New.Function NewXls() As Excel.Application
'QLib.Std.MXls_Pt.Function PtCpyToLo(A As PivotTable, At As Range) As ListObject
'QLib.Std.MXls_Pt.Sub SetPtffOri(A As PivotTable, FF, Ori As XlPivotFieldOrientation)
'QLib.Std.MXls_Pt.Private Sub FmtPt(Pt As PivotTable)
'QLib.Std.MXls_Pt.Function WbNm$(A As Workbook)
'QLib.Std.MXls_Pt.Function LasWb() As Workbook
'QLib.Std.MXls_Pt.Function LasWbz(A As Excel.Application) As Workbook
'QLib.Std.MXls_Pt.Sub ShwWb(Wb As Workbook)
'QLib.Std.MXls_Pt.Sub ThwHasWbzWs(Wb As Workbook, Wsn$, Fun$)
'QLib.Std.MXls_Pt.Function PtzRg(A As Range, Optional Wsn$, Optional PtNm$) As PivotTable
'QLib.Std.MXls_Pt.Function PivCol(Pt As PivotTable, PivColNm) As PivotField
'QLib.Std.MXls_Pt.Function PivRow(Pt As PivotTable, PivRowNm) As PivotField
'QLib.Std.MXls_Pt.Function PivFld(A As PivotTable, F) As PivotField
'QLib.Std.MXls_Pt.Function ColEntPt(A As PivotTable, PivColNm) As Range
'QLib.Std.MXls_Pt.Function PivColEnt(Pt As PivotTable, ColNm) As Range
'QLib.Std.MXls_Pt.Sub SetPtWdt(A As PivotTable, Colss$, ColWdt As Byte)
'QLib.Std.MXls_Pt.Sub SetPtOutLin(A As PivotTable, Colss$, Optional Lvl As Byte = 2)
'QLib.Std.MXls_Pt.Sub SetPtRepeatLbl(A As PivotTable, Rowss$)
'QLib.Std.MXls_Pt.Sub ShwPt(A As PivotTable)
'QLib.Std.MXls_Pt.Function WbPt(A As PivotTable) As Workbook
'QLib.Std.MXls_Pt.Function WsPt(A As PivotTable) As Worksheet
'QLib.Std.MXls_Pt.Function SampPt() As PivotTable
'QLib.Std.MXls_Pt.Function SampRg() As Range
'QLib.Std.MXls_Pt.Function RgSetVis(A As Range, Vis As Boolean) As Range
'QLib.Std.MXls_Pt.Sub SetAppVis(A As Excel.Application, Vis As Boolean)
'QLib.Std.MXls_Pt.Function AtAddSq(At As Range, Sq()) As Range
'QLib.Std.MXls_Pt.Function NewPtLoAtRDCP(A As ListObject, At As Range, Rowss$, Dtass$, Optional Colss$, Optional Pagss$) As PivotTable
'QLib.Std.MXls_Pt.Function NewPtLoAtRDCPNm$(A As ListObject)
'QLib.Std.MXls_Qt.Function FbtStrQt$(A As QueryTable)
'QLib.Std.MXls_Rfh.Private Sub ClsWc(A As WorkbookConnection)
'QLib.Std.MXls_Rfh.Private Sub ClsWczWb(Wb As Workbook)
'QLib.Std.MXls_Rfh.Private Sub SetWczFb(A As WorkbookConnection, ToUseFb)
'QLib.Std.MXls_Rfh.Private Sub RfhWc(A As WorkbookConnection, ToUseFb)
'QLib.Std.MXls_Rfh.Private Sub RfhPc(A As PivotCache)
'QLib.Std.MXls_Rfh.Sub RfhFx(Fx, Fb$)
'QLib.Std.MXls_Rfh.Private Sub RfhWs(A As Worksheet)
'QLib.Std.MXls_Rfh.Function RfhWb(Wb As Workbook, Fb) As Workbook
'QLib.Std.MXls_Rfh.Private Sub RplLozFb(Wb As Workbook, Fb)
'QLib.Std.MXls_Rfh.Private Function RplLozT(A As ListObject, Db As Database, T) As ListObject
'QLib.Std.MXls_Rfh.Private Function OupLoAy(A As Workbook) As ListObject()
'QLib.Std.MXls_Rg.Function CvRg(A) As Range
'QLib.Std.MXls_Rg.Function RgA1(A As Range) As Range
'QLib.Std.MXls_Rg.Function A1(Ws As Worksheet) As Range
'QLib.Std.MXls_Rg.Function A1zRg(A As Range) As Range
'QLib.Std.MXls_Rg.Function IsA1(A As Range) As Boolean
'QLib.Std.MXls_Rg.Function WsRgAdr$(A As Range)
'QLib.Std.MXls_Rg.Sub AsgRRRCCRg(A As Range, OR1, OR2, OC1, OC2)
'QLib.Std.MXls_Rg.Sub BdrRg(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
'QLib.Std.MXls_Rg.Sub BdrRgAround(A As Range)
'QLib.Std.MXls_Rg.Sub BdrRgBottom(A As Range)
'QLib.Std.MXls_Rg.Sub BdrRgInner(A As Range)
'QLib.Std.MXls_Rg.Sub BdrRgInside(A As Range)
'QLib.Std.MXls_Rg.Sub BdrRgAlign(A As Range, H As XlHAlign)
'QLib.Std.MXls_Rg.Sub BdrRgLeft(A As Range)
'QLib.Std.MXls_Rg.Sub BdrRgRight(A As Range)
'QLib.Std.MXls_Rg.Sub BdrRgTop(A As Range)
'QLib.Std.MXls_Rg.Function RgC(A As Range, C) As Range
'QLib.Std.MXls_Rg.Function RgCC(A As Range, C1, C2) As Range
'QLib.Std.MXls_Rg.Function RgCRR(A As Range, C, R1, R2) As Range
'QLib.Std.MXls_Rg.Function EntRgC(A As Range, C) As Range
'QLib.Std.MXls_Rg.Function EntRgRR(A As Range, R1, R2) As Range
'QLib.Std.MXls_Rg.Function FstColRg(A As Range) As Range
'QLib.Std.MXls_Rg.Function FstRowRg(A As Range) As Range
'QLib.Std.MXls_Rg.Function IsHBarRg(A As Range) As Boolean
'QLib.Std.MXls_Rg.Function IsVbarRg(A As Range) As Boolean
'QLib.Std.MXls_Rg.Function LasColRg%(A As Range)
'QLib.Std.MXls_Rg.Function LasHBarRg(A As Range) As Range
'QLib.Std.MXls_Rg.Function LasRowRg&(A As Range)
'QLib.Std.MXls_Rg.Function LasVbarRg(A As Range) As Range
'QLib.Std.MXls_Rg.Function LozSq(Sq(), At As Range, Optional LoNm$) As ListObject
'QLib.Std.MXls_Rg.Function LozRg(Rg As Range, Optional LoNm$) As ListObject
'QLib.Std.MXls_Rg.Sub MgeRg(A As Range)
'QLib.Std.MXls_Rg.Function NColRg%(A As Range)
'QLib.Std.MXls_Rg.Function RgzMoreBelow(A As Range, Optional N% = 1)
'QLib.Std.MXls_Rg.Function RgzMoreTop(A As Range, Optional N = 1)
'QLib.Std.MXls_Rg.Function NRowRg&(A As Range)
'QLib.Std.MXls_Rg.Function RgR(A As Range, R)
'QLib.Std.MXls_Rg.Function CellBelow(Cell As Range, Optional N = 1) As Range
'QLib.Std.MXls_Rg.Sub SwapValOfRg(Cell1 As Range, Cell2 As Range)
'QLib.Std.MXls_Rg.Function CellAbove(Cell As Range, Optional Above = 1) As Range
'QLib.Std.MXls_Rg.Function CellRight(A As Range, Optional Right = 1) As Range
'QLib.Std.MXls_Rg.Function RgRC(A As Range, R, C) As Range
'QLib.Std.MXls_Rg.Function RgRCC(A As Range, R, C1, C2) As Range
'QLib.Std.MXls_Rg.Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
'QLib.Std.MXls_Rg.Function RgRR(A As Range, R1, R2) As Range
'QLib.Std.MXls_Rg.Function RgzResz(At As Range, Sq) As Range
'QLib.Std.MXls_Rg.Function SqzRg(A As Range) As Variant()
'QLib.Std.MXls_Rg.Function WbzRg(A As Range) As Workbook
'QLib.Std.MXls_Rg.Function WszRg(A As Range) As Worksheet
'QLib.Std.MXls_Rg.Private Sub Z_RgzMoreBelow()
'QLib.Std.MXls_Rg.Private Sub Z()
'QLib.Std.MXls_RgCell.Function CellAyH(A As Range) As Variant()
'QLib.Std.MXls_RgCell.Sub CellClrDown(A As Range)
'QLib.Std.MXls_RgCell.Sub CellFillSeqDown(A As Range, FmNum&, ToNum&)
'QLib.Std.MXls_RgCell.Function IsCellInRg(A As Range, Rg As Range) As Boolean
'QLib.Std.MXls_RgCell.Function IsCellInRgAp(Cell As Range, ParamArray RgAp()) As Boolean
'QLib.Std.MXls_RgCell.Function IsCellInRgAv(A As Range, RgAv()) As Boolean
'QLib.Std.MXls_RgCell.Sub MgeCellAbove(Cell As Range)
'QLib.Std.MXls_RgCell.Function VbarRgAt(At As Range, Optional AtLeastOneCell As Boolean) As Range
'QLib.Std.MXls_RgVBar.Sub Vbar_MgeBottomEmpCell(A As Range)
'QLib.Std.MXls_RgVBar.Function VbarAy(A As Range) As Variant()
'QLib.Std.MXls_RgVBar.Function VbarIntAy(A As Range) As Integer()
'QLib.Std.MXls_RgVBar.Function VbarSy(A As Range) As String()
'QLib.Std.MXls_Sq.Function NewSq(R&, C&) As Variant()
'QLib.Std.MXls_Sq.Function SqAddSngQuote(A)
'QLib.Std.MXls_Sq.Sub BrwSq(A)
'QLib.Std.MXls_Sq.Function Sq_Col(A, C%) As Variant()
'QLib.Std.MXls_Sq.Function IntoSqC(A, C%, Into) As String()
'QLib.Std.MXls_Sq.Function SyzSq(Sq, Optional C% = 0) As String()
'QLib.Std.MXls_Sq.Function DrzSqr(Sq, R) As Variant()
'QLib.Std.MXls_Sq.Function SqInsDr(A, Dr, Optional Row& = 1)
'QLib.Std.MXls_Sq.Function IsEmpSq(Sq) As Boolean
'QLib.Std.MXls_Sq.Function IsEqSq(A, B) As Boolean
'QLib.Std.MXls_Sq.Function LySq(A) As String()
'QLib.Std.MXls_Sq.Function NColSq&(A)
'QLib.Std.MXls_Sq.Function NewLoSqAt(Sq(), At As Range) As ListObject
'QLib.Std.MXls_Sq.Function NewLoSq(Sq(), Optional Wsn$ = "Data") As ListObject
'QLib.Std.MXls_Sq.Function WszSq(Sq(), Optional Wsn$) As Worksheet
'QLib.Std.MXls_Sq.Function NRowSq&(A)
'QLib.Std.MXls_Sq.Sub SetSqrzDr(OSq, R, Dr, Optional NoTxtSngQ As Boolean)
'QLib.Std.MXls_Sq.Function SqSyz(A) As String()
'QLib.Std.MXls_Sq.Function SqTranspose(A) As Variant()
'QLib.Std.MXls_Sq.Private Sub ZZ()
'QLib.Std.MXls_Sq.Property Get SampSq() As Variant()
'QLib.Std.MXls_TreeWs.Sub Change(Target As Range)
'QLib.Std.MXls_TreeWs.Sub SelectionChange(Target As Range)
'QLib.Std.MXls_TreeWs.Private Sub ShwEntzHom(Hom$)
'QLib.Std.MXls_TreeWs.Private Sub ShwEnt(Pth$, Cno%)
'QLib.Std.MXls_TreeWs.Private Sub ShwEntzPut(Cno%, FdrAy$(), FnAy$())
'QLib.Std.MXls_TreeWs.Private Sub ShwFstHomFdr(Hom$)
'QLib.Std.MXls_TreeWs.Private Sub ShwCurCol(Cur As Range)
'QLib.Std.MXls_TreeWs.Private Sub ShwNxtEnt()
'QLib.Std.MXls_TreeWs.Private Sub ShwCurEnt(Cur As Range)
'QLib.Std.MXls_TreeWs.Private Function PthzCur$(Cur As Range)
'QLib.Std.MXls_TreeWs.Private Sub PutCurEnt(Cur As Range, SubPthAy$(), FnAy$())
'QLib.Std.MXls_TreeWs.Private Function EntRg(Cur As Range, EntCnt%) As Range
'QLib.Std.MXls_TreeWs.Private Sub ClrCurCol(Cur As Range)
'QLib.Std.MXls_TreeWs.Private Function CurColCC() As Range
'QLib.Std.MXls_TreeWs.Private Function MgeCurSubPthCol(SubPthSz&)
'QLib.Std.MXls_TreeWs.Private Function MgeCurFnCol(SubPthSz&, FnSz&)
'QLib.Std.MXls_TreeWs.Private Sub ShwNxtCol(Cur As Range)
'QLib.Std.MXls_TreeWs.Private Sub ShwRow(Cur As Range)
'QLib.Std.MXls_TreeWs.Private Function MaxR%(Ws As Worksheet)
'QLib.Std.MXls_TreeWs.Private Function MaxC%(Ws As Worksheet)
'QLib.Std.MXls_TreeWs.Private Sub EnsA1(A1 As Range)
'QLib.Std.MXls_TreeWs.Private Sub Clear(Ws As Worksheet)
'QLib.Std.MXls_TreeWs.Private Function IsAction(Ws As Worksheet) As Boolean
'QLib.Std.MXls_TreeWs.Private Function IsActionWs(Ws As Worksheet) As Boolean
'QLib.Std.MXls_TreeWs.Private Function IsActionA1(A1 As Range) As Boolean
'QLib.Std.MXls_TreeWs_Install.Function TreeWsMdLines$()
'QLib.Std.MXls_TreeWs_Install.Sub InstallTreeWs()
'QLib.Std.MXls_TreeWs_Install.Sub InstallTreeWbz(Wb As Workbook)
'QLib.Std.MXls_TreeWs_Install.Function IsTreeWb(A As Workbook) As Boolean
'QLib.Std.MXls_TreeWs_Install.Function TreeWbAy() As Workbook()
'QLib.Std.MXls_TreeWs_Install.Function TreeWsAy() As Worksheet()
'QLib.Std.MXls_TreeWs_Install.Sub InstallTreeWsz(A As Worksheet)
'QLib.Std.MXls_Wb.Property Get CurWb() As Workbook
'QLib.Std.MXls_Wb.Function CvWb(A) As Workbook
'QLib.Std.MXls_Wb.Function FstWs(A As Workbook) As Worksheet
'QLib.Std.MXls_Wb.Function FxWb$(A As Workbook)
'QLib.Std.MXls_Wb.Function LasWs(A As Workbook) As Worksheet
'QLib.Std.MXls_Wb.Function LoItr(A As Workbook)
'QLib.Std.MXls_Wb.Function LozAyH(Ay, Wb As Workbook, Optional Wsn$, Optional LoNm$) As ListObject
'QLib.Std.MXls_Wb.Function MainLo(A As Workbook) As ListObject
'QLib.Std.MXls_Wb.Function MainQt(A As Workbook) As QueryTable
'QLib.Std.MXls_Wb.Function MainWs(A As Workbook) As Worksheet
'QLib.Std.MXls_Wb.Function Wbs(A As Workbook) As Workbooks
'QLib.Std.MXls_Wb.Function PtNy(A As Workbook) As String()
'QLib.Std.MXls_Wb.Function TxtWc(A As Workbook) As TextConnection
'QLib.Std.MXls_Wb.Function TxtWcCnt%(A As Workbook)
'QLib.Std.MXls_Wb.Function TxtWcStr$(A As Workbook)
'QLib.Std.MXls_Wb.Function OleWcAy(A As Workbook) As OLEDBConnection()
'QLib.Std.MXls_Wb.Function WcNyWb(A As Workbook) As String()
'QLib.Std.MXls_Wb.Function WcStrAyWbOLE(A As Workbook) As String()
'QLib.Std.MXls_Wb.Function WszWb(A As Workbook, Wsn) As Worksheet
'QLib.Std.MXls_Wb.Function WsNyzRg(A As Range) As String()
'QLib.Std.MXls_Wb.Function WsNy(A As Workbook) As String()
'QLib.Std.MXls_Wb.Function WszCdNm(A As Workbook, WsCdNm$) As Worksheet
'QLib.Std.MXls_Wb.Function WsCdNy(A As Workbook) As String()
'QLib.Std.MXls_Wb.Function WbFullNm$(A As Workbook)
'QLib.Std.MXls_Wb.Function WbAddTT(A As Workbook, Db As Database, TT, Optional UseWc As Boolean) As Workbook
'QLib.Std.MXls_Wb.Function WszWbDt(A As Workbook, Dt As Dt) As Worksheet
'QLib.Std.MXls_Wb.Function WczWbFb(A As Workbook, LnkToFb$, WcNm) As WorkbookConnection
'QLib.Std.MXls_Wb.Function AddWs(A As Workbook, Optional Wsn$, Optional AtBeg As Boolean, Optional AtEnd As Boolean, Optional BefWsn$, Optional AftWsn$) As Worksheet
'QLib.Std.MXls_Wb.Sub ThwWbMisOupNy(A As Workbook, OupNy$())
'QLib.Std.MXls_Wb.Sub ClsWbNoSav(A As Workbook)
'QLib.Std.MXls_Wb.Sub DltWc(A As Workbook)
'QLib.Std.MXls_Wb.Sub DltWs(A As Workbook, Wsn)
'QLib.Std.MXls_Wb.Function WbMax(A As Workbook) As Workbook
'QLib.Std.MXls_Wb.Function NewA1Wb(A As Workbook, Optional Wsn$) As Range
'QLib.Std.MXls_Wb.Sub WbQuit(A As Workbook)
'QLib.Std.MXls_Wb.Function WbSav(A As Workbook) As Workbook
'QLib.Std.MXls_Wb.Function WbSavAs(A As Workbook, Fx, Optional Fmt As XlFileFormat = xlOpenXMLWorkbook) As Workbook
'QLib.Std.MXls_Wb.Sub SetWbFcsvCn(A As Workbook, Fcsv$)
'QLib.Std.MXls_Wb.Function HasWbzWs(A As Workbook, Wsn) As Boolean
'QLib.Std.MXls_Wb.Private Sub ZZ_WbWcSy()
'QLib.Std.MXls_Wb.Private Sub ZZ_LozAyH()
'QLib.Std.MXls_Wb.Private Sub Z_TxtWcCnt()
'QLib.Std.MXls_Wb.Private Sub Z_SetWbFcsvCn()
'QLib.Std.MXls_Wb.Private Sub ZZ()
'QLib.Std.MXls_Wb.Private Sub Z()
'QLib.Std.MXls_Wc.Function WszWc(Wc As WorkbookConnection) As Worksheet
'QLib.Std.MXls_Wc.Sub RgzWc(Wc As WorkbookConnection, At As Range)
'QLib.Std.MXls_Wc.Sub AddWcTpWFb()
'QLib.Std.MXls_Wc.Sub AddWcFxFbtt(Fx, LnkFb$, TT)
'QLib.Std.MXls_Wc.Private Function WbzFbOupTbl(Fb) As Workbook
'QLib.Std.MXls_Wc.Sub CrtFxzFbOupTbl(Fx$, Fb$)
'QLib.Std.MXls_Ws.Sub ShwWs(A As Worksheet)
'QLib.Std.MXls_Ws.Function WsAdd(Wb As Workbook, Optional Wsn$) As Worksheet
'QLib.Std.MXls_Ws.Function WsC(A As Worksheet, C) As Range
'QLib.Std.MXls_Ws.Function WsCC(A As Worksheet, C1, C2) As Range
'QLib.Std.MXls_Ws.Sub DltLo(A As Worksheet)
'QLib.Std.MXls_Ws.Sub ClsWsNoSav(A As Worksheet)
'QLib.Std.MXls_Ws.Property Get CurWs() As Worksheet
'QLib.Std.MXls_Ws.Function WsCRR(A As Worksheet, C, R1, R2) As Range
'QLib.Std.MXls_Ws.Function DltWs(A As Workbook, WsIx) As Boolean
'QLib.Std.MXls_Ws.Function RgzWs(A As Worksheet) As Range
'QLib.Std.MXls_Ws.Function HasLo(A As Worksheet, LoNm$) As Boolean
'QLib.Std.MXls_Ws.Function LasCell(A As Worksheet) As Range
'QLib.Std.MXls_Ws.Function LasCno%(A As Worksheet)
'QLib.Std.MXls_Ws.Function LasRno&(A As Worksheet)
'QLib.Std.MXls_Ws.Function PtNyzWs(A As Worksheet) As String()
'QLib.Std.MXls_Ws.Function WsRC(A As Worksheet, R, C) As Range
'QLib.Std.MXls_Ws.Function WsRCC(A As Worksheet, R, C1, C2) As Range
'QLib.Std.MXls_Ws.Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
'QLib.Std.MXls_Ws.Function WsRR(A As Worksheet, R1, R2) As Range
'QLib.Std.MXls_Ws.Function SetWsNm(A As Worksheet, Nm$) As Worksheet
'QLib.Std.MXls_Ws.Function SqzWs(A As Worksheet) As Variant()
'QLib.Std.MXls_Ws.Function WsSetVis(A As Worksheet, Vis As Boolean) As Worksheet
'QLib.Std.MXls_Ws.Function A1Wb(A As Workbook, Optional Wsn$) As Range
'QLib.Std.MXls_Ws.Function A1zWs(A As Worksheet) As Range
'QLib.Std.MXls_Ws.Function CvWs(A) As Worksheet
'QLib.Std.MXls_Ws.Function WbzWs(A As Worksheet) As Workbook
'QLib.Std.MXls_Ws.Function WbNmzWs$(A As Worksheet)
'QLib.Std.MXls_Ws.Sub DltColFm(Ws As Worksheet, FmCol)
'QLib.Std.MXls_Ws.Sub DltRowFm(Ws As Worksheet, FmRow)
'QLib.Std.MXls_Ws.Sub HidColFm(Ws As Worksheet, FmCol)
'QLib.Std.MXls_Ws.Sub HidRowFm(Ws As Worksheet, FmRow&)
'QLib.Std.MXls_Ws.Function CnozBefFstHid%(Ws As Worksheet)
'QLib.Std.MXls_Xls.Private Sub Z_XlszGetObj()
'QLib.Std.MXls_Xls.Function XlszGetObj() As Excel.Application
'QLib.Std.MXls_Xls.Function Xls() As Excel.Application
'QLib.Std.MXls_Xls.Function HasAddinFn(A As Excel.Application, AddinFn) As Boolean
'QLib.Std.MXls_Xls.Sub XlsQuit(A As Excel.Application)
'QLib.Std.MXls_Xls.Sub ClsAllWb(A As Excel.Application)
'QLib.Std.MXls_Xls.Function DftXls(A As Excel.Application) As Excel.Application
'QLib.Cls.Rel.Friend Function Init(RelLy$()) As Rel
'QLib.Cls.Rel.Sub PushParChd(P, C)
'QLib.Cls.Rel.Sub PushRelLin(A)
'QLib.Cls.Rel.Property Get CycParDotChdAy() As String()
'QLib.Cls.Rel.Property Get IsCyc() As Boolean
'QLib.Cls.Rel.Property Get MulChdRel() As Rel
'QLib.Cls.Rel.Property Get Srt() As Rel
'QLib.Cls.Rel.Property Get SwapParChd() As Rel
'QLib.Cls.Rel.Sub Vc()
'QLib.Cls.Rel.Sub Brw()
'QLib.Cls.Rel.Function Clone() As Rel
'QLib.Cls.Rel.Sub Dmp()
'QLib.Cls.Rel.Property Get Fmt() As String()
'QLib.Cls.Rel.Function IsEq(A As Rel) As Boolean
'QLib.Cls.Rel.Sub ThwIfNE(A As Rel, Optional Msg$ = "Two rel are diff", Optional ANm$ = "Rel-B")
'QLib.Cls.Rel.Sub ThwNotVdt()
'QLib.Cls.Rel.Property Get NItm&()
'QLib.Cls.Rel.Function IsLeaf(Itm) As Boolean
'QLib.Cls.Rel.Function IsNoChdPar(Itm) As Boolean
'QLib.Cls.Rel.Function IsPar(Itm) As Boolean
'QLib.Cls.Rel.Function SetOfItm() As Aset
'QLib.Cls.Rel.Function InDpdOrdItms() As Aset
'QLib.Cls.Rel.Function SetOfPar() As Aset
'QLib.Cls.Rel.Function SetOfLeaf() As Aset
'QLib.Cls.Rel.Function NoChdPar() As Aset
'QLib.Cls.Rel.Sub ThwIfNotPar(Par, Fun$)
'QLib.Cls.Rel.Property Get NPar&()
'QLib.Cls.Rel.Function ParHasChd(P, C) As Boolean
'QLib.Cls.Rel.Function ParChd(P) As Aset
'QLib.Cls.Rel.Function ParIsNoChd(P) As Boolean
'QLib.Cls.Rel.Function ParLin$(P)
'QLib.Cls.Rel.Function RmvChdAy&(P, ChdAy())
'QLib.Cls.Rel.Function RmvChd(P, C) As Boolean
'QLib.Cls.Rel.Property Get SetOfChd() As Aset
'QLib.Cls.Rel.Function RmvAllLeaf&()
'QLib.Cls.Rel.Function RmvNoChdPar&()
'QLib.Cls.Rel.Function RmvPar(P) As Boolean
'QLib.Cls.Rel.Property Get SampRel() As Rel
'QLib.Cls.Rel.Friend Sub Z_Itms()
'QLib.Cls.Rel.Friend Sub Z_InDpdOrdItms()
'QLib.Cls.Rel.Private Sub ZZ()
'QLib.Cls.Rel.Friend Sub Z()
'QLib.Cls.Rel.Property Get SetOfSngChdPar() As Aset
'QLib.Cls.RRCC.Friend Function Init(R1, R2, C1, C2) As RRCC
'QLib.Cls.RRCC.Property Get IsEmp() As Boolean
'QLib.Cls.S1S2.Friend Function Init(S1, S2) As S1S2
'QLib.Cls.S1S2.Property Get ToStr$()
'QLib.Cls.SampS1S2.Private Sub X(O1$(), O2$())
'QLib.Cls.SampS1S2.Property Get S1S2AyzLines() As S1S2()
'QLib.Cls.SampS1S2.Property Get S1S2AyzLin() As S1S2()
'QLib.Cls.SampSqt.Private Property Get PmLy() As String()
'QLib.Cls.SampSqt.Property Get Pm() As Dictionary
'QLib.Cls.SampSqt.Property Get SwLnxAy() As Lnx()
'QLib.Cls.SampSqt.Property Get Sw() As Dictionary
'QLib.Cls.SampSqt.Property Get FldSw() As Dictionary
'QLib.Cls.SampSqt.Property Get StmtSw() As Dictionary
'QLib.Cls.SampSqt.Property Get SqTp$()
'QLib.Cls.SwBrk.Friend Property Get TermAy() As String()
'QLib.Cls.SwBrk.Friend Property Let TermAy(A$())
'QLib.Cls.SwBrk.Friend Function Init(Ix%, Nm$, OpStr$, TermAy$()) As SwBrk
'QLib.Cls.SwBrk.Property Get Lin$()
'QLib.Cls.SyPair.Function Init(Sy1, Sy2) As SyPair
'QLib.Cls.SyPair.Property Get Sy1() As String()
'QLib.Cls.SyPair.Property Get Sy2() As String()
'QLib.Cls.TpPos.Property Get Lin$()
'QLib.Cls.WhMd.Function Init(CmpTy() As vbext_ComponentType, Nm As WhNm) As WhMd
'QLib.Cls.WhMd.Property Get CmpTy() As vbext_ComponentType()
'QLib.Cls.WhMth.Function Init(ShtMdy$(), ShtKd$(), Nm As WhNm) As WhMth
'QLib.Cls.WhMth.Property Get WhNm() As WhNm
'QLib.Cls.WhMth.Property Get IsEmp() As Boolean
'QLib.Cls.WhMth.Property Get ShtKdAy() As String()
'QLib.Cls.WhMth.Property Get ShtMthMdyAy() As String()
'QLib.Cls.WhMth.Property Get ToStr$()
'QLib.Cls.WhNm.Property Get IsEmp() As Boolean
'QLib.Cls.WhNm.Friend Function Init(Patn$, LikeAy$(), ExlLikAy$()) As WhNm
'QLib.Cls.WhNm.Property Get Re() As RegExp
'QLib.Cls.WhNm.Property Get LikeAy() As String()
'QLib.Cls.WhNm.Property Get ExlLikAy() As String()
'QLib.Cls.WhNm.Property Get ToStr$()
'QLib.Std.MIde_Pj_Backup.Sub BackupPj()
'QLib.Std.MIde_Pj_Backup.Function PjfBackupzPj$(A As VBProject)
'QLib.Std.MDao_Db_Get.Function LngAyzQ(A As Database, Q) As Long()
'QLib.Std.MDao_Db_Get.Function SyzQ(A As Database, Q) As String()
'QLib.Std.MDao_Db_Get.Private Sub ZZ_Rs()
'QLib.Std.MDao_Db_Get_Dta.Function DrzQ(A As Database, Q) As Variant()
'QLib.Std.MDao_Db_Get_Dta.Function DryzQ(A As Database, Q) As Variant()
'QLib.Std.MDao_Db_Get_Fny.Function FnyzQ(A As Database, Q) As String()
'QLib.Std.MDao_Db_Get_Fny.Private Sub Z_FnyzQ()
'QLib.Std.MDao_Db_Brw.Sub BrwQ(A As Database, Q)
'QLib.Std.MDao_Db_Run.Sub RunSqy(A As Database, Sqy$())
'QLib.Std.MDao_Db_Get_Col.Function IntAyzQ(A As Database, Q) As Integer()
'QLib.Std.MDao_Db_Get_Col.Function SyzTF(A As Database, T, F) As String()
'QLib.Std.MDao_Db_Get_Col.Function IntozTF(OInto, A As Database, T, F)
'QLib.Std.MIde_Mth_Nm_Dup_X.Function SamMthLinesMthDNmDry(MthQNmLDrs As Drs, Vbe As Vbe) As Variant()
'QLib.Std.MIde_Mth_Nm_Dup_X.Private Function IfShwNoDupMsg(MthDNy$(), MthNm) As Boolean
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNyGp_IsDup(Ny) As Boolean
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNyGp_IsVdt(A) As Boolean
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNyGpAyAllSameCnt%(A)
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupPjLinesIdMthNy(A As VBProject) As String()
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNmDrsPj() As Drs
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNmDrszPj(A As VBProject) As Drs
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNmDryPj() As Variant()
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNmDryzPj(A As VBProject) As Variant()
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupIxAyzDry(Dry(), CC) As Long()
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNmDryVbe() As Variant()
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNmDryzMthQNy(MthQNy$()) As Variant()
'QLib.Std.MIde_Mth_Nm_Dup_X.Function DupMthQNmDryzVbe(A As Vbe) As Variant()
'QLib.Std.MIde_Mth_Nm_Dup_X.Private Sub Z()
'QLib.Std.MDta_Dry_ReSzToSamColCnt.Function DryReSzToSamColCnt(Dry()) As Variant()
'QLib.Std.MDta_Col_Get.Function ColzDrs(A As Drs, ColNm$) As Variant()
'QLib.Std.MDta_Col_Get.Function StrColzDrs(A As Drs, ColNm$) As String()
'QLib.Std.MDta_Col_Get.Function DrLinAy(Dry(), Optional CC, Optional FldSep$ = vbFldSep) As String()
'QLib.Std.MDta_Col_Get.Function DrLin$(Dr, Optional CC, Optional FldSep$ = vbFldSep)
'QLib.Std.MDta_Col_Get.Function SqzDry(A()) As Variant()
'QLib.Std.MDta_Col_Get.Function StrColzDry(Dry(), C) As String()
'QLib.Std.MDta_Col_Get.Function SqzDrySkip(A(), SkipNRow%)
'QLib.Std.MDta_Col_Get.Function IntAyDryC(A(), C) As Integer()
'QLib.Std.MDta_Col_Get.Function ColzDry(Dry(), C) As Variant()
'QLib.Std.MDta_Col_Get.Function IntoColzDry(Into, Dry(), C)
'QLib.Std.MDta_ValId.Function AddColzValIdzCntzDrs(A As Drs, ColNm$, Optional ColNmPfx$) As Drs
'QLib.Std.MDta_ValId.Function AddColzValIdzCntzDry(A(), ValColIx) As Variant()
'QLib.Std.MDta_Col_Add.Function DrsAddColzNmVy(A As Drs, ColNm$, ColVy) As Drs
'QLib.Std.MDta_Col_Add.Private Function DryAddColzColVy(Dry(), ColVy, AtIx&) As Variant()
'QLib.Std.MDta_Col_Add.Function DrsAddColzMap(A As Drs, NewFldEqFunQuoteFmFldSsl$) As Drs
'QLib.Std.MDao_Lid_Mis_Samp.Function SampLidMis() As LidMis
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ffn() As Aset
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Tbl() As LidMisTbl()
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Col() As LidMisCol()
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty() As LidMisTy()
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty1() As LidMisTy
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty2() As LidMisTy
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty3() As LidMisTy
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty4() As LidMisTy
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty4Col() As LidMisTyc()
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty3Col() As LidMisTyc()
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty2Col() As LidMisTyc()
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty1Col() As LidMisTyc()
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty1Col1() As LidMisTyc
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty1Col2() As LidMisTyc
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty2Col1() As LidMisTyc
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty2Col2() As LidMisTyc
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty3Col1() As LidMisTyc
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty3Col2() As LidMisTyc
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty4Col1() As LidMisTyc
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Ty4Col2() As LidMisTyc
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Col1() As LidMisCol
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Col2() As LidMisCol
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Col3() As LidMisCol
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Tbl1() As LidMisTbl
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Tbl2() As LidMisTbl
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Tbl3() As LidMisTbl
'QLib.Std.MDao_Lid_Mis_Samp.Private Function Tbl4() As LidMisTbl
'QLib.Cls.LidMisTyc.Friend Function Init(ExtNm, ActShtTy$, EptShtTyLis$) As LidMisTyc
'QLib.Cls.LidMisTyc.Property Get MisMsg$()
'QLib.Std.MDao_Li_Er_MsgzLiMis.Function MsgzLiMis(A As LiMis) As String()
'QLib.Std.MDao_Li_Er_MsgzLiMis.Private Function MisMsgTbl(A() As LiMisTbl) As String()
'QLib.Std.MDao_Li_Er_MsgzLiMis.Private Function MisMsgCol(A() As LiMisCol) As String()
'QLib.Std.MDao_Li_Er_MsgzLiMis.Private Function MisMsgTy(A() As LiMisTy) As String()
'QLib.Std.MDao_Li_Er_MsgzLiMis.Private Sub Z_ChkCol()
'QLib.Std.MDao_Ty_ShtTyDic.Function ShtTyDic(FxOrFb, TblNm) As Dictionary
'QLib.Std.MDao_Ty_ShtTyDic.Private Function ShtTyDiczFbt(Fb, T) As Dictionary
'QLib.Std.MDao_Ty_ShtTyDic.Private Function ShtTyDiczFxw(Fx, W) As Dictionary
'QLib.Std.AShpCst_Pm_LidPm.Property Get RptAppFb$()
'QLib.Std.AShpCst_Pm_LidPm.Private Sub Z_RptLidPrmSrc()
'QLib.Std.AShpCst_Pm_LidPm.Property Get RptLidPmSrc() As String()
'QLib.Std.AShpCst_Pm_LidPm.Private Function RptLidPmSrczAppFb(AppFb$) As String()
'QLib.Std.AShpCst_Pm_LidPm.Property Get RptLidPm() As LidPm
'QLib.Std.AShpCst_Pm_LidPm.Private Function RptLidPmzAppFb(AppFb$) As LidPm
'QLib.Std.AShpCst_Pm_LidPm.Private Function RptLidFilLinAy() As String()
'QLib.Std.AShpCst_Pm_LidPm.Private Function LidFilLinAy(AppFb$) As String()
'QLib.Std.AShpCst_Pm_LidPm.Private Function LidFilLin$(AppDb As Database, Itm$)
'QLib.Std.MDta_ReduceCol.Function ReduceCol(A As Drs) As ReduceCol
'QLib.Std.MDta_ReduceCol.Private Function FnyzReducibleCol(A As Drs) As String()
'QLib.Std.MDta_ReduceCol.Sub BrwReduceCol(A As ReduceCol)
'QLib.Std.MDta_ReduceCol.Private Function FmtReduceCol(A As ReduceCol) As String()
'QLib.Std.MDta_ReduceCol.Sub BrwDrszRedCol(A As Drs)
'QLib.Std.MIde_ConstMth__Fun.Function FtzConstQNm$(ConstQNm$)
'QLib.Std.MIde_ConstMth__Fun.Private Function ConstPrpPth$(MdNm$)
'QLib.Std.MIde_ConstMth__Fun.Function IsMthLinzConstStr(Lin) As Boolean
'QLib.Std.MIde_ConstMth__Fun.Function IsMthLinzConstLy(Lin) As Boolean
'QLib.Std.MIde_Mth_Lin_Is.Function IsMthLin(A) As Boolean
'QLib.Std.MIde_Mth_Lin_Is.Function IsMthLinzNm(Lin, Nm$) As Boolean
'QLib.Std.MIde_Src_Lin_Tak_VbStr.Function TakVbStr$(S)
'QLib.Std.MIde_Src_Lin_Tak_VbStr.Private Function EndPos%(Fm%, S, Lvl%)
'QLib.Std.MIde_Src_Lin_Tak_VbStr.Private Sub Z_TakVbStr()
'QLib.Std.MIde_VbCd_Expr.Function ExprLyzStr(Str, Optional MaxCdLinWdt% = 200) As String()
'QLib.Std.MIde_VbCd_Expr.Private Function ExprLyzLin(Lin, W%) As String()
'QLib.Std.MIde_VbCd_Expr.Private Function ShfLin(Str$, OvrFlwTerm$, W%) As LinRslt
'QLib.Std.MIde_VbCd_Expr.Private Function Z_ShfTermzPrintable()
'QLib.Std.MIde_VbCd_Expr.Private Function ShfTermzPrintable$(OStr$)
'QLib.Std.MIde_VbCd_Expr.Private Function LinRslt(ExprLin$, OvrFlwTerm$, S$) As LinRslt
'QLib.Std.MIde_VbCd_Expr.Private Function ExprzQuote$(BytAy() As Byte)
'QLib.Std.MIde_VbCd_Expr.Private Function ExprzAndChr$(BytAy() As Byte)
'QLib.Std.MIde_VbCd_Expr.Private Function Term(ExprTerm$, S$) As Term
'QLib.Std.MIde_VbCd_Expr.Private Sub Z_ExprLyzStr()
'QLib.Std.MIde_VbCd_Expr.Private Sub AAA()
'QLib.Std.MIde_VbCd_Expr.Private Sub Z_BrwRepeatedBytes()
'QLib.Std.MIde_VbCd_Expr.Function AscStr$(S)
'QLib.Std.MIde_VbCd_Expr.Private Sub Z_BrkAyzPrintable1()
'QLib.Std.MIde_VbCd_Expr.Function FmtPrintableStr$(T)
'QLib.Std.MIde_VbCd_Expr.Private Sub Z_BrkAyzPrintable()
'QLib.Std.MIde_VbCd_Expr.Private Function BrkAyzRepeat(S) As String()
'QLib.Std.MIde_VbCd_Expr.Private Function BrkAyzPrintable(S) As String()
'QLib.Std.MIde_VbCd_Expr.Private Function PrintableSts$(T)
'QLib.Std.MIde_VbCd_Expr.Private Function RepeatSts$(T)
'QLib.Std.MIde_VbCd_Expr.Private Function ShfTermzRepeatedOrNot$(OStr$)
'QLib.Std.MIde_VbCd_Expr.Private Sub BrwRepeatedBytes(S)
'QLib.Std.MIde_VbCd_Expr.Sub BBB()
'QLib.Std.MVb_Ay_Map_Align_Pm.Function FmtAyPm(Ay, PmStr$) As String() 'PmStr [FF..] [AlignNCol:FF..] ..
'QLib.Std.MVb_Ay_Map_Align_Pm.Private Function FmtAyPmzT1(Ay, T1, AlignNCol) As String()
'QLib.Std.MVb_Ay_Map_Align_Pm.Private Function T1ToAlignNColDic(PmStr$) As Dictionary
'QLib.Std.MVb_Ay_Map_Align_Pm.Private Function T1ToAlignNColDiczNoSrt(PmLy$()) As Dictionary
'QLib.Std.MVb_Fs_Ffn_Backup.Function FfnBackup$(Ffn)
'QLib.Std.MVb_Fs_Ffn_Backup.Function FfnRpl$(Ffn, ByFfn)
'QLib.Std.MVb_Zip.Sub ZipPth(Pth, Optional PthKd$ = "Path")
'QLib.Std.MDao_Bql_Read.Private Sub Z_CrtTTzPth()
'QLib.Std.MDao_Bql_Read.Sub CrtTTzPth(A As Database, FbqlPth)
'QLib.Std.MDao_Bql_Read.Sub CrtTTzPthTT(A As Database, FbqlPth, TT)
'QLib.Std.MDao_Bql_Read.Private Sub Z_CrtTblzFbql()
'QLib.Std.MDao_Bql_Read.Sub CrtFbzBqlPth(BqlPth, Optional Fb0$)
'QLib.Std.MDao_Bql_Read.Sub CrtTblzFbql(A As Database, T, Fbq)
'QLib.Std.MDao_Bql_Read.Sub CrtTblzShtTyBql(A As Database, T, ShtTyBql$)
'QLib.Std.MDao_Bql_Read.Private Function FdzShtTyscf(A) As Dao.Field
'QLib.Std.MDao_Bql_Read.Function ShtTyBqlzT$(A As Database, T)
'QLib.Std.MDao_Bql_Read.Private Function ShtTyszFd$(A As Dao.Field)
'QLib.Std.MDao_Bql_Write.Private Sub Z_WrtFbqlzDb()
'QLib.Std.MDao_Bql_Write.Private Sub Z_WrtFbqlzT()
'QLib.Std.MDao_Bql_Write.Sub WrtFbqlzDb(Pth, Db As Database)
'QLib.Std.MDao_Bql_Write.Sub WrtFbqlzTT(Pth, Db As Database, TT)
'QLib.Std.MDao_Bql_Write.Sub WrtFbql(Fbql, Db As Database, T)
'QLib.Std.MDao_Bql_Dr.Sub InsRszBql(R As Dao.Recordset, Bql)
'QLib.Std.MDao_Bql_Dr.Function BqlzRs$(A As Dao.Recordset)
'QLib.Std.MDao_Schm_Samp.Property Get SampSchm() As String()
'QLib.Std.MVb_Lin_Vy.Function ShfVy(OLin$, Lblss$) As Variant() ' 'Return Ay, which is '   Same sz as Lblss-cnt '   Ay-ele will be either string of boolean '   Each element is corresponding to terms-lblss 'Update OLin '   if the term match, it will removed from OLin 'Lblss: is in *LLL ?LLL or LLL '   *LLL is always first at beginning, mean the value OLin has not lbl '   ?LLL means the value is in OLin is using LLL => it is true, '   LLL  means the value in OLin is LLL=VVV 'OLin is '   VVV VVV=LLL [VVV=L L]
'QLib.Std.MVb_Lin_Vy.Private Function ShfTxtOpt(OAy$(), Lbl) As StrRslt
'QLib.Std.MVb_Lin_Vy.Private Function ShfBool(OAy$(), Lbl)
'QLib.Std.MVb_Lin_Vy.Private Function ShfTxt(OAy$(), Lbl)
'QLib.Std.MVb_Lin_Vy.Private Sub Z_ShfVy()
'QLib.Std.MIde_Gen_Push.Sub GenPushMd()
'QLib.Std.MIde_Gen_Push.Sub GenPushPj()
'QLib.Std.MIde_Gen_Push.Private Sub GenPushzMd(A As CodeModule)
'QLib.Std.MIde_Gen_Push.Sub MdEnsMth(A As CodeModule, MthDic As Dictionary)
'QLib.Std.MIde_Gen_Push.Private Function TyNyzGen(A As CodeModule) As String()
'QLib.Std.MIde_Gen_Push.Private Function TyNyzDlt(A As CodeModule) As String()
'QLib.Std.MIde_Gen_Push.Private Sub GenPushzPj(A As VBProject)
'QLib.Std.MIde_Gen_Push.Private Function MthDic(TyNyzGen$()) As Dictionary
'QLib.Std.MIde_Gen_Push.Private Function MthNyzDltTyNy(TyNyzDlt$()) As String()
'QLib.Std.MXls_TxtCn.Function TxtCnzWc(A As WorkbookConnection) As TextConnection
'QLib.Std.MDao_Lid_ErzLidPmzV1.Function ErzLidPmzV1(LidPm As LidPm) As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisFfn(MisFfn As Aset, Optional FilKind$ = "file") As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisFfnAy(FfnAy$(), Optional FilKind$ = "file") As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisTbl(T() As Tbl) As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisTbl1(A As Tbl) As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisCol(T() As Tbl) As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisCol1(T As Tbl) As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisCol2(T As Tbl) As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisTy(T() As Tbl) As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function MsgzMisTy1(T As Tbl) As String()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function T1Ay(A As LidPm) As Tbl()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function FfnDic(A() As LidFil) As Dictionary
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function T1Fx(A As LidFx, B() As LidFil, FfnDic As Dictionary) As Tbl
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function FsetzFxc(A() As LidFxc) As Aset
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function FldNmToEptShtTyLisDiczFxc(A() As LidFxc) As Dictionary
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function T1Fb(A As LidFb, B() As LidFil) As Tbl
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function T2Ay(T() As Tbl, ExistFfnAy$()) As Tbl()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function T3Ay(T() As Tbl) As Tbl()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function T4Ay(T() As Tbl) As Tbl()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function T5Ay(T() As Tbl) As Tbl()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function Tyc(ExistFset As Aset, FldNmToEptShtTyLisDic As Dictionary, Ffn$, TblNm$) As LidMisTyc()
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function Tyci(ActShtTy$, EptShtTyLis$, ExtNm) As TycOpt
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function FfnzFilNm$(FilNm$, A() As LidFil)
'QLib.Std.MDao_Lid_ErzLidPmzV1.Private Function FfnAyzLidFil(A() As LidFil) As String()
'QLib.Std.MSudoku.Private Function RRCC(R1 As Byte, R2 As Byte, C1 As Byte, C2 As Byte) As RRCC
'QLib.Std.MSudoku.Private Function SolveFstRound(Sq()) As Variant()
'QLib.Std.MSudoku.Private Function Solve(SudokuSq()) As Variant()
'QLib.Std.MSudoku.Private Function SolveRow(Sq()) As SolveRslt
'QLib.Std.MSudoku.Private Property Get NineEleOfRow(Sq(), Row%) As Variant()
'QLib.Std.MSudoku.Private Property Let NineEleOfRow(Sq(), Row%, NineEle())
'QLib.Std.MSudoku.Private Function SolveSmallSq(Sq()) As SolveRslt
'QLib.Std.MSudoku.Property Get NineEleOfSmallSq(Sq(), J%) As Variant()
'QLib.Std.MSudoku.Private Function RRCCzJ(J%) As RRCC
'QLib.Std.MSudoku.Property Let NineEleOfSmallSq(Sq(), J%, NineEle())
'QLib.Std.MSudoku.Private Function SolveCol(Sq()) As SolveRslt
'QLib.Std.MSudoku.Property Get NineEleOfCol(Sq(), Col%) As Variant()
'QLib.Std.MSudoku.Property Let NineEleOfCol(Sq(), Col%, NineEle())
'QLib.Std.MSudoku.Private Function SolveDiag(Sq()) As SolveRslt
'QLib.Std.MSudoku.Private Property Get NineEleOfDiag1(Sq()) As Variant()
'QLib.Std.MSudoku.Private Property Let NineEleOfDiag1(Sq(), NineEle())
'QLib.Std.MSudoku.Private Property Get NineEleOfDiag2(Sq()) As Variant()
'QLib.Std.MSudoku.Private Property Let NineEleOfDiag2(Sq(), NineEle())
'QLib.Std.MSudoku.Private Function SolveNineEleOfFstRnd(NineEle()) As Variant()
'QLib.Std.MSudoku.Private Function SolveNineEle(NineEle()) As NineEleRslt
'QLib.Std.MSudoku.Private Function Intersect(A() As Byte, B() As Byte)
'QLib.Std.MSudoku.Private Function ShouldBe(NineEle()) As Byte()
'QLib.Std.MSudoku.Sub SolveSudoku(Ws As Worksheet)
'QLib.Std.MSudoku.Private Function SudokuSq(Ws As Worksheet) As Variant()
'QLib.Std.MSudoku.Private Sub PutSudokuSolution(Ws As Worksheet, Sq())
'QLib.Std.MSudoku.Private Function SolutionRg(Ws As Worksheet) As Range
'QLib.Std.MSudoku.Private Property Get SampSudokuSq() As Variant()
'QLib.Std.MSudoku.Private Sub Z_PutSampSudoku()
'QLib.Std.MSudoku.Sub PutSampSudoku(At As Range)
'QLib.Std.MSudoku.Sub FmtSudoku(At As Range)
'QLib.Std.MSudoku.Private Sub Z_SolveSudoku()
'QLib.Std.MVb_RRCC.Function RRCC(R1, R2, C1, C2) As RRCC
'QLib.Std.MXls_Lo_Fmtr_Fmt.Sub BrwSampLof()
'QLib.Std.MXls_Lo_Fmtr_Fmt.Function FmtLof(Lof$()) As String()
'QLib.Std.MXls_Lo_Fmtr_Fmt.Function FmtSpec(Spec$(), Optional T1nn, Optional FmtFstNTerm% = 1) As String()
'QLib.Std.MXls_Vis.Function SetViszWb(A As Workbook, Vis As Boolean) As Workbook
'QLib.Std.MXls_Vis.Private Sub SetViszXls(A As Excel.Application, Vis As Boolean)
'QLib.Std.MXls_Vis.Function SetViszWs(A As Worksheet, Vis As Boolean) As Worksheet
'QLib.Std.MXls_Vis.Function WsVis(A As Worksheet) As Worksheet
'QLib.Std.MXls_Vis.Function WbVis(A As Workbook) As Workbook
'QLib.Std.MXls_Vis.Sub XlsVis(A As Excel.Application)
'QLib.Std.MXls_Vis.Function RgVis(Rg As Range) As Range
'QLib.Std.MXls_Vis.Function LoVis(A As ListObject) As ListObject
'QLib.Std.MIde_Gen_ErMsg.Sub GenErMsgzNm(MdNm$)
'QLib.Std.MIde_Gen_ErMsg.Sub GenErMsgMd()
'QLib.Std.MIde_Gen_ErMsg.Private Sub Init(Src$())
'QLib.Std.MIde_Gen_ErMsg.Private Sub Z_SrcGenErMsg()
'QLib.Std.MIde_Gen_ErMsg.Private Sub A_Prim()
'QLib.Std.MIde_Gen_ErMsg.Private Sub MdGenErMsg(Md As CodeModule)  'eMthNmTy.eeNve
'QLib.Std.MIde_Gen_ErMsg.Private Sub Z_MdGenErMsg()
'QLib.Std.MIde_Gen_ErMsg.Private Sub Z_ErConstDic()
'QLib.Std.MIde_Gen_ErMsg.Function ConstFTIxzMd(A As CodeModule, ConstNm$) As FTIx
'QLib.Std.MIde_Gen_ErMsg.Function ConstFTIx(DclLy$(), ConstNm) As FTIx
'QLib.Std.MIde_Gen_ErMsg.Function SrcGenErMsg(Src$(), Optional MdNm$ = "?") As String()
'QLib.Std.MIde_Gen_ErMsg.Function SrcRplConstDic(Src$(), ConstDic As Dictionary) As String()
'QLib.Std.MIde_Gen_ErMsg.Function DclRmvConstzSngLinConst(Dcl$(), ConstNmDic As Aset) As String() 'Assume: the const in Dcl to be remove is SngLin
'QLib.Std.MIde_Gen_ErMsg.Sub AsgDclAndBdy(Src$(), ODcl$(), OBdy$())
'QLib.Std.MIde_Gen_ErMsg.Function SrcRplMthDic(Src$(), MthDic As Dictionary) As String()
'QLib.Std.MIde_Gen_ErMsg.Function SrcRplMth(Src$(), MthNm, MthLines) As String()
'QLib.Std.MIde_Gen_ErMsg.Function MdRplMthDic(A As CodeModule, MthDic As Dictionary) As CodeModule
'QLib.Std.MIde_Gen_ErMsg.Function MdRplConstDic(A As CodeModule, ConstDic As Dictionary) As CodeModule
'QLib.Std.MIde_Gen_ErMsg.Function MdLines(StartLine, Lines, Optional InsLno0 = 0) As MdLines
'QLib.Std.MIde_Gen_ErMsg.Function EmpMdLines(A As CodeModule) As MdLines
'QLib.Std.MIde_Gen_ErMsg.Function MdLineszMdLno(A As CodeModule, Lno) As MdLines
'QLib.Std.MIde_Gen_ErMsg.Function MdLineszConst(A As CodeModule, ConstNm) As MdLines
'QLib.Std.MIde_Gen_ErMsg.Sub MdRplLines(A As CodeModule, B As MdLines, NewLines, Optional LinesNm$ = "MdLines")
'QLib.Std.MIde_Gen_ErMsg.Sub MdRplConst(A As CodeModule, ConstNm, NewLines)
'QLib.Std.MIde_Gen_ErMsg.Private Property Get ErMthNmSet() As Aset
'QLib.Std.MIde_Gen_ErMsg.Private Property Get ErMthNy() As String()
'QLib.Std.MIde_Gen_ErMsg.Private Property Get ErConstDic() As Dictionary
'QLib.Std.MIde_Gen_ErMsg.Private Function ErConstNm$(ErNm)
'QLib.Std.MIde_Gen_ErMsg.Private Sub Z_ErMthLinAy()
'QLib.Std.MIde_Gen_ErMsg.Private Sub Z_Init()
'QLib.Std.MIde_Gen_ErMsg.Private Function ZZ_Src() As String()
'QLib.Std.MIde_Gen_ErMsg.Private Function ErMthLinAy() As String() 'One ErMth is one-MulStmtLin
'QLib.Std.MIde_Gen_ErMsg.Private Function ErMthLinesByNm$(ErNm$, ErMsg$)
'QLib.Std.MIde_Gen_ErMsg.Private Function ErMthNm$(ErNm)
'QLib.Std.MIde_Dim.Function IsDimItmzAs(DimItm) As Boolean
'QLib.Std.MIde_Dim.Function DimNmzSht$(DimShtItm)
'QLib.Std.MIde_Dim.Function DimNmzAs$(DimAsItm)
'QLib.Std.MIde_Dim.Function DimTy$(DimItm)
'QLib.Std.MIde_Dim.Function DimNm$(DimItm)
'QLib.Std.MIde_Dim.Function IsDimItmzSht(DimItm) As Boolean
'QLib.Std.MIde_Dim.Function DimItmAy(Lin) As String()
'QLib.Std.MIde_Dim.Function DimNy(Lin) As String()
'QLib.Std.MIde_Dim.Function DimNyzDimItmAy(DimItmAy$()) As String()
'QLib.Std.MIde_Dim.Function DimNyzSrc(Src$()) As String()
'QLib.Cls.MdLines.Friend Function Init(StartLine, Lines, InsLno) As MdLines
'QLib.Cls.MdLines.Property Get Count&()
'QLib.Std.MVb_AscTbl.Function Chr99$()
'QLib.Std.MVb_AscTbl.Function AscWs() As Worksheet
'QLib.Std.MVb_AscTbl.Property Get AscSqOfNoNonPrt() As Variant()
'QLib.Std.MVb_AscTbl.Property Get AscSq() As Variant()
'QLib.Std.MVb_AscTbl.Function AscSqRplNonPrt(AscSq(), RplByAsc%) As Variant()
'QLib.Std.MVb_AscTbl.Sub BrwAsc()
'QLib.Std.MVb_AscTbl.Function IsVdtAsc(AscSq()) As Boolean
'QLib.Std.MVb_AscTbl.Property Get HexDigAy() As String()
'QLib.Std.MVb_AscTbl.Function AscSqAddLbl(AscSq()) As Variant()
'QLib.Std.MVb_AscTbl.Function FmtAsc(Optional RplNonPrtByAsc% = 8) As String()
'QLib.Std.MVb_AscTbl.Function FmtSq(Sq(), Optional SepChr$ = " ") As String()
'QLib.Std.MVb_AscTbl.Property Get FmtAscSq() As String()
'QLib.Std.MVb_AscTbl.Sub DmpAsc(S, Optional MaxLen& = 100)
'QLib.Std.MVb_AscTbl.Sub DmpAscSq()
'QLib.Std.MVb_AscTbl.Function RRCCzSq(Sq()) As RRCC
'QLib.Std.Module2.Sub FmtLozStdWb(A As Workbook)
'QLib.Std.Module2.Sub FmtLozStd(A As ListObject)
'QLib.Std.Module2.Property Get StdLof() As String()
'QLib.Std.MVb_Ay_Op_Rpl.Function AyRplFTIx(Ay, B As FTIx, ByAy)
'QLib.Std.MVb_Dic_CntDic.Function FmtCntDic(Ay, Optional IgnCas As Boolean, Optional Opt As eCntOpt) As String()
'QLib.Std.MIde_Fun_FmtMulLinSrc.Private Function DryzMulStmtSrc(MulStmtSrc$()) As Variant()
'QLib.Std.MIde_Fun_FmtMulLinSrc.Function FmtMulStmtSrc(MulStmtSrc$()) As String()
'QLib.Std.MVb_Rel.Property Get SampRel() As Rel
'QLib.Std.MVb_Rel.Property Get SampRelLy() As String()
'QLib.Std.MVb_Rel.Property Get SampMthRel() As Rel
'QLib.Std.MVb_Rel.Property Get SampMthRelLy() As String()
'QLib.Std.MIde_MthId.Function FmtMthQidLyOfVbe() As String()
'QLib.Std.MIde_MthId.Private Sub Z_FmtMthQidLyOfVbe()
'QLib.Std.MIde_MthId.Function MthRetNmRstLinzMthNmRstLin$(MthNmRstLin$, IsRetVal As Boolean)
'QLib.Std.MIde_MthId.Function MthQidLyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_MthId.Private Sub Z_MthQidLyOfVbe()
'QLib.Std.MIde_MthId.Function MthSrtKey$(ShtMthMdy$, MthNm$)
'QLib.Std.MIde_MthId.Private Function DicOf_PjMdTyMdNm_To_MthQLy() As Dictionary
'QLib.Std.MIde_MthId.Private Function MthSrtKeyzLin$(MthLin) ' MthKey is Mdy.Nm
'QLib.Std.MIde_MthId.Private Function MthQidLy(MthQLy$()) As String()
'QLib.Std.MIde_MthId.Function DotLinRmvSegN$(DotLin, Optional SegN% = 1)
'QLib.Std.MIde_MthId.Function FstNDotSeg$(DotLin$, Optional NSeg% = 1)
'QLib.Std.MIde_MthId.Function DotLyInsSep(DotLy$(), Optional UpToNSeg% = 1, Optional Sfx$ = "------") As String()
'QLib.Std.MIde_MthId.Function DotLyRmvSegN(DotLy$(), Optional SegN% = 1) As String()
'QLib.Std.MIde_MthId.Private Function MthQMLy(MthQLy$()) As String()
'QLib.Std.MIde_MthId.Private Sub Z_MthSQMLin()
'QLib.Std.MIde_MthId.Private Function MthSQMLin$(MthQLin)
'QLib.Std.MIde_MthId.Private Sub Asg_ShtMthMdy_ShtMthTy_MthNm_MthNmRst(OShtMthMdy$, OShtMthTy$, OMthNm$, OMthNmRst$, MthLin$)
'QLib.Std.MIde_MthId.Private Function MthQidLin$(MthQMLin, Id$)
'QLib.Std.MIde_MthId.Private Function MthQidLyzSamMdMthQLy(SamMdMthQLy$()) As String() 'Assume the MthQLy are from same module
'QLib.Std.MIde_Fun_VerbPatn.Property Get BRKCmlASet() As Aset
'QLib.Std.MIde_Fun_VerbPatn.Property Get MthVNyOfVbe() As String()
'QLib.Std.MIde_Fun_VerbPatn.Sub Z_MthVNsetOfVbe()
'QLib.Std.MIde_Fun_VerbPatn.Property Get MthDNmToMdDNmRelOfVbe() As Rel
'QLib.Std.MIde_Fun_VerbPatn.Private Function MthDNmToMdDNmRelzVbe(A As Vbe) As Rel
'QLib.Std.MIde_Fun_VerbPatn.Property Get MthVNsetOfVbe() As Aset
'QLib.Std.MIde_Fun_VerbPatn.Function MthQVNsetOfVbe(Optional WhStr$) As Aset
'QLib.Std.MIde_Fun_VerbPatn.Sub VcMthQVNsetOfVbe(Optional WhStr$)
'QLib.Std.MIde_Fun_VerbPatn.Sub VcMthQVNyOfVbe(Optional WhStr$)
'QLib.Std.MIde_Fun_VerbPatn.Function MthQVNyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Fun_VerbPatn.Function MthQVNyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Fun_VerbPatn.Function QVNy(Ny$()) As String()
'QLib.Std.MIde_Fun_VerbPatn.Function QBNm$(Nm)
'QLib.Std.MIde_Fun_VerbPatn.Function QVBNm$(Nm) 'Quote-Verb-and-cmlBrk-Nm.
'QLib.Std.MIde_Fun_VerbPatn.Function QVNm$(Nm)
'QLib.Std.MIde_Fun_VerbPatn.Function MthVNm$(MthNm)
'QLib.Std.MIde_Fun_VerbPatn.Property Get VerbRe() As RegExp
'QLib.Std.MIde_Fun_VerbPatn.Sub BrwVerb()
'QLib.Std.MIde_Fun_VerbPatn.Sub VcNVTDNmAsetOfVbe()
'QLib.Std.MIde_Fun_VerbPatn.Property Get NVTDNmAsetOfVbe() As Aset
'QLib.Std.MIde_Fun_VerbPatn.Property Get NVTDNyOfVbe() As String()
'QLib.Std.MIde_Fun_VerbPatn.Private Function NVTDNyzVbe(A As Vbe) As String()
'QLib.Std.MIde_Fun_VerbPatn.Private Function NVTDNy(Ny$()) As String()
'QLib.Std.MIde_Fun_VerbPatn.Private Function NVTDNm$(Nm) 'Nm.Verb.Ty.Dot-Nm
'QLib.Std.MIde_Fun_VerbPatn.Function FstVerbSubNyOfVbe() As String()
'QLib.Std.MIde_Fun_VerbPatn.Function NVTy$(Nm) 'Nm.Verb-Ty
'QLib.Std.MIde_Fun_VerbPatn.Function IsNoVerbNm(Nm) As Boolean
'QLib.Std.MIde_Fun_VerbPatn.Function IsMidVerbNm(Nm) As Boolean
'QLib.Std.MIde_Fun_VerbPatn.Function IsFstVerbNm(Nm) As Boolean
'QLib.Std.MIde_Fun_VerbPatn.Function IsVerb(S) As Boolean
'QLib.Std.MIde_Fun_VerbPatn.Property Get VerbAset() As Aset
'QLib.Std.MIde_Fun_VerbPatn.Function RmvEndDig$(S)
'QLib.Std.MIde_Fun_VerbPatn.Function Verb$(Nm)
'QLib.Std.MIde_Fun_VerbPatn.Property Get NormVerbss$()
'QLib.Std.MIde_Fun_VerbPatn.Function NormSsl$(Ssl, Optional IsDes As Boolean)
'QLib.Std.MIde_Fun_VerbPatn.Function PatnzVerbss$(Verbss$)
'QLib.Std.MIde_Fun_VerbPatn.Private Function PatnzVerb$(Verb)
'QLib.Std.MIde_Fun_VerbPatn.Private Sub ThwIfNotVerb(S, Fun$)
'QLib.Std.MIde_Fun_VerbPatn.Function QuoteVerb$(Nm)
'QLib.Std.AA.Sub XXX1A()
'QLib.Std.AA.Private Sub XXX1B()
'QLib.Std.AA.Private Function VerbPatn$(XX)
'QLib.Std.MIde_Mth_Lin_Brk.Private Function ShfRetTyzAftPm$(OAftPm$)
'QLib.Std.MIde_Mth_Lin_Brk.Private Function RmkzAftRetTy$(AftRetTy$)
'QLib.Std.MIde_Mth_Lin_Brk.Function MthLinRec(MthLin) As MthLinRec
'QLib.Std.MIde_Mth_Lin_Brk.Function MthFLin$(MthQLin)
'QLib.Std.MIde_Mth_Lin_Brk.Function MthFLyOfVbe(Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Lin_Brk.Function MthFLyzVbe(A As Vbe, Optional WhStr$) As String()
'QLib.Std.MIde_Mth_Lin_Brk.Function MthFLy(MthQLy$()) As String()
'QLib.Std.MIde_Mth_Lin_Brk.Function MthFLinzMthLin$(MthLin)
'QLib.Std.MIde_Mth_Lin_Brk.Function FmtPm(Pm$, Optional IsNoBkt As Boolean) 'Pm is wo bkt.
'QLib.Std.MIde_Mth_Lin_Brk.Function ShtRetTyAsetOfVbe(Optional WhStr$) As Aset
'QLib.Std.MIde_Mth_Lin_Brk.Function ShtRetTyAsetzVbe(A As Vbe, Optional WhStr$) As Aset
'QLib.Std.MIde_Mth_Lin_Brk.Function ShtRetTyAset(MthLinAy$()) As Aset
'QLib.Std.MIde_Mth_Lin_Brk.Function ShtRetTyAy(MthLinAy$()) As String()
'QLib.Std.MIde_Mth_Lin_Brk.Function ShtRetTyzLin$(MthLin)
'QLib.Std.MIde_Mth_Lin_Brk.Function ShtRetTy$(TyChr$, RetTy$, IsRetVal As Boolean, Optional ExlColon As Boolean)
'QLib.Std.MIde_Ens_PrpEr1.Private Sub EnsLinzExit(OMthLy$())
'QLib.Std.MIde_Ens_PrpEr1.Private Sub EnsLinLblX(OMthLy$())
'QLib.Std.MIde_Ens_PrpEr1.Private Sub EnsLinzOnEr(OMthLy$())
'QLib.Std.MIde_Ens_PrpEr1.Private Function IxOfExit&(MthLy$())
'QLib.Std.MIde_Ens_PrpEr1.Private Function IxOfInsExit&(MthLy$())
'QLib.Std.MIde_Ens_PrpEr1.Private Function LinzLblX$()
'QLib.Std.MIde_Ens_PrpEr1.Private Function IxOfLblX&(MthLy$())
'QLib.Std.MIde_Ens_PrpEr1.Private Function IxOfOnEr&(MthLy$())
'QLib.Std.MIde_Ens_PrpEr1.Private Sub Z_SrcEnsPrpOnEr()
'QLib.Std.MIde_Ens_PrpEr1.Private Function SrcEnsPrpOnEr(Src$()) As String()
'QLib.Std.MIde_Ens_PrpEr1.Function SyPairOfTopRmkOooMthLy(MthLyWiTopRmk$()) As SyPair
'QLib.Std.MIde_Ens_PrpEr1.Private Function MthLyEnsPrpOnEr(MthLyWiTopRmk$()) As String()
'QLib.Std.MIde_Ens_PrpEr1.Private Function MthLyRmvPrpOnEr(MthLy$()) As String()
'QLib.Std.MIde_Ens_PrpEr1.Private Function RmvPrpOnErzSrc(Src$()) As String()
'QLib.Std.MIde_Ens_PrpEr1.Private Sub RmvPrpOnErzMd(A As CodeModule)
'QLib.Std.MIde_Ens_PrpEr1.Sub RmvPrpOnErOfMd()
'QLib.Std.MIde_Ens_PrpEr1.Sub EnsPrpOnErzMd(A As CodeModule)
'QLib.Std.MIde_Ens_PrpEr1.Sub EnsPrpOnErOfMd1()
'QLib.Std.MIde_Ens_PrpEr1.Private Sub Z_EnsPrpOnErzMd()
'QLib.Cls.Pos.Friend Function Init(Cno1, Cno2) As Pos
'QLib.Cls.LinPos.Friend Function Init(Lno, Pos As Pos) As LinPos
'QLib.Cls.MdPos.Friend Function Init(Md As CodeModule, Pos As LinPos) As MdPos

