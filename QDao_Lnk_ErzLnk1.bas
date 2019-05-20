Attribute VB_Name = "QDao_Lnk_ErzLnk1"
Option Compare Text
Option Explicit
Dim D As Database:
Private Const M_Stru_DupFld$ = "Lin#[?] have dup fld: Stru[?] Fld[?]."
Private Const M_Stru_DupStru$ = "Lin#[?] are all have Stru[?]"
Private Const M_Stru_ErFldTy$ = "Lin#[?] has invalid FldTy[?].  See VdtFldTy."

Private Const M_Stru_ErFld_VdtFldTy$ = "VdtFldTy: ...."
Private Const M_Stru_ExcessStru$ = "Lin#[?] is exccess stru."
Private Const M_Stru_ExcessStru__StruInUse$ = "   These are stru in use: [?]"
Private Const M_Stru_MisgExtn$ = "Lin#[?] Stru[?] Extn[?] is not found in Fxn[?] Wsn[?]."
Private Const M_Stru_MisFldTy$ = "Lin#[?] Stru[?] Extn[?] has Ty[?] which is not as expected Ty[?]"
Private Const M_Stru_NoStru$ = "There is no Stru.XXX sections"

Private Const M_Stru_NoFld$ = "Lin#[?] Stru[?] has no field"
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

Private Sub B_TblWh_MisTn()
End Sub

Private Property Get C_CMSrc() As String()
Erase XX
X "Stru_NoFld"
X "Inp_DupFbxn"

End Property

Private Sub B_Stru__IsNoStru()

End Sub
Private Sub B_Stru__NoFldStruAy()

End Sub
Private Sub B_Stru__DupFld()

End Sub

Private Sub B_Stru__DupFlds()
End Sub
Private Sub B_Ins__Lines(Ly$(), T$)
Dim L, J%, Q$
For Each L In Ly
    J = J + 1
    If Not HasPfx(LTrim(L), "--") Then
        Q = FmtQQ("Insert into [?] (Lno,L) Values(?,""?"")", T, J, RTrim(L)): D.Execute Q
    End If
Next
End Sub
Private Sub B_Brk_ImpLnkSrc() ' Add Colun IsHdr K
'--
D.Execute "Alter Table [#LnkImpSrc] Add Column IsHdr YesNo, K Text(50)"
D.Execute "Update [#LnkImpSrc] set IsHdr=True where Left(L,1)<>''"
D.Execute "Update [#LnkImpSrc] set K=L"
D.Execute "Update [#LnkImpSrc] set K=Left(L,Instr(K,' ')-1) where Instr(L,' ')>0 and IsHdr"
D.Execute "Update [#LnkImpSrc] set L=Trim(L)                where Not IsHdr"
With D.OpenRecordset("Select IsHdr,K from [#LnkImpSrc]")
    Dim K$
    While Not .EOF
        If !IsHdr Then
            K = !K
        Else
            .Edit
            !K = K
            .Update
        End If
        .MoveNext
    Wend
    .Close
End With

End Sub

Function ErzLnk1(InpFilSrc$(), LnkImpSrc$()) As String()
ThwIf_KFsEr KFs(InpFilSrc), CSub
Set D = TmpDb("Lnk", "Lnk")
D.Execute "Create Table [#InpFilSrc] (Er Memo, Lno Integer,L Text(255))"
D.Execute "Create Table [#LnkImpSrc] (Er Memo, Lno Integer,L Text(255))"
B_Ins__Lines InpFilSrc, "#InpFilsrc"
B_Ins__Lines LnkImpSrc, "#LnkImpsrc"
B_Brk_ImpLnkSrc
'-- [#InpFilSrc]
    D.Execute "Alter Table [#InpFilSrc] Add Column SpcPos Integer, Tn Text(20), Ffn Text(235)"
    D.Execute "Update [#InpFilSrc] set SpcPos = Instr(L,' ')"
    D.Execute "Update [#InpFilSrc] set Tn = Trim(Left(L,SpcPos-1))"
    D.Execute "Update [#InpFilSrc] set Ffn = Trim(Mid(L,SpcPos+1))"
'-- [TblWh]
    D.Execute "Create Table [TblWh] (Er Memo, Lno Integer, L Text(255), SpcPos Integer, Tn Text(20),Bexpr Text(234))"
    D.Execute "Insert into [TblWh] Select Lno,Trim(x.L) as L from [#LnkImpSrc] x where K='Tbl.Where' and Not IsHdr"
    D.Execute "Update [TblWh] set SpcPos = Instr(L,' ')"
    D.Execute "Update [TblWh] set Tn = Trim(Left(L,SpcPos-1))"
    D.Execute "Update [TblWh] set Bexpr = Trim(Mid(L,SpcPos+1))"
'-- [FxTbl]
    D.Execute "Create Table [FxTbl] (Er Memo, Lno Integer, L Text(255), W Text(255), P Integer, Tn Text(20),Fxn Text(20), Wsn Text(20),Stru Text(30))"
    D.Execute "Insert into [FxTbl] Select Lno,Trim(x.L) as L from [#LnkImpSrc] x where K='FxTbl' and Not IsHdr"
    D.Execute "Update [FxTbl] Set W = L"
    
    D.Execute "Update [FxTbl] set P = Instr(W,' ')"
    D.Execute "Update [FxTbl] set Tn = Trim(Left(W,P-1)) where P>0"
    D.Execute "Update [FxTbl] set Tn = Trim(W)           where P=0"
    D.Execute "Update [FxTbl] set W = Trim(Mid(W,P+1))   where P>0"
    D.Execute "Update [FxTbl] set W = ''                 where P=0"
    
    D.Execute "Update [FxTbl] set P = Instr(W,'.')"
    D.Execute "Update [FxTbl] set Fxn = Trim(Left(W,P-1)) where P>0"
    D.Execute "Update [FxTbl] set Fxn = Tn                where P=0"
    D.Execute "Update [FxTbl] set W = Trim(Mid(W,P+1))    where P>0"
    D.Execute "Update [FxTbl] set W = Trim(W)             where P=0"
    
    D.Execute "Update [FxTbl] set P = Instr(W,' ')"
    D.Execute "Update [FxTbl] set Wsn = Trim(Left(W,P-1)) where P>0"
    D.Execute "Update [FxTbl] set Wsn = 'Sheet1'          where P=0"
    D.Execute "Update [FxTbl] set Stru = Trim(Mid(W,P+1)) where P>0"
    D.Execute "Update [FxTbl] set Stru = Tn               where Stru is null"
    D.Execute "Alter Table [FxTbl] drop column W,P"
'-- [FxTbl]
    D.Execute "Create Table [FbTbl] (Er Memo, Lno Integer, Fbn Text(20),Tn Text(20))"
    Dim Lno&, Fbn_Tnss$, Rs As DAO.Recordset
    Set Rs = D.TableDefs("FbTbl").OpenRecordset
    With D.OpenRecordset("Select Lno,L from [#LnkImpSrc] where Not IsHdr and K = 'FbTbl'")
        While Not .EOF
            Lno = !Lno
            Fbn_Tnss = Trim(!L)
            B_Ins_FbTbl Rs, Lno, Fbn_Tnss
            .MoveNext
        Wend
        .Close
    End With
    Rs.Close
'-- [Stru]
    D.Execute "Create Table [Stru] (Er Memo, L Text(255), Lno Integer, Stru Text(30))"
    D.Execute "Insert into [Stru] Select Lno,L from [#LnkImpSrc] where Left(K,5)='Stru.' and IsHdr"
    D.Execute "Update [Stru] set Stru = Trim(Mid(L,6)) where Left(L,5)='Stru.'"

'-- [StruF]
    D.Execute "Create Table [StruF] (Er Memo, K Text(50), L Text(255), W Text(255), P Integer, Lno Integer, Stru Text(30), Fld Text(30),Ty Text(20), Extn Text(30))"
    D.Execute "Insert into [StruF] Select Lno,K,L from [#LnkImpSrc] where Left(K,5)='Stru.' and Not IsHdr"
    D.Execute "Update [StruF] set W = L"
    D.Execute "Update [StruF] set Stru = Trim(Mid(K,6))"
    D.Execute "Update [StruF] set P = Instr(W,' ')"
    D.Execute "Update [StruF] set Fld = Left(W,P-1) where P>0"
    D.Execute "Update [StruF] set Fld = W           where P=0"
    D.Execute "Update [StruF] set W = Trim(Mid(W,P+1)) where P>0"
    D.Execute "Update [StruF] set W = ''               where P=0"
    
    D.Execute "Update [StruF] set P = Instr(W,' ')"
    D.Execute "Update [StruF] set Ty = Left(W,P-1) where P>0"
    D.Execute "Update [StruF] set Ty = W           where P=0"
    D.Execute "Update [StruF] set W = Trim(Mid(W,P+1))"
    
    D.Execute "Update [StruF] set Extn = W"
    D.Execute "Update [StruF] set Extn = Mid(W,2,Len(W)-2) where Left(W,1)='[' and Right(W,1)=']'"
    D.Execute "Alter Table [StruF] drop column W,P"
    
    BrwDb D
    Stop
'-- @Er
    D.Execute "Create Table [@Er] (ErGp Text(10),Ern Text(20),Lnoss Text(30),Msg Text(100))"
    
'-- NoStru
If Not HasReczT(D, "Stru") Then
    D.Execute "Insert Into [@Er] Values('Stru','NoStru','','No Stru.XXX')"
End If

'-- DupStru
    D.Execute "Select Distinct Stru into [#A] from [Stru] group by Stru having Count(*)>1"
    D.Execute "Alter Table [#A] Add Column Lnoss Text(200)"
    With RszT(D, "#A")
        While Not .EOF
            Dim Lnoss$: Lnoss = JnSpc(LngAyzQ(D, FmtQQ("Select Lno from [Stru] where Stru = '?'", !Stru)))
            D.Execute "Insert into [@Er] (ErGp,Ern,Lnoss,Msg) Values ('Stru','DupStru','?','Stru[?] are duplicated')"
            .MoveNext
        Wend
    End With
    
    BrwDb D
    Stop

B_Stru_DupFld
B_Stru_ErFldTy
B_Stru_ExcessStru
B_Stru_MisgExtn
B_Stru_MisMchFxFldTy
B_Stru_MisgFxFldTy
B_Stru_NoFld
    
B_FbTbl_DupFbt
B_FbTbl_DupFbn
    
B_FxTbl_DupFxt
B_FxTbl_MisFxn
B_FxTbl_MisStru
B_FxTbl_MisWsn
    
B_TblWh_MisTn
B_TblWh_DupTbl
End Function
Private Sub B_Ins_FbTbl(Rs As DAO.Recordset, Lno&, Fbn_Tnss$)
With Rs
    Dim L$:     L = Fbn_Tnss
    Dim Fbn$: Fbn = ShfT1(L)
    Dim T
    For Each T In SyzSS(L)
        .AddNew
        !Lno = Lno
        !Fbn = Fbn
        !Tn = T
        .Update
    Next
End With
End Sub
            
Private Sub B_FxTbl_MisWsn()

End Sub

Private Sub B_FxTbl_MisStru()

End Sub

Private Sub B_FbTbl_DupFbn()

End Sub

Private Sub B_FbTbl_DupFbt()

End Sub

Private Sub B_StruSyzNoFld()
End Sub

Private Sub B_Stru_NoFld()
End Sub

Private Sub B_Stru_DupFld()
End Sub

Private Sub B_TblWh_DupTbl()
End Sub

Private Sub B_FxTbl_DupFxt()
End Sub

Private Sub B_FxTbl_MisFxn()
End Sub

Private Sub B_Stru_DupStru()
End Sub

Private Sub B_Stru_ErFldTy()
End Sub

Private Sub B_Stru_ExcessStru()
End Sub

Private Sub B_Stru_MisgExtn()
End Sub

Private Sub B_Stru_MisgFxFldTy()
End Sub

Private Sub B_Stru_MisMchFxFldTy()
End Sub

Private Sub B_LnoAyzStru()
End Sub

Sub Z3()
ZZ_ErzLnk
End Sub
Private Sub ZZ_ErzLnk()
B ErzLnk(Y_InpFilSrc, Y_LnkImpSrc)
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



