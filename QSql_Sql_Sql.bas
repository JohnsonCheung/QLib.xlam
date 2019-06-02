Attribute VB_Name = "QSql_Sql_Sql"
Option Explicit
Option Compare Text
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Sql."
Const KwBet$ = "between"
Const KwDist$ = "distinct"
Const KwUpd$ = "update"
Const KwInto$ = "into"
Const KwSel$ = "select"
Const KwSelDis$ = KwSel & " " & KwDist
Const KwFm$ = "from"
Const KwGp$ = "group by"
Const KwWh$ = "where"
Const KwAnd$ = "and"
Const KwJn$ = "join"
Const KwOr$ = "or"
Const KwOrd$ = "order by"
Const KwLeftJn$ = "left join"

Private Function PSel_Fny_Extny_NOFMT(Fny$(), Extny$(), Optional IsDis As Boolean)
Dim O$(), J%, E$, F$
For J = 0 To UB(Fny)
    F = Fny(J)
    E = Trim(Extny(J))
    Select Case True
    Case E = "", E = F: PushI O, F
    Case Else: PushI O, QuoteSq(E) & " As " & F
    End Select
Next
PSel_Fny_Extny_NOFMT = KwSelzIsDis(IsDis) & " " & JnCommaSpc(O)
End Function

Private Property Get C_CNL$()
C_CNL = "," & vbCrLf  'Comma-NewLin-Tab
End Property

Private Property Get C_CNLT$()
C_CNLT = "," & vbCrLf & C_T  'Comma-NewLin-Tab
End Property
Private Property Get C_Fm$()
C_Fm = C_NLT & KwFm & C_T
End Property

Private Property Get C_Into$()
C_Into = C_NLT & KwInto & C_T
End Property

Private Property Get C_NL$() ' New Line
If Fmt Then
    C_NL = vbCrLf
Else
    C_NL = " "
End If
End Property

Private Property Get C_T$()
If Fmt Then
    C_T = "    "
Else
    C_T = " "
End If
End Property
Private Property Get C_NLT$() ' New Line Tabe
If Fmt Then
    C_NLT = C_NL & C_T
Else
    C_NLT = " "
End If
End Property

Private Property Get C_NLTT$() ' New Line Tabe
If Fmt Then
    C_NLTT = C_NLT & C_T
Else
    C_NLTT = " "
End If
End Property

Private Property Get C_And$()
If Fmt Then
    C_And = C_NLT & KwAnd & C_T
Else
    C_And = " " & KwAnd & " "
End If
End Property

Private Property Get C_Wh$()
C_Wh = C_NLT & KwWh & C_NLT
End Property

Private Property Get C_CommaSpc$()
If Fmt Then
    C_CommaSpc = C_CNLT
Else
    C_CommaSpc = ", "
End If
End Property


Private Function PBkt_FF$(FF$)
PBkt_FF = QuoteBkt(SyzSS(FF))
End Function

Private Function PBkt_Av$(Av())
Dim O$(), I
For Each I In Av
    PushI O, SqlQuote(I)
Next
PBkt_Av = QuoteBktJnComma(Av)
End Function


Private Function PFldInX_F_InAset_Wdt(F, S As Aset, Wdt%) As String()
Dim A$
    A = "[F] in ("
Dim I
'For Each I In LyJnQSqlCommaAsetW(S, Wdt - Len(A))
    PushI PFldInX_F_InAset_Wdt, I
'Next
End Function

Private Function PFm_T$(T)
PFm_T = C_Fm & QuoteSq(T)
End Function

Private Function PFm_X$(X)
PFm_X = C_Fm & X
End Function


Private Function PGp_ExprVblAy$(ExprVblAy$())
PGp_ExprVblAy = "|  Group By " & JnCrLf(FmtExprVblAy(ExprVblAy))
End Function


Private Function PIns_T$(T)
PIns_T = "Insert into [" & T & "]"
End Function



Private Function PInto_T$(T)
PInto_T = C_Into & "[" & T & "]"
End Function


Private Function POrd_MinusSfxFF$(OrdMinusSfxFF$)
If OrdMinusSfxFF = "" Then Exit Function
Dim O$(): O = SyzSS(OrdMinusSfxFF)
Dim I, J%
For Each I In O
    If HasSfx(O(J), "-") Then
        O(J) = RmvSfx(O(J), "-") & " desc"
    End If
    J = J + 1
Next
POrd_MinusSfxFF = C_NLT & "order by " & JnCommaSpc(O)
End Function

Private Function PSel_F$(F$)
PSel_F = "Select [" & F & "]"
End Function


Private Function PSel_FF$(FF, Optional Dis As Boolean)
PSel_FF = PSel_Fny(SyzSS(FF), Dis)
End Function

Private Function PSel_FF_Extny$(FF$, Extny$())
PSel_FF_Extny = PSel_X(PSel_Fny_Extny(Ny(FF), Extny))
End Function

Private Function PSel_Fny_Extny$(Fny$(), Extny$(), Optional IsDis As Boolean)
If Not Fmt Then PSel_Fny_Extny = PSel_Fny_Extny_NOFMT(Fny, Extny): Exit Function
Dim E$(), F$()
F = Fny
E = Extny
FEs_SetExtNm_ToBlank_IfEqToFld F, E
FEs_SqQuoteExtNm_IfNonBlank E
FEs_AlignExtNm E
FEs_AddAs_Or4Spc_ToExtNm E
FEs_AddTab2Spc_ToExtNm E
FEs_AlignFld F
PSel_Fny_Extny = KwSelzIsDis(IsDis) & C_NL & Join(JnAyab(E, F), C_CNL)
End Function

Private Sub FEs_AddTab2Spc_ToExtNm(OE$())
OE = AddPfxzAy(OE, C_T & "  ")
End Sub
Private Sub FEs_SetExtNm_ToBlank_IfEqToFld(F$(), OE$())
Dim J%
For J = 0 To UB(OE)
    If OE(J) = F(J) Then OE(J) = ""
Next
End Sub
Private Sub FEs_SqQuoteExtNm_IfNonBlank(OE$())
Dim J%
For J = 0 To UB(OE)
    If OE(J) <> "" Then
        OE(J) = QuoteSq(OE(J))
    End If
Next
End Sub
Private Sub FEs_AlignExtNm(OE$())
OE = AlignLzAy(OE)
End Sub
Private Sub FEs_AddAs_Or4Spc_ToExtNm(OE$())
Dim J%, C$
For J = 0 To UB(OE)
    
    If Trim(OE(J)) = "" Then
        C = "    "
    Else
        C = " As "
    End If
    OE(J) = OE(J) & C
Next
End Sub
Private Sub FEs_AlignFld(OF$())
OF = AlignLzAy(OF)
End Sub

Private Function PSel_Fny$(Fny$(), Optional Dis As Boolean)
'PSel_FF = PSel_Dis(Dis) & C_NLTT & JnCommaSpc(Fny)
Stop
End Function

Private Sub ZZ_PSel_Fny_Extny()
Dim Fny$()
Dim Extny$()
GoSub ZZ
Exit Sub
ZZ:
    Fny = SyzSS("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Extny = TermAy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Debug.Print PSel_Fny_Extny(Fny, Extny)
    Return
End Sub

Private Function PSel_T$(T)
PSel_T = KwSel & C_T & "*" & PFm_T(T)
End Function


Private Function PSel_X$(X$, Optional Dis As Boolean)
PSel_X = KwSelzIsDis(Dis) & X
End Function

Private Function PSet_FF_EqDr$(FF$, EqDr)

End Function


Private Function PSet_FF_Evy$(FF$, EqVy)

End Function

Private Property Get Fmt() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then X = True: Y = Cfg.Sql.FmtSql
Fmt = Y
End Property

Private Function PExpr_T_RecId$(T, RecId)
PExpr_T_RecId = FmtQQ("?Id=?", T, RecId)
End Function

Private Function PSet_Fny_Vy$(Fny$(), Vy())
Dim F$(): F = QuoteSqzAy(Fny)
Dim V$(): V = SqlQuoteVy(Vy)
PSet_Fny_Vy = JnComma(JnAyab(F, V, "="))
End Function

Private Property Get C_Comma$()
If Fmt Then
    C_Comma = "," & vbCrLf
Else
    C_Comma = ", "
End If
End Property

Private Function PAnd_Bexpr$(Bexpr$)
If Bexpr = "" Then Exit Function
'PAnd_Bexpr = NxtLin & "and " & NxtLin_Tab & Bexpr
End Function

Private Function AddPfxNLTT$(Sy$())
AddPfxNLTT = Jn(AddPfxzAy(Sy, C_NLTT), "")
End Function

Private Function Bexpr_E_InLis$(Expr$, InLisStr$)
If InLisStr = "" Then Exit Function
Bexpr_E_InLis = FmtQQ("? in (?)", Expr, InLisStr)
End Function

Private Sub Z_PGp_ExprVblAy()
Dim ExprVblAy$()
    Push ExprVblAy, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push ExprVblAy, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push ExprVblAy, "3sldkfjsdf"
DmpAy SplitVBar(PGp_ExprVblAy(ExprVblAy))
End Sub


Private Function KwSelzIsDis$(IsDis As Boolean)
If IsDis Then
    KwSelzIsDis = KwSelDis
Else
    KwSelzIsDis = KwSel
End If
End Function


Function JnCommaSpcFF$(FF$)
JnCommaSpcFF = JnCommaSpc(SyzSS(FF))
End Function



Private Sub Z_PSel()
Dim Fny$(), ExprVblAy$()
ExprVblAy = Sy("F1-Expr", "F2-Expr   AA|BB    X|DD       Y", "F3-Expr  x")
Fny = SplitSpc("F1 F2 F3xxxxx")
'Debug.Print LineszVbl(PSelFFFldLvs(Fny, ExprVblAy))
End Sub

Function SqlSel_FF_T$(FF, T, Optional IsDis As Boolean)
SqlSel_FF_T = PSel_FF(FF, IsDis) & PFm_T(T)
End Function

Private Function PSet_FF_ExprAy$(FF, Ey$())
Const CSub$ = CMod & "PSet_FF_Ey"
Dim Fny$(): Fny = SyzSS(FF)
Ass IsVblAy(Ey)
If Si(Fny) <> Si(Ey) Then Thw CSub, "[FF-Sz} <> [Si-Ey], where [FF],[Ey]", Si(Fny), Si(Ey), FF, Ey
Dim AFny$()
    AFny = AlignLzAy(Fny)
    AFny = AddSfxzAy(AFny, " = ")
Dim W%
    'W = VblWdtAy(Ey)
Dim Ident%
    W = WdtzAy(AFny)
Dim Ay$()
    Dim J%, U%, S$
    U = UB(AFny)
    For J = 0 To U
        If J = U Then
            S = ""
        Else
            S = ","
        End If
        'Push Ay, VblAlign(Ey(J), Pfx:=AFny(J), IdentOpt:=Ident, WdtOpt:=W, Sfx:=S)
    Next
Dim Vbl$
    Dim Ay1$()
    Dim P$
    For J = 0 To U
        If J = 0 Then P = "|  Set" Else P = ""
'        Push Ay1, VblAlign(Ay(J), Pfx:=P, IdentOpt:=6)
    Next
    Vbl = JnVBar(Ay1)
PSet_FF_ExprAy = Vbl
End Function

Private Sub Z_PSetFFEqvy()
Dim Fny$(), ExprVblAy$()
Fny = SyzSS("a b c d")
Push ExprVblAy, "1sdfkl|lskdfj|skldfjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "2sdfkl|lskdfjdf| sdf"
Push ExprVblAy, "3sdfkl|fjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "4sf| sdf"
    Act = PSet_Fny_Evy(Fny, ExprVblAy)
'Debug.Print LineszVbl(Act)
End Sub


Private Function PSet_Fny_Evy$(Fny$(), EqVy)

End Function

Private Function QNm$(T)
QNm = QuoteSq(T)
End Function

Private Function PUpd_T$(T)
PUpd_T = KwUpd & C_T & QNm(T)
End Function

Private Function PWh_F_Eqv(F$, EqVal) ' Ssk is single-Sk-value
PWh_F_Eqv = C_Wh & QNm(F) & "=" & QV(EqVal)
End Function

Private Function PWh_T_EqK$(T, K&)
PWh_T_EqK = PWh_F_Eqv(T & "Id", K)
End Function

Private Function PWhBet_F_Fm_To$(F$, FmV, ToV)
PWhBet_F_Fm_To = C_Wh & QNm(F) & " " & KwBet & QV(FmV) & " " & KwAnd & " " & QV(ToV)
End Function

Private Function QV$(V)
QV = SqlQuote(V)
End Function

Private Function PExpr_F_InAy$(F, InVy)

End Function

Private Function PWh_F_InVy$(F$, InVy)
PWh_F_InVy = C_Wh & PExpr_F_InAy(F, InVy)
End Function

Private Sub Z_PWh_F_InVy()
Dim F$, Vy()
F = "A"
Vy = Array(1, "2", #2/1/2017#)
Ept = " where A=1 and B='2' and C=#2017-2-1#"
GoSub Tst
Exit Sub
Tst:
    Act = PWh_F_InVy(F, Vy)
    C
    Return
End Sub

Private Function PBexpr_Fny_EqVy$(Fny$(), EqVy)

End Function

Private Function PWh_Fny_EqVy$(Fny$(), EqVy)
PWh_Fny_EqVy = C_Wh & PBexpr_Fny_EqVy(Fny, EqVy)
End Function

Private Function PWh$(Bexpr$)
If Bexpr = "" Then Exit Function
PWh = C_Wh & Bexpr
End Function

Private Sub Z_PSet_Fny_VyFmt()
Dim Fny$(), Vy()
Ept = LineszVbl("|  Set|" & _
"    [A xx] = 1                     ,|" & _
"    B      = '2'                   ,|" & _
"    C      = #2018-12-01 12:34:56# ")
Fny = TermAy("[A xx] B C"): Vy = Array(1, "2", #12/1/2018 12:34:56 PM#): GoSub Tst
Exit Sub
Tst:
    Act = PSet_Fny_Vy(Fny, Vy)
    C
    Return
End Sub

Private Sub Z_PWhFldInVy_StrPAy()

End Sub

Function FmtExprVblAy(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional Sep$ = ",") As String()
Ass IsVblAy(ExprVblAy)
Dim Ident%
    If IdentOpt > 0 Then
        Ident = IdentOpt
    Else
        Ident = 0
    End If
    If Ident = 0 Then
        If Pfx <> "" Then
            Ident = Len(Pfx)
        End If
    End If
Dim O$(), P$, S$, U&, J&
U = UB(ExprVblAy)
Dim W%
'    W = VblWdtAy(ExprVblAy)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If J = U Then S = "" Else S = Sep
'    Push O, VblAlign(ExprVblAy(J), IdentOpt:=Ident, Pfx:=P, WdtOpt:=W, Sfx:=S)
Next
FmtExprVblAy = O
End Function

Private Sub ZZ_SqlSel_Fny_Ey_Into_T_OB()
Dim Fny$(), Ey$(), Into$, T$, Bexpr$
GoSub ZZ
Exit Sub
ZZ:
    Fny = SyzSS("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Ey = TermAy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Into = "#IZHT086"
    T = ">ZHT086"
    Bexpr = ""
    Debug.Print SqlSel_Fny_Extny_Into_T_OB(Fny, Ey, Into, T, Bexpr)
    Return
End Sub

Function SqlSel_Fny_Extny_Into_T_OB$(Fny$(), Extny$(), Into, T, Optional Bexpr$)
SqlSel_Fny_Extny_Into_T_OB = PSel_Fny_Extny(Fny, Extny) & PInto_T(Into) & PFm_T(T) & PWh(Bexpr)
End Function

Function SqlSel_Dist_Fny_EDict_Into_T_Wh_Gp_Ord$(IsDist As Boolean, Fny$(), EDic As Dictionary, T$, Wh$, Gp$, Ord$)

End Function
Function SqlSel_FF_EDic_Into_T_OB$(FF$, EDic As Dictionary, Into, T, Optional Bexpr$)
Dim Fny$(): Fny = SyzSS(FF)
Dim ExprAy$(): ExprAy = SyzDicKy(EDic, Fny)
Stop
SqlSel_FF_EDic_Into_T_OB = SqlSel_Fny_Extny_Into_T_OB(Fny, ExprAy, Into, T, Bexpr)
End Function

Function FnyzPfxN(Pfx$, N%) As String()
Dim J%
For J = 1 To N
    PushI FnyzPfxN, Pfx & J
Next
End Function

Function NsetzNN(NN$) As Aset
Set NsetzNN = AsetzAy(SyzSS(NN))
End Function

Function SqlQuote$(V)
Dim O$, C$
C = SqlQuoteChr(V)
If C <> "" Then SqlQuote = Quote(CStr(V), C): Exit Function
Select Case True
Case IsBool(V): O = IIf(V, "true", "false")
Case IsEmpty(V), IsNull(V), IsNothing(V): O = "null"
Case Else: O = V
End Select
SqlQuote = O
End Function

Function SqlQuoteChrzT$(A As Dao.DataTypeEnum)
Select Case A
Case _
    Dao.DataTypeEnum.dbBigInt, _
    Dao.DataTypeEnum.dbByte, _
    Dao.DataTypeEnum.dbCurrency, _
    Dao.DataTypeEnum.dbDecimal, _
    Dao.DataTypeEnum.dbDouble, _
    Dao.DataTypeEnum.dbFloat, _
    Dao.DataTypeEnum.dbInteger, _
    Dao.DataTypeEnum.dbLong, _
    Dao.DataTypeEnum.dbNumeric, _
    Dao.DataTypeEnum.dbSingle: Exit Function
Case _
    Dao.DataTypeEnum.dbChar, _
    Dao.DataTypeEnum.dbMemo, _
    Dao.DataTypeEnum.dbText: SqlQuoteChrzT = "'"
Case _
    Dao.DataTypeEnum.dbDate: SqlQuoteChrzT = "#"
Case Else
    Thw CSub, "Invalid DaoTy", "DaoTy", A
End Select
End Function

Function SqlQuoteChr$(V)
Dim O$
Select Case True
Case IsStr(V): O = "'"
Case IsDate(V): O = "#"
End Select
SqlQuoteChr = O
End Function
Function SqlQuoteVy(Vy) As String()
Dim V
For Each V In Vy
    PushI SqlQuoteVy, SqlQuote(V)
Next
End Function

Function SqlUpd_T_FF_EqDr_Whff_Eqvy$(T, FF$, Dr, WhFF$, EqVy)
SqlUpd_T_FF_EqDr_Whff_Eqvy = PUpd_T(T) & PSet_FF_EqDr(FF$, Dr) & PWh_FF_Eqvy(WhFF, EqVy)
End Function

Private Function PWh_FF_Eqvy$(FF$, EqVy)

End Function


Function SqlSel_FF_T_Bexpr$(FF$, T, Bexpr$)

End Function

Function SqlAddCol_T_Fny_FzDiSqlTy$(T, Fny$(), FzDiSqlTy As Dictionary)
Dim O$(), F
For Each F In Fny
    PushI O, F & " " & ValzDicK(FzDiSqlTy, F, "FzDiSqlTy", "Fld")
Next
SqlAddCol_T_Fny_FzDiSqlTy = FmtQQ("Alter Table [?] add column ?", T, JnComma(O))
End Function

Function SqlCrtPk_T$(T)
SqlCrtPk_T = FmtQQ("Create Index PrimaryKey on [?] (?Id) with Primary", T, T)
End Function

Function SqlCrtSk_T_SkFF$(T, Skff$)
SqlCrtSk_T_SkFF = SqlCrtSk_T_SkFny(T, Ny(Skff))
End Function

Function SqlCrtSk_T_SkFny$(T, SkFny$())
SqlCrtSk_T_SkFny = FmtQQ("Create unique Index SecondaryKey on [?] (?)", T, JnComma(QuoteSqzAy(SkFny)))
End Function

Function SqlCrtTbl_T_X$(T, X$)
SqlCrtTbl_T_X = FmtQQ("Create Table [?] (?)", T, X)
End Function

Function SqlDrpCol_T_F$(T, F$)
SqlDrpCol_T_F = FmtQQ("Alter Table [?] drop column [?]", T, F$)
End Function

Function SqlDrpTbl_T$(T)
SqlDrpTbl_T = "Drop Table [" & T & "]"
End Function

Function SqlIns_T_FF_Dr$(T, FF$, Dr)
Dim Fny$(): Fny = SyzSS(FF)
ThwIf_DifSi Fny, Dr, CSub
Dim A$, B$
A = JnComma(QuoteSqzAyIf(Fny))
B = JnComma(SqlQuoteVy(Dr))
SqlIns_T_FF_Dr = FmtQQ("Insert Into [?] (?) Values(?)", T, A, B)
End Function
Function SqlSel_T$(T, Optional Bexpr$)
SqlSel_T = "Select *" & PFm_T(T) & PWh(Bexpr)
End Function

Function SqlSel_T_Wh$(T, Bexpr$)
SqlSel_T_Wh = SqlSel_T(T) & PWh(Bexpr)
End Function

Function SqlSel_Into_T_WhFalse(Into, T)
SqlSel_Into_T_WhFalse = FmtQQ("Select * Into [?] from [?] where false", Into, T)
End Function

Function SqlSel_F$(F$)
SqlSel_F = SqlSel_F_T(F, F)
End Function

Function SqlSel_F_T$(F$, T, Optional Bexpr$)
SqlSel_F_T = FmtQQ("Select [?] from [?]?", F, T, PWh(Bexpr))
End Function


Function SqlSel_FF_T_Ord(FF$, T, OrdMinusSfxFF$)
SqlSel_FF_T_Ord = PSel_FF(FF) & PFm_T(T) & POrd_MinusSfxFF(OrdMinusSfxFF)
End Function

Function SqlUpd_T_Sk_Fny_Dr$(T, Sk$(), Fny$(), Dr)
If Si(Sk) = 0 Then Stop
Dim PUpd_T$, Set_$, Wh$: GoSub X_PUpd_T_Set_Wh
'UpdSql = PUpd_T & Set_ & Wh
Exit Function
X_PUpd_T_Set_Wh:
    Dim Fny1$(), Dr1(), Skvy(): GoSub X_Fny1_Dr1_SkVy
    PUpd_T = "Update [" & T & "]"
    Set_ = PSet_Fny_Vy(Fny1, Dr1)
    Wh = PWh_Fny_EqVy(Sk, Skvy)
    Return
X_Ay:
    Dim L$(), R$()
    L = AlignQuoteSq(Fny)
    R = SqlQuoteVy(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, Ixy%(), I%
    For Each Ski In Sk
'        I = IxzAy(Fny, Ski)
        If I = -1 Then Stop
        Push Ixy, I
        Push Skvy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not HasEle(Ixy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Private Function PSet_Fny_Vy1$(Fny$(), Vy())
Dim A$: GoSub X_A
PSet_Fny_Vy1 = "  Set " & A
Exit Function
X_A:
    Dim L$(): L = QuoteSqzAy(Fny)
    Dim R$(): R = SqlQuoteVy(Vy)
    Dim J%, O$()
    For J = 0 To UB(L)
        Push O, L(J) & " = " & R(J)
    Next
    A = JnComma(O)
    Return
End Function

Private Sub Z_SqlDtlTWhfInAset()
Dim T$, F$, S As Aset, SqlWdt%
T = "Tbl-1"
F = "Fld-1"
T1:
    Set S = AsetNRndStr(1000)
    GoTo Tst
T2:
    Set S = AsetNRndInt(1000)
Tst:
    D SqyDlt_T_WhFld_InAset(T, F, S)
    Return
End Sub

Function SqlDlt_T$(T)
SqlDlt_T = "Delete * from [" & T & "]"
End Function

Function SqlDlt_T_Wh$(T, Bexpr$)
SqlDlt_T_Wh = SqlDlt_T(T) & PWh(Bexpr)
End Function

Function SqyDlt_T_WhFld_InAset(T, F, S As Aset, Optional SqlWdt% = 3000) As String()
Dim A$
Dim Ey$()
    A = SqlDlt_T(T) & " Where "
    Ey = PFldInX_F_InAset_Wdt(F, S, SqlWdt - Len(A))
Dim E
For Each E In Ey
    PushI SqyDlt_T_WhFld_InAset, A & E & vbCrLf
Next
End Function


Function LyJnSqlCommaAsetW(A As Aset, W%) As String()

End Function



Function SqlIns_T_FF_ValAp$(T, FF$, ParamArray ValAp())
Dim Av(): Av = ValAp
SqlIns_T_FF_ValAp = PIns_T(T) & PBkt_FF(FF) & " Values" & PBkt_Av(Av)
End Function


Function SqlSel_Fny_T(Fny$(), T, Optional Bexpr$, Optional IsDis As Boolean)
SqlSel_Fny_T = PSel_Fny(Fny, IsDis) & PFm_T(T) & PWh(Bexpr)
End Function

Function SqlSel_FF_T_WhF_InVy$(FF, T, WhF$, InVy, Optional IsDis As Boolean)
Dim W$
W = PExpr_F_InAy(WhF$, InVy)
SqlSel_FF_T_WhF_InVy = SqlSel_FF_T(FF, T, IsDis)
End Function

Function SqlSelDis_FF_T$(FF$, T)
SqlSelDis_FF_T = SqlSel_FF_T(FF$, T, IsDis:=True)
End Function

Function SqlSel_FF_ExprDic_T$(FF$, E As Dictionary, T, Optional IsDis As Boolean)
'SelFFExprDicP = "Select" & vbCrLf & FFExprDicAsLines(FF$, ExprDic)
End Function

Function SqlSel_T_WhId$(T, Id&)
SqlSel_T_WhId = PSel_T(T) & " " & PWh_T_Id(T, Id)
End Function


Private Function PWh_T_Id$(T, Id)
PWh_T_Id = PWh(FmtQQ("[?]Id=?", T, Id))
End Function


Function SqlSel_FF_Into_T$(FF$, Into$, T, Optional Bexpr$, Optional Dis As Boolean)
SqlSel_FF_Into_T = PSel_FF(FF) & PInto_T(Into) & PFm_T(T) & PWh(Bexpr)
End Function

Function SqlSel_Fny_T_WhFny_EqVy$(Fny$(), T, WhFny$(), EqVy)
SqlSel_Fny_T_WhFny_EqVy = SqlSel_Fny_T(Fny, T, PWh_Fny_EqVy(WhFny, EqVy))
End Function

Function SqlSel_X_Into_T_OB_OGp_OOrd_ODis$(X$, Into$, T$, Optional OBexpr$, Optional OGp$, Optional OOrd$, Optional ODis As Boolean)
SqlSel_X_Into_T_OB_OGp_OOrd_ODis$ = PSel_X(X, ODis) & PInto_T(Into) & PFm_T(T) & PWh(OBexpr) & PGp(OGp) & POrd(OOrd$)
End Function
Private Function PGp$(Gp$)

End Function
Private Function POrd$(Ord$)

End Function
Function SqlSel_Fny_Into_T_OB$(Fny$(), Into$, T, Optional Bexpr$)

End Function

Function SqlSel_X_Into_T$(X$, Into$, T, Optional Bexpr$)
SqlSel_X_Into_T = PSel_X(X) & PFm_T(T) & PWh(Bexpr)
End Function

Function SqlSel_X_T$(X$, T, Optional Bexpr$)
SqlSel_X_T = PSel_X(X) & PFm_T(T) & PWh(Bexpr)
End Function

Function SqlSel_FF_T_Ordff$(FF$, T, OrdMinusSfxFF$)
SqlSel_FF_T_Ordff = PSel_FF(FF) & PFm_T(T) & POrd_MinusSfxFF(OrdMinusSfxFF)
End Function

Function SqlSelCnt_T_OB$(T, Optional Bexpr$)
SqlSelCnt_T_OB = "Select Count(*)" & PFm_T(T) & PWh(Bexpr)
End Function

Function SqyCrtPkzTny(Tny$()) As String()
Dim T
For Each T In Itr(Tny)
    PushI SqyCrtPkzTny, SqlCrtPk_T(T)
Next
End Function

Function SqlSel_F_T_F_Ev$(F$, T, WhFld$, Ev())
SqlSel_F_T_F_Ev = SqlSel_F_T(F, T, PExpr_F_InAy(WhFld, Ev))
End Function

Private Function Bexpr_Fny_Vy$(Fny$(), Vy())
End Function

Private Function Bexpr_F_Ev$(F$, Ev)
Bexpr_F_Ev = QuoteSq(F) & "=" & SqlQuote(Ev)
End Function

Private Sub ZZZ()

End Sub
