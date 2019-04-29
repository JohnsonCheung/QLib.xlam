Attribute VB_Name = "MDao_Sql"
Option Explicit
Const CMod$ = "MSql."
Public Const DocOfNN$ = "It a Str or Sy will give Ny.  See TermAy"
Const KwBet$ = "between"
Const KwUpd$ = "update"
Const KwInto$ = "into"
Const KwSel$ = "select"
Const KwSelDis$ = "select distinct"
Const KwFm$ = "from"
Const KwGp$ = "group by"
Const KwWh$ = "where"
Const KwAnd$ = "and"
Const KwJn$ = "join"
Const KwOr$ = "or"
Const KwOrd$ = "order by"
Const KwLeftJn$ = "left join"
Public FmtSql As Boolean

Private Property Get C_Fm$()
C_Fm = C_NLT & KwFm & C_T
End Property

Private Property Get C_Into$()
C_Into = C_NLT & KwInto & C_T
End Property

Private Property Get C_NL$() ' New Line
If FmtSql Then
    C_NL = vbCrLf
Else
    C_NL = " "
End If
End Property

Private Property Get C_T$()
If FmtSql Then
    C_T = vbTab
Else
    C_T = " "
End If
End Property

Private Property Get C_NLT$() ' New Line Tabe
If FmtSql Then
    C_NLT = C_NL & C_T
Else
    C_NLT = " "
End If
End Property

Private Property Get C_NLTT$() ' New Line Tabe
If FmtSql Then
    C_NLTT = C_NLT & C_T
Else
    C_NLTT = " "
End If
End Property

Function Sqp_F_InAy$(F$, InAy)
Stop
'Sqp_F_InAy = QuoteSqBkt(F) & "(" & JnComma(SyQuote(InAy, SqlQuoteChr(InAy(0)))) & ")"
End Function

Function FFJnComma$(FF$)
FFJnComma = JnComma(TermAy(FF))
End Function

Function SqpInto_T$(T$)
SqpInto_T = C_Into & "[" & T & "]"
End Function

Function BexprRecId$(T$, RecId)
BexprRecId = FmtQQ("?Id=?", T, RecId)
End Function

Function SqpSet_Fny_Vy$(Fny$(), Vy())
Dim F$(): F = SyQuoteSq(Fny)
Dim V$(): V = SyQuoteSql(Vy)
SqpSet_Fny_Vy = JnComma(JnAyab(F, V, "="))
End Function

Function SqpAnd_Bexpr$(Bexpr$)
If Bexpr = "" Then Exit Function
'SqpAnd_Bexpr = NxtLin & "and " & NxtLin_Tab & Bexpr
End Function

Private Function AddPfxNLTT$(Sy$())
AddPfxNLTT = Jn(SyAddPfx(Sy, C_NLTT), "")
End Function

Private Function ExprInLis_InLisBexpr$(Expr$, InLis$)
If InLis = "" Then Exit Function
ExprInLis_InLisBexpr = FmtQQ("? in (?)", Expr, InLis)
End Function

Function SqpSel_F$(F$)
SqpSel_F = "Select [" & F & "]"
End Function

Function SqpSel_X$(X$, Optional Dis As Boolean)
SqpSel_X = SqpSel_Dis(Dis) & X
End Function

Function SqpFm$(T$)
SqpFm = C_Fm & QuoteSq(T$)
End Function

Function SqpGp_ExprVblAy$(ExprVblAy$())
SqpGp_ExprVblAy = VblFmtAyAsLines(ExprVblAy, "|  Group By")
End Function

Private Sub Z_SqpGp_ExprVblAy()
Dim ExprVblAy$()
    Push ExprVblAy, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push ExprVblAy, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push ExprVblAy, "3sldkfjsdf"
DmpAy SplitVBar(SqpGp_ExprVblAy(ExprVblAy))
End Sub


Function SqpSelX_FF_ExtNy$(FF$, ExtNy$())
Dim Fny$(): Fny = TermAy(FF)
Dim P1$()
    Dim J%, M$
    For J = 0 To UB(Fny)
        If Fny(J) = ExtNy(J) Or ExtNy(J) = "" Then
            M = ""
        Else
            M = QuoteSq(ExtNy(J))
        End If
        PushI P1, M
    Next
    If FmtSql Then
        P1 = FmtLinesAy(P1)
        For J = 0 To UB(P1)
            If Trim(P1(J)) = "" Then
                P1(J) = P1(J) & "    "
            Else
                P1(J) = P1(J) & " As "
            End If
        Next
    Else
        For J = 0 To UB(P1)
            P1(J) = ApdIf(P1(J), " As ")
        Next

    End If
Dim P2$(): If FmtSql Then P2 = FmtAySamWdt(Fny) Else P2 = Fny
SqpSelX_FF_ExtNy = KwSel & C_T & JnComma(SyAddPfx(JnAyab(P1, P2), C_NLTT))
End Function

Function SqpSel_FF_Ey$(FF$, ExprAy$())
SqpSel_FF_Ey = SqpSel_X(SqpSelX_FF_ExtNy(FF$, ExprAy))
End Function

Function JnCommaSpcFF$(FF$)
JnCommaSpcFF = JnCommaSpc(TermAy(FF))
End Function

Function SqpSel_FF$(FF$, Optional Dis As Boolean)
SqpSel_FF = SqpSel_Fny(SySsl(FF), Dis)
End Function

Function SqpSel_Fny$(Fny$(), Optional Dis As Boolean)
'SqpSel_FF = SqpSel_Dis(Dis) & C_NLTT & JnCommaSpc(Fny)
Stop
End Function

Function SqpSel_Dis$(Dis As Boolean)
If Dis Then
    SqpSel_Dis = KwSelDis
Else
    SqpSel_Dis = KwSel
End If
End Function

Private Sub Z_SqpSel()
Dim Fny$(), ExprVblAy$()
ExprVblAy = Sy("F1-Expr", "F2-Expr   AA|BB    X|DD       Y", "F3-Expr  x")
Fny = SplitSpc("F1 F2 F3xxxxx")
'Debug.Print LineszVbl(SqpSelFFFldLvs(Fny, ExprVblAy))
End Sub

Function SqlSel_FF$(FF$, Optional IsDis As Boolean)
SqlSel_FF = SqpSel_X(FFJnComma(FF), IsDis)
End Function

Function SqpSet_FF_Ey$(FF$, Ey$())
Const CSub$ = CMod & "SqpSet_FF_Ey"
Dim Fny$(): Fny = SySsl(FF)
Ass IsVblAy(Ey)
If Si(Fny) <> Si(Ey) Then Thw CSub, "[FF-Sz} <> [Si-Ey], where [FF],[Ey]", Si(Fny), Si(Ey), FF, Ey
Dim AFny$()
    AFny = FmtAySamWdt(Fny)
    AFny = SyAddSfx(AFny, " = ")
Dim W%
    'W = VblWdtAy(Ey)
Dim Ident%
    W = AyWdt(AFny)
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
SqpSet_FF_Ey = Vbl
End Function

Private Sub Z_SqpSetFFEqvy()
Dim Fny$(), ExprVblAy$()
Fny = SySsl("a b c d")
Push ExprVblAy, "1sdfkl|lskdfj|skldfjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "2sdfkl|lskdfjdf| sdf"
Push ExprVblAy, "3sdfkl|fjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "4sf| sdf"
    Act = SqpSet_Fny_Evy(Fny, ExprVblAy)
'Debug.Print LineszVbl(Act)
End Sub

Function SqpSet_FF_Evy$(FF$, EqVy)

End Function
Function SqpSet_Fny_Evy$(Fny$(), EqVy)

End Function

Private Function QNm$(T$)
QNm = QuoteSq(T$)
End Function

Function SqpUpd_T$(T$)
SqpUpd_T = KwUpd & C_T & QNm(T$)
End Function

Function SqpWhfv(F$, V) ' Ssk is single-Sk-value
SqpWhfv = C_Wh & QNm(F) & "=" & QV(V)
End Function

Function SqpWhK$(K&, T$)
SqpWhK = SqpWhfv(T & "Id", K)
End Function

Function SqpWhBet_F_Fm_To$(F$, FmV, ToV)
SqpWhBet_F_Fm_To = C_Wh & QNm(F) & " " & KwBet & QV(FmV) & " " & KwAnd & " " & QV(ToV)
End Function

Private Function QV$(V)
QV = SqlVal(V)
End Function

Private Property Get C_And$()
If FmtSql Then
    C_And = C_NLT & KwAnd & C_T
Else
    C_And = " " & KwAnd & " "
End If
End Property

Private Property Get C_Wh$()
C_Wh = C_NLT & KwWh & C_NLT
End Property

Function SqpWh_F_InVy$(F$, InVy)
SqpWh_F_InVy = C_Wh & Sqp_F_InAy(F, InVy)
End Function

Private Sub Z_SqpWh_F_InVy()
Dim F$, Vy()
F = "A"
Vy = Array(1, "2", #2/1/2017#)
Ept = " where A=1 and B='2' and C=#2017-2-1#"
GoSub Tst
Exit Sub
Tst:
    Act = SqpWh_F_InVy(F, Vy)
    C
    Return
End Sub

Private Function FnyEqVy_Bexpr$(Fny$(), EqVy)

End Function

Function SqpWh_FnyEqVy$(Fny$(), EqVy)
SqpWh_FnyEqVy = C_Wh & FnyEqVy_Bexpr(Fny, EqVy)
End Function

Function SqpWh$(A$)
If A = "" Then Exit Function
SqpWh = C_Wh & A
End Function

Private Sub Z_SqpSet_Fny_VyFmt()
Dim Fny$(), Vy()
Ept = LineszVbl("|  Set|" & _
"    [A xx] = 1                     ,|" & _
"    B      = '2'                   ,|" & _
"    C      = #2018-12-01 12:34:56# ")
Fny = TermAy("[A xx] B C"): Vy = Array(1, "2", #12/1/2018 12:34:56 PM#): GoSub Tst
Exit Sub
Tst:
    Act = SqpSet_Fny_Vy(Fny, Vy)
    C
    Return
End Sub

Private Sub Z_SqpWhFldInVy_StrSqpAy()

End Sub

Function VblFmtAyAsLines$(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAy, Optional Sep$ = ",")
VblFmtAyAsLines = JnVBar(VblFmtAyAsLy(ExprVblAy, Pfx, IdentOpt, SfxAy, Sep))
End Function

Function VblFmtAyAsLy(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAyOpt, Optional Sep$ = ",") As String()
Dim NoSfxAy As Boolean
Dim SfxWdt%
Dim SfxAy$()
    NoSfxAy = IsEmp(SfxAy)
    If Not NoSfxAy Then
        Ass IsSy(SfxAyOpt)
        SfxAy = FmtAySamWdt(SfxAyOpt)
        Dim U%, J%: U = UB(SfxAy)
        For J = 0 To U
            If J <> U Then
                SfxAy(J) = SfxAy(J) & Sep
            End If
        Next
    End If
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
Dim O$(), P$, S$
U = UB(ExprVblAy)
Dim W%
'    W = VblWdtAy(ExprVblAy)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If NoSfxAy Then
        If J = U Then S = "" Else S = Sep
    Else
        If J = U Then S = SfxAy(J) Else S = SfxAy(J) & Sep
    End If
'    Push O, VblAlign(ExprVblAy(J), IdentOpt:=Ident, Pfx:=P, WdtOpt:=W, Sfx:=S)
Next
VblFmtAyAsLy = O
End Function

Function SqlSel_FF_EDic_T$(FF$, EDic As Dictionary, T$, Optional IsDis As Boolean)
SqlSel_FF_EDic_T = SqlSel_FF_Ey_T(FF$, SyzDicKy(EDic, TermAy(FF)), T, IsDis)
End Function

Function SqlSel_FF_T$(FF$, T$, Optional IsDis As Boolean, Optional Bexpr$)
SqlSel_FF_T = SqpSel_FF(FF$, IsDis) & SqpFm(T$) & SqpWh(Bexpr)
End Function

Function SqlSel_FF_Ey_T$(FF$, Ey$(), T$, Optional IsDis As Boolean, Optional Bexpr$)
SqlSel_FF_Ey_T = SqpSel_X(SqpSelX_FF_ExtNy(FF$, Ey), IsDis) & SqpFm(T$) & SqpWh(Bexpr)
End Function

Function ItrzTT(TT$)
Asg Itr(TermAyzTT(TT)), ItrzTT
End Function

Function FnyzPfxN(Pfx$, N%) As String()
Dim J%
For J = 1 To N
    PushI FnyzPfxN, Pfx & J
Next
End Function

Function NsetzNN(NN$) As Aset
Set NsetzNN = AsetzAy(TermAy(NN))
End Function

Function SqlVal$(V)
Dim O$, C$
C = SqlQuoteChr(V)
If C <> "" Then SqlVal = Quote(CStr(V), C): Exit Function
Select Case True
Case IsBool(V): O = IIf(V, "true", "false")
Case IsEmpty(V), IsNull(V), IsNothing(V): O = "null"
Case Else: O = V
End Select
SqlVal = O
End Function
Function SqlQuoteChr$(V)
Dim O$
Select Case True
Case IsStr(V): O = "'"
Case IsDate(V): O = "#"
End Select
SqlQuoteChr = O
End Function
Function SyQuoteSql(Ay) As String()
Dim V
For Each V In Ay
    PushI SyQuoteSql, SqlVal(V)
Next
End Function

Function SqlUpd_T_FF_EqDr_Whff_Eqvy$(T$, FF$, Dr, WhFF$, EqVy)
SqlUpd_T_FF_EqDr_Whff_Eqvy = SqpUpd_T(T$) & SqpSet_FF_EqDr(FF$, Dr) & SqpWh_FF_Eqvy(WhFF, EqVy)
End Function

Function SqpWh_FF_Eqvy$(FF$, EqVy)

End Function

Function SqpSet_FF_EqDr$(FF$, EqDr)

End Function

Function SqlSel_FF_T_Bexpr$(FF$, T$, Bexpr$)

End Function

Function QAddCol$(T$, Fny$(), F As Drs, E As Dictionary)
Dim O$(), Fld
For Each Fld In Fny
'    PushI O, Fld & " " & QAddCol1(Fld, F, E)
Next
QAddCol = FmtQQ("Alter Table [?] add column ?", T, JnComma(O))
End Function

Function SqlCrtPkzT$(T$)
SqlCrtPkzT = FmtQQ("Create Index PrimaryKey on [?] (?Id) with Primary", T, T)
End Function

Function SqlCrtSk_T_SkFF$(T$, Skff$)
SqlCrtSk_T_SkFF = SqlCrtSk_T_SkFny(T, Ny(Skff))
End Function

Function SqlCrtSk_T_SkFny$(T$, SkFny$())
SqlCrtSk_T_SkFny = FmtQQ("Create unique Index SecondaryKey on [?] (?)", T, JnComma(SyQuoteSq(SkFny)))
End Function

Function SqlCrtTbl_T_X$(T$, X$)
SqlCrtTbl_T_X = FmtQQ("Create Table [?] (?)", T, X)
End Function

Function SqlDrpCol_T_F$(T$, F$)
SqlDrpCol_T_F = FmtQQ("Alter Table [?] drop column [?]", T, F$)
End Function

Function SqlDrpTbl_T$(T$)
SqlDrpTbl_T = "Drop Table [" & T & "]"
End Function

Function SqlIns_T_FF_Dr$(T$, FF$, Dr)
Dim Fny$(): Fny = TermAy(FF)
ThwDifSz Fny, Dr, CSub
Dim A$, B$
A = JnComma(SyQuoteSqIf(Fny))
B = JnComma(SyQuoteSql(Dr))
SqlIns_T_FF_Dr = FmtQQ("Insert Into [?] (?) Values(?)", T, A, B)
End Function
Function SqlSel_T$(T$, Optional Bexpr$)
SqlSel_T = "Select *" & SqpFm(T$) & SqpWh(Bexpr)
End Function

Function SqlSel_T_Wh$(T$, Bexpr$)
SqlSel_T_Wh = SqlSel_T(T$) & SqpWh(Bexpr)
End Function

Function SqlSel_Into_T_WhFalse(Into$, T$)
SqlSel_Into_T_WhFalse = FmtQQ("Select * Into [?] from [?] where false", Into, T)
End Function

Function SqlSel_F$(F$)
SqlSel_F = SqlSel_F_T(F, F)
End Function

Function SqlSel_F_T$(F$, T$, Optional Bexpr$)
SqlSel_F_T = FmtQQ("Select [?] from [?]?", F, T, SqpWh(Bexpr))
End Function

Function SqpOrd_MinusSfxFF$(OrdMinusSfxFF$)
If OrdMinusSfxFF = "" Then Exit Function
Dim O$(): O = SySsl(OrdMinusSfxFF)
Dim I, J%
For Each I In O
    If HasSfx(O(J), "-") Then
        O(J) = RmvSfx(O(J), "-") & " desc"
    End If
    J = J + 1
Next
SqpOrd_MinusSfxFF = C_NLT & "order by " & JnCommaSpc(O)
End Function

Function SqlSel_FF_T_Ord(FF$, T$, OrdMinusSfxFF$)
SqlSel_FF_T_Ord = SqpSel_FF(FF) & SqpFm(T$) & SqpOrd_MinusSfxFF(OrdMinusSfxFF)
End Function

Function SqlUpd_T_Sk_Fny_Dr$(T$, Sk$(), Fny$(), Dr)
If Si(Sk) = 0 Then Stop
Dim SqpUpd_T$, Set_$, Wh$: GoSub X_SqpUpd_T_Set_Wh
'UpdSql = SqpUpd_T & Set_ & Wh
Exit Function
X_SqpUpd_T_Set_Wh:
    Dim Fny1$(), Dr1(), Skvy(): GoSub X_Fny1_Dr1_SkVy
    SqpUpd_T = "Update [" & T & "]"
    Set_ = SqpSet_Fny_Vy(Fny1, Dr1)
    Wh = SqpWh_FnyEqVy(Sk, Skvy)
    Return
X_Ay:
    Dim L$(), R$()
    L = FnyAlignQuote(Fny)
    R = SyQuoteSql(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, IxAy%(), I%
    For Each Ski In Sk
'        I = IxzAy(Fny, Ski)
        If I = -1 Then Stop
        Push IxAy, I
        Push Skvy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not HasEle(IxAy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Function SqpSet_Fny_Vy1$(Fny$(), Vy())
Dim A$: GoSub X_A
SqpSet_Fny_Vy1 = "  Set " & A
Exit Function
X_A:
    Dim L$(): L = SyQuoteSq(Fny)
    Dim R$(): R = SyQuoteSql(Vy)
    Dim J%, O$()
    For J = 0 To UB(L)
        Push O, L(J) & " = " & R(J)
    Next
    A = JnComma(O)
    Return
End Function

Function FnyAlignQuote(Fny$()) As String()
FnyAlignQuote = FmtAySamWdt(SyQuoteSq(Fny))
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
    D SqyDlt_T_WhFld_InAset(T$, F, S)
    Return
End Sub

Function SqlDlt_T$(T$)
SqlDlt_T = "Delete * from [" & T & "]"
End Function

Function SqlDlt_T_Wh$(T$, Bexpr$)
SqlDlt_T_Wh = SqlDlt_T(T$) & SqpWh(Bexpr)
End Function

Function SqyDlt_T_WhFld_InAset(T$, F, S As Aset, Optional SqlWdt% = 3000) As String()
Dim A$
Dim Ey$()
    A = SqlDlt_T(T$) & " Where "
    Ey = SqpFldInX_F_InAset_Wdt(F, S, SqlWdt - Len(A))
Dim E
For Each E In Ey
    PushI SqyDlt_T_WhFld_InAset, A & E & vbCrLf
Next
End Function

Function SqpFldInX_F_InAset_Wdt(F, S As Aset, Wdt%) As String()
Dim A$
    A = "[F] in ("
Dim I
'For Each I In LyJnQSqlCommaAsetW(S, Wdt - Len(A))
    PushI SqpFldInX_F_InAset_Wdt, I
'Next
End Function

Function LyJnSqlCommaAsetW(A As Aset, W%) As String()

End Function

Function SqpBexpr_F_Ev$(F, Ev)

End Function

Function SqpBktFF$(FF)
'SqpBktFF = QuoteBkt(JnCommaFF(FF))
End Function

Function JnCommaFF$(FF$)
JnCommaFF = JnComma(TermAy(FF))
End Function

Function SqlIns_T_FF_Valap$(T$, FF$, ParamArray ValAp())
Dim Av(): Av = ValAp
SqlIns_T_FF_Valap = SqpIns_T(T$) & SqpBktFF(FF) & " Values" & SqpBktAv(Av)
End Function

Function SqpIns_T$(T$)
SqpIns_T = "Insert into [" & T & "]"
End Function

Function SqpBktAv$(Av())
Dim O$(), I
For Each I In Av
    PushI O, SqlVal(I)
Next
SqpBktAv = QuoteBktJnComma(Av)
End Function

Function SqlSel_Fny_T(Fny$(), T$, Optional Bexpr$, Optional IsDis As Boolean)
SqlSel_Fny_T = SqpSel_Fny(Fny, IsDis) & SqpFm(T$) & SqpWh(Bexpr)
End Function

Function SqlSel_FF_T_WhF_InVy$(FF$, T$, WhF$, InVy, Optional IsDis As Boolean)
Dim W$
W = Sqp_F_InAy(WhF$, InVy)
SqlSel_FF_T_WhF_InVy = SqlSel_FF_T(FF$, T, IsDis, W)
End Function

Function QSelDis_FF_T$(FF$, T$)
QSelDis_FF_T = SqlSel_FF_T(FF$, T, IsDis:=True)
End Function

Function SqlSel_FF_ExprDic_T$(FF$, E As Dictionary, T, Optional IsDis As Boolean)
'SelFFExprDicSqp = "Select" & vbCrLf & FFExprDicAsLines(FF$, ExprDic)
End Function

Function SqlSel_T_WhId$(T$, Id&)
SqlSel_T_WhId = SqpSel_T(T) & " " & SqpWh_T_Id(T, Id)
End Function

Function SqpSel_T$(T$)
SqpSel_T = KwSel & C_T & "*" & SqpFm(T$)
End Function

Function SqpWh_T_Id$(T$, Id)
SqpWh_T_Id = SqpWh(FmtQQ("[?]Id=?", T, Id))
End Function

Function QSelDis_FF_ExprDic_T$(FF$, E As Dictionary, T$)
QSelDis_FF_ExprDic_T = SqlSel_FF_ExprDic_T(FF$, E, T, IsDis:=True)
End Function
Function SqpInto$(Into$)

End Function
Function SqlSel_FF_Into_T$(FF$, Into$, T$, Optional Bexpr$, Optional Dis As Boolean)
SqlSel_FF_Into_T = SqpSel_FF(FF) & SqpInto(Into) & SqpFm(T) & SqpWh(Bexpr)
End Function

Function SqlSel_FF_T_WhFny_EqVy$(FF$, T$, Fny$(), EqVy)
SqlSel_FF_T_WhFny_EqVy = SqlSel_FF_T(FF$, T, SqpWh_FnyEqVy(Fny, EqVy))
End Function

Function SqlSel_FF_ExtNy_Into_T$(FF$, ExtNy$(), Into$, T$, Optional Bexpr$)
SqlSel_FF_ExtNy_Into_T = SqpSelX_FF_ExtNy(FF$, ExtNy) & SqpInto_T(Into) & SqpFm(T$) & SqpWh(Bexpr)
End Function

Function SqlSel_FF_Into_T_WhFalse$(FF$, Into$, T$)

End Function

Function SqlSel_X_Into_T$(X$, Into$, T$, Optional Bexpr$)
SqlSel_X_Into_T = SqpSel_X(X) & SqpFm(T$) & SqpWh(Bexpr)
End Function

Function SqlSel_X_T$(X$, T$, Optional Bexpr$)
SqlSel_X_T = SqpSel_X(X) & SqpFm(T$) & SqpWh(Bexpr)
End Function

Function SqlSel_FF_T_Ordff$(FF$, T$, OrdMinusSfxFF$)
SqlSel_FF_T_Ordff = SqpSel_FF(FF) & SqpFm(T$) & SqpOrd_MinusSfxFF(OrdMinusSfxFF)
End Function

Function SqlSelCnt_T$(T$, Optional Bexpr$)
SqlSelCnt_T = "Select Count(*) " & SqpFm(T$) & SqpWh(Bexpr)
End Function

Function SqyCrtPkzTny(A$()) As String()
Dim T
For Each T In A
    PushI SqyCrtPkzTny, SqlCrtPkzT(CStr(T))
Next
End Function

Function SqlSel_F_T_F_Ev$(F$, T$, WhFld$, Ev())
SqlSel_F_T_F_Ev = SqlSel_F_T(F, T, Bexpr(WhFld, Ev))
End Function

Function BexprzFnyVy$(Fny$(), Vy())

End Function

Function Bexpr$(F$, Ev)
Bexpr = QuoteSq(F) & "=" & SqlVal(Ev)
End Function

