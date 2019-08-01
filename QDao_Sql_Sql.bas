Attribute VB_Name = "QDao_Sql_Sql"
Option Explicit
Option Compare Text
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Sql."
Const KwBet$ = "between"
Const KwSet$ = "set"
Const KwDis$ = "distinct"
Const KwUpd$ = "update"
Const KwInto$ = "into"
Const KwSel$ = "select"
Const KwFm$ = "from"
Const KwGp$ = "group by"
Const KwWh$ = "where"
Const KwAnd$ = "and"
Const KwOn$ = "on"
Const KwLJn$ = "left join"
Const KwIJn$ = "inner join"
Const KwOr$ = "or"
Const KwOrd$ = "order by"
Const KwLeftJn$ = "left join"
Type SelIntoPm: Fny() As String: Ey() As String: Into As String: T As String: Bexp As String: End Type
Type SelIntoPms: N As Byte: Ay() As SelIntoPm: End Type
Function SelIntoPm(Fny$(), Ey$(), Into$, T$, Optional Bexp$) As SelIntoPm
With SelIntoPm
    .Fny = Fny
    .Ey = Ey
    .Into = Into
    .T = T
    .Bexp = Bexp
End With
End Function

Sub PushIelIntoPm(O As SelIntoPms, M As SelIntoPm)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function SqyzSelIntoPms(A As SelIntoPms) As String()
Dim J As Byte
For J = 0 To A.N - 1
    PushI SqyzSelIntoPms, SqlzSelIntoPm(A.Ay(J))
Next
End Function

Function SqlzSelIntoPm$(A As SelIntoPm)
With A
'SqlzSelIntoPm = SqlSel_Fny_Extny_Into_T(.Fny, .Extny, .Into, .T, .Bexp)
End With
End Function


Private Function PSel_Fny_Extny_NOFMT(Fny$(), Extny$(), Optional IsDis As Boolean)
Dim O$(), J%, E$, F$
For J = 0 To UB(Fny)
    F = Fny(J)
    E = Trim(Extny(J))
    Select Case True
    Case E = "", E = F: PushI O, F
    Case Else: PushI O, QteSq(E) & " As " & F
    End Select
Next
PSel_Fny_Extny_NOFMT = KwSel & PDis(IsDis) & " " & JnCommaSpc(O)
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
Private Property Get C_TT$()
C_TT = C_T & C_T
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
C_Wh = C_NLT & KwWh & C_T
End Property

Private Property Get C_CommaSpc$()
If Fmt Then
    C_CommaSpc = C_CNLT
Else
    C_CommaSpc = ", "
End If
End Property


Private Function PBkt_FF$(FF$)
PBkt_FF = QteBkt(SyzSS(FF))
End Function

Private Function PBkt_Av$(Av())
Dim O$(), I
For Each I In Av
    PushI O, SqlQte(I)
Next
PBkt_Av = QteBktJnComma(Av)
End Function


Private Function PFldInX_F_InAset_Wdt(F, S As Aset, Wdt%) As String()
Dim A$
    A = "[F] in ("
Dim I
'For Each I In LyJnQSqlCommaAsetW(S, Wdt - Len(A))
    PushI PFldInX_F_InAset_Wdt, I
'Next
End Function

Private Function PFmzX$(FmX$)
PFmzX = PFm(FmX) & " x"
End Function

Private Function PFm$(Fm)
PFm = C_Fm & QteSq(Fm)
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
PSel_FF_Extny = PSelzX(PSel_Fny_Extny(Ny(FF), Extny))
End Function

Private Function PSel_Fny_Extny$(Fny$(), Extny$(), Optional IsDis As Boolean)
If Not Fmt Then PSel_Fny_Extny = PSel_Fny_Extny_NOFMT(Fny, Extny): Exit Function
Dim E$(), F$()
F = Fny
E = Extny
FEs_SetExtNm_ToBlnk_IfEqToFld F, E
FEs_SqQteExtNm_IfNB E
FEs_AlignExtNm E
FEs_AddAs_Or4Spc_ToExtNm E
FEs_AddTab2Spc_ToExtNm E
FEs_AlignFld F
PSel_Fny_Extny = KwSel & PDis(IsDis) & C_NL & Join(LyzAyab(E, F), C_CNL)
End Function

Private Sub FEs_AddTab2Spc_ToExtNm(OE$())
OE = AddPfxzAy(OE, C_T & "  ")
End Sub
Private Sub FEs_SetExtNm_ToBlnk_IfEqToFld(F$(), OE$())
Dim J%
For J = 0 To UB(OE)
    If OE(J) = F(J) Then OE(J) = ""
Next
End Sub
Private Sub FEs_SqQteExtNm_IfNB(OE$())
Dim J%
For J = 0 To UB(OE)
    If OE(J) <> "" Then
        OE(J) = QteSq(OE(J))
    End If
Next
End Sub
Private Sub FEs_AlignExtNm(OE$())
OE = AlignAy(OE)
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
OF = AlignAy(OF)
End Sub
Private Function PDis$(Dis As Boolean)
If Dis Then PDis = " " & KwDis & C_NLTT Else PDis = C_NLTT
End Function
Private Function PSel_Fny$(Fny$(), Optional Dis As Boolean)
PSel_Fny = KwSel & PDis(Dis) & C_NLTT & JnCommaSpc(Fny)
End Function

Private Sub Z_PSel_Fny_Extny()
Dim Fny$()
Dim Extny$()
GoSub Z
Exit Sub
Z:
    Fny = SyzSS("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Extny = TermAy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Debug.Print PSel_Fny_Extny(Fny, Extny)
    Return
End Sub

Private Function PSel_T$(T)
PSel_T = KwSel & C_T & "*" & PFm(T)
End Function


Private Function PSelzX$(X$, Optional Dis As Boolean)
PSelzX = KwSel & PDis(Dis) & X
End Function

Private Function PSet_FF_EqDr$(FF$, EqDr)

End Function

Private Function PSet_FF_Ey$(FF$, Ey$())
PSet_FF_Ey = PSet_Fny_Ey(SyzSS(FF), Ey)
End Function

Private Function PSet_Fny_Ey$(Fny$(), Ey$())
Dim J$(): J = LyzAyab(SyzQteSq(Fny), Ey, " = ")
Dim J1$(): J1 = AddPfxzAy(J, C_TT)
Dim S$: S = Jn(J, "," & C_NL)
PSet_Fny_Ey = C_NLT & KwSet & C_NL & S
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
Dim F$(): F = SyzQteSq(Fny)
Dim V$(): V = SqlQteVy(Vy)
PSet_Fny_Vy = JnComma(LyzAyab(F, V, "="))
End Function

Private Property Get C_Comma$()
If Fmt Then
    C_Comma = "," & vbCrLf
Else
    C_Comma = ", "
End If
End Property

Private Function PAnd_Bexp$(Bexp$)
If Bexp = "" Then Exit Function
'PAnd_Bexp = NxtLin & "and " & NxtLin_Tab & Bexp
End Function

Private Function AddPfxNLTT$(Sy$())
AddPfxNLTT = Jn(AddPfxzAy(Sy, C_NLTT), "")
End Function

Private Function Bexp_E_InLis$(Expr$, InLisStr$)
If InLisStr = "" Then Exit Function
Bexp_E_InLis = FmtQQ("? in (?)", Expr, InLisStr)
End Function

Private Sub Z_PGp_ExprVblAy()
Dim ExprVblAy$()
    Push ExprVblAy, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push ExprVblAy, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push ExprVblAy, "3sldkfjsdf"
DmpAy SplitVBar(PGp_ExprVblAy(ExprVblAy))
End Sub


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
SqlSel_FF_T = PSel_FF(FF, IsDis) & PFm(T)
End Function

Private Function PSet_FF_ExprAy$(FF, Ey$())
Const CSub$ = CMod & "PSet_FF_Ey"
Dim Fny$(): Fny = SyzSS(FF)
Ass IsVblAy(Ey)
If Si(Fny) <> Si(Ey) Then Thw CSub, "[FF-Sz} <> [Si-Ey], where [FF],[Ey]", Si(Fny), Si(Ey), FF, Ey
Dim AFny$()
    AFny = AlignAy(Fny)
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
QNm = QteSq(T)
End Function

Private Function PUpd$(T)
PUpd = KwUpd & C_T & QNm(T)
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
QV = SqlQte(V)
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

Private Function PBexp_Fny_EqVy$(Fny$(), EqVy)

End Function

Private Function PWh_Fny_EqVy$(Fny$(), EqVy)
PWh_Fny_EqVy = C_Wh & PBexp_Fny_EqVy(Fny, EqVy)
End Function

Private Function PWh$(Bexp$)
If Bexp = "" Then Exit Function
PWh = C_Wh & Bexp
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

Private Sub Z_SqlSel_Fny_Ey_Into_T_OB()
Dim Fny$(), Ey$(), Into$, T$, Bexp$
GoSub Z
Exit Sub
Z:
    Fny = SyzSS("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Ey = TermAy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Into = "#IZHT086"
    T = ">ZHT086"
    Bexp = ""
    Debug.Print SqlSel_Fny_Extny_Into_T_OB(Fny, Ey, Into, T, Bexp)
    Return
End Sub

Function SqlSel_Fny_Extny_Into_T_OB$(Fny$(), Extny$(), Into, T, Optional Bexp$)
SqlSel_Fny_Extny_Into_T_OB = PSel_Fny_Extny(Fny, Extny) & PInto_T(Into) & PFm(T) & PWh(Bexp)
End Function

Function SqlSel_Dist_Fny_EDict_Into_T_Wh_Gp_Ord$(IsDist As Boolean, Fny$(), EDic As Dictionary, T$, Wh$, Gp$, Ord$)

End Function
Function SqlSel_FF_EDic_Into_T_OB$(FF$, EDic As Dictionary, Into, T, Optional Bexp$)
Dim Fny$(): Fny = SyzSS(FF)
Dim ExprAy$(): ExprAy = SyzDicKy(EDic, Fny)
Stop
SqlSel_FF_EDic_Into_T_OB = SqlSel_Fny_Extny_Into_T_OB(Fny, ExprAy, Into, T, Bexp)
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

Function SqlQte$(V)
Dim O$, C$
C = SqlQteChr(V)
If C <> "" Then SqlQte = Qte(CStr(V), C): Exit Function
Select Case True
Case IsBool(V): O = IIf(V, "true", "false")
Case IsEmpty(V), IsNull(V), IsNothing(V): O = "null"
Case Else: O = V
End Select
SqlQte = O
End Function

Function SqlQteChrzT$(A As Dao.DataTypeEnum)
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
    Dao.DataTypeEnum.dbText: SqlQteChrzT = "'"
Case _
    Dao.DataTypeEnum.dbDate: SqlQteChrzT = "#"
Case Else
    Thw CSub, "Invalid DaoTy", "DaoTy", A
End Select
End Function

Function SqlQteChr$(V)
Dim O$
Select Case True
Case IsStr(V): O = "'"
Case IsDate(V): O = "#"
End Select
SqlQteChr = O
End Function
Function SqlQteVy(Vy) As String()
Dim V
For Each V In Vy
    PushI SqlQteVy, SqlQte(V)
Next
End Function

Function SqlUpd_T_FF_EqDr_Whff_Eqvy$(T, FF$, Dr, WhFF$, EqVy)
SqlUpd_T_FF_EqDr_Whff_Eqvy = PUpd(T) & PSet_FF_EqDr(FF, Dr) & PWh_FF_Eqvy(WhFF, EqVy)
End Function

Function SqlDrpFld$(T, Fny$())
Dim S$: S = JnCommaSpc(SyzQteSq(Fny))
SqlDrpFld = "Alter Table [" & T & "] drop column " & S
End Function

Function SqlUpdzEy$(T, Fny$(), Ey$(), Optional OBexp$)
SqlUpdzEy = PUpd(T) & PSet_Fny_Ey(Fny, Ey) & PWh(OBexp)
End Function

Private Function PWh_FF_Eqvy$(FF$, EqVy)

End Function


Function SqlSel_FF_T_Bexp$(FF$, T, Bexp$)

End Function
Private Function JnAnd$(Sy$())
JnAnd = Jn(Sy, " " & KwAnd & " ")
End Function
Private Function POnzJnXA(JnFny$())
Dim X$(): X = SyzQAy("x.[?]", JnFny)
Dim A$(): A = SyzQAy("a.[?]", JnFny)
Dim J$(): J = LyzAyab(X, A, " = ")
Dim S$: S = JnAnd(J)
POnzJnXA = KwOn & " " & S
End Function
Private Function PTblzXAJn$(TblX$, TblA$, JnFny$())
PTblzXAJn = C_TT & "[" & TblX & "] x" & C_NLTT & KwIJn & " [" & TblA & "] a " & POnzJnXA(JnFny)
End Function

Function PUpdzXAJn$(TblX$, TblA$, JnFny$())
Dim X$: X = PTblzXAJn(TblX, TblA, JnFny)
PUpdzXAJn = PUpdzX(X)
End Function
Private Function PSetzXAFny(Fny$())
PSetzXAFny = PSetzXA(Fny, Fny)
End Function

Private Function PSetzXA(FnyX$(), FnyA$())
Dim X$(): X = AddPfxSzAy(FnyX, "x.[", "]")
Dim A$(): A = AddPfxSzAy(FnyA, "a.[", "]")
Dim J$(): J = LyzAyab(X, A, " = ")
          J = AddPfxzAy(J, C_TT)
Dim S$:   S = Jn(J, "," & C_NL)
PSetzXA = PSetzX(S)
End Function

Function SqlUpdzJn$(T$, FmA$, JnFny$(), SetFny$())
'Fm T     : Table nm to be update.  It will have alias x.
'Fm FmA   : Table nm used to update @T.  It will has alias a.
'Fm JnFny : Fld nm common in @T & @FmA.  It will use to bld the jn clause with alias x and a.
'Fm SetX  : Fny in @T to be updated.  No alias, by the ret sql will put the alias x.  Sam ele as @EqA.
'Ret      : upd sql stmt updating @T from @FmA using @JnFny as jn clause setting @T fld as stated in @SetX eq to @FmA fld as stated in @EqA
Dim U$: U = PUpdzXAJn(T, FmA, JnFny)
Dim S$: S = PSetzXAFny(SetFny)
SqlUpdzJn = U & C_NL & S
End Function

Private Function PUpdzX$(TblX$)
PUpdzX = KwUpd & C_NL & TblX
End Function

Function PSetzX$(SetX$)
PSetzX = C_T & KwSet & C_NL & SetX
End Function

Function SqlUpdzXSet$(TblX$, SetX$)
SqlUpdzXSet = PUpdzX(TblX) & PSetzX(SetX)
End Function

Function SqlAddColzLis$(T, ColLis$)
SqlAddColzLis = FmtQQ("Alter Table [?] add column ?", T, ColLis)
End Function

Function SqlAddColzAy$(T, ColAy$())
SqlAddColzAy = SqlAddColzLis(T, JnCommaSpc(ColAy))
End Function

Function SqlAddCol_T_Fny_FzDiSqlTy$(T, Fny$(), FzDiSqlTy As Dictionary)
Dim O$(), F
For Each F In Fny
    PushI O, F & " " & VzDicK(FzDiSqlTy, F, "FzDiSqlTy", "Fld")
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
SqlCrtSk_T_SkFny = FmtQQ("Create unique Index SecondaryKey on [?] (?)", T, JnComma(SyzQteSq(SkFny)))
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
A = JnComma(SyzQteSqIf(Fny))
B = JnComma(SqlQteVy(Dr))
SqlIns_T_FF_Dr = FmtQQ("Insert Into [?] (?) Values(?)", T, A, B)
End Function
Function SqlSel_T$(T, Optional Bexp$)
SqlSel_T = "Select *" & PFm(T) & PWh(Bexp)
End Function

Function SqlSel_T_Wh$(T, Bexp$)
SqlSel_T_Wh = SqlSel_T(T) & PWh(Bexp)
End Function

Function SqlSel_Into_T_WhFalse(Into, T)
SqlSel_Into_T_WhFalse = FmtQQ("Select * Into [?] from [?] where false", Into, T)
End Function

Function SqlSel_F$(F$)
SqlSel_F = SqlSel_F_T(F, F)
End Function

Function SqlSel_F_T$(F$, T, Optional Bexp$)
SqlSel_F_T = FmtQQ("Select [?] from [?]?", F, T, PWh(Bexp))
End Function


Function SqlSel_FF_T_Ord(FF$, T, OrdMinusSfxFF$)
SqlSel_FF_T_Ord = PSel_FF(FF) & PFm(T) & POrd_MinusSfxFF(OrdMinusSfxFF)
End Function

Function SqlUpd_T_Sk_Fny_Dr$(T, Sk$(), Fny$(), Dr)
If Si(Sk) = 0 Then Stop
Dim PUpd$, Set_$, Wh$: GoSub X_PUpd_Set_Wh
'UpdSql = PUpd & Set_ & Wh
Exit Function
X_PUpd_Set_Wh:
    Dim Fny1$(), Dr1(), Skvy(): GoSub X_Fny1_Dr1_SkVy
    PUpd = "Update [" & T & "]"
    Set_ = PSet_Fny_Vy(Fny1, Dr1)
    Wh = PWh_Fny_EqVy(Sk, Skvy)
    Return
X_Ay:
    Dim L$(), R$()
    L = AlignQteSq(Fny)
    R = SqlQteVy(Dr)
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
    Dim L$(): L = SyzQteSq(Fny)
    Dim R$(): R = SqlQteVy(Vy)
    Dim J%, O$()
    For J = 0 To UB(L)
        Push O, L(J) & " = " & R(J)
    Next
    A = JnComma(O)
    Return
End Function

Function SqlDlt_T$(T)
SqlDlt_T = "Delete * from [" & T & "]"
End Function

Function SqlDlt$(T, Bexp$)
SqlDlt = SqlDlt_T(T) & PWh(Bexp)
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


Function SqlSel_Fny_T(Fny$(), T, Optional Bexp$, Optional IsDis As Boolean)
SqlSel_Fny_T = PSel_Fny(Fny, IsDis) & PFm(T) & PWh(Bexp)
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

Function SqpAEqB_Fny_AliasAB$(Fny$(), Optional AliasAB$ = "x a")
Dim A1$: A1 = BefSpc(AliasAB) ' Alias1
Dim A2$: A2 = BefSpc(AliasAB) ' Alias2
Dim A$(): A = AddPfxzAy(Fny, A1 & ".")
Dim B$(): B = AddPfxzAy(Fny, A2 & ".")
Dim J$(): J = LyzAyab(A, B, " = ")
SqpAEqB_Fny_AliasAB = JnCommaSpc(J)
End Function

Function SqlSelzIntoFF$(Into$, FF$, Fm$, Optional OBexp$, Optional ODis As Boolean)
SqlSelzIntoFF = PSel_FF(FF, ODis) & PInto_T(Into) & PFm(Fm) & PWh(OBexp)
End Function

Function SqlSel_Fny_T_WhFny_EqVy$(Fny$(), T, WhFny$(), EqVy)
SqlSel_Fny_T_WhFny_EqVy = SqlSel_Fny_T(Fny, T, PWh_Fny_EqVy(WhFny, EqVy))
End Function

Function SqlSel_X_Into_T_OB_OGp_OOrd_ODis$(X$, Into$, T$, Optional OBexp$, Optional OGp$, Optional OOrd$, Optional ODis As Boolean)
SqlSel_X_Into_T_OB_OGp_OOrd_ODis$ = PSelzX(X, ODis) & PInto_T(Into) & PFm(T) & PWh(OBexp) & PGp(OGp) & POrd(OOrd$)
End Function

Function SqlSel_X_Into_T_OB$(X$, Into$, T$, Optional OBexp$)
SqlSel_X_Into_T_OB$ = PSelzX(X) & PInto_T(Into) & PFm(T) & PWh(OBexp)
End Function

Private Function PGp$(Gp$)
If Gp = "" Then Exit Function
PGp = C_T & KwGp & C_T & Gp
End Function
Private Function POrd$(Ord$)

End Function
Function SqlSel_Fny_Into_T_OB$(Fny$(), Into$, T, Optional Bexp$)

End Function
Function SqlSel_X_Into_T_T_Jn$(X$, Into$, T1$, T2$, Jn$)

End Function

Function SqlSelzInto$(Into$, SelX$, Fm$, Optional Gp$, Optional Bexp$)
Dim Dis As Boolean: If Gp <> "" Then Dis = True
SqlSelzInto = PSelzX(SelX, Dis) & PFm(Fm) & PWh(Bexp) & PGp(Gp)
End Function

Function SqlSelzIntoCpy$(Into$, Fm$)
SqlSelzIntoCpy = PSelzX("*") & PInto_T(Into) & PFm(Fm)
End Function

Function SqlSelzIntoFmX$(Into$, SelX$, FmX$, Optional Gp$, Optional Bexp$)
Dim Dis As Boolean: If Gp <> "" Then Dis = True
SqlSelzIntoFmX = PSelzX(SelX, Dis) & PInto_T(Into) & PFmzX(FmX) & PWh(Bexp) & PGp(Gp)
End Function

Function SqlSel_X_T$(X$, T, Optional Bexp$)
SqlSel_X_T = PSelzX(X) & PFm(T) & PWh(Bexp)
End Function

Function SqlSel_FF_T_Ordff$(FF$, T, OrdMinusSfxFF$)
SqlSel_FF_T_Ordff = PSel_FF(FF) & PFm(T) & POrd_MinusSfxFF(OrdMinusSfxFF)
End Function

Function SqlSelCnt_T_OB$(T, Optional Bexp$)
SqlSelCnt_T_OB = "Select Count(*)" & PFm(T) & PWh(Bexp)
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

Private Function Bexp_Fny_Vy$(Fny$(), Vy())
End Function

Private Function Bexp_F_Ev$(F$, Ev)
Bexp_F_Ev = QteSq(F) & "=" & SqlQte(Ev)
End Function

Private Sub Z()

End Sub
