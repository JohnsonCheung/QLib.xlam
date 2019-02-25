Attribute VB_Name = "MTp_SqyRslt_4_SqyRsltzGpAy"
Option Explicit
Type SqlRslt: Er() As String: Sql As String: End Type
Const Msg_Sq_1_NotInEDic = "These items not found in ExprDic [?]"
Const Msg_Sq_1_MustBe1or0 = "For %?xxx, 2nd term must be 1 or 0"

Private Enum eStmtTy
    eUpdStmt = 1
    eDrpStmt = 2
    eSelStmt = 3
End Enum
Const U_Into$ = "INTO"
Const U_Sel$ = "SEL"
Const U_SelDis$ = "SELECT DISTINCT"
Const U_Fm$ = "FM"
Const U_Gp$ = "GP"
Const U_Wh$ = "WH"
Const U_And$ = "AND"
Const U_Jn$ = "JN"
Const U_LeftJn$ = "LEFT JOIN"
Private Pm As Dictionary
Private StmtSw As Dictionary
Private FldSw As Dictionary
Private Function SqlRsltzEr(Sql$, Er$()) As SqlRslt
SqlRsltzEr.Er = Er
SqlRsltzEr.Sql = Sql
End Function

Private Sub AAMain()
Z_SqyRslt
End Sub
Function SqyRsltzGpAy(SqGpAy() As Gp, PmDic As Dictionary, StmtSwDic As Dictionary, FldSwDic As Dictionary) As SqyRslt
Set Pm = Pm
Set StmtSw = StmtSwDic
Set FldSw = FldSwDic

Dim I
Dim R As LyRslt
Dim O As SqyRslt
For Each I In SqGpAy
    R = SqLyRslt(CvGp(I))
    O = PushSqlRslt(O, SqlRsltzSqLy(R.Ly))
Next
End Function

Private Function SqlRsltzSqLy(SqLy$()) As SqlRslt
Dim Ty As eStmtTy
    Ty = StmtTy(SqLy)

Dim IsSkip As Boolean
    IsSkip = StmtSw.Exists(StmtSwKey(SqLy, Ty))
    If IsSkip Then Exit Function

Dim NoExprLinSqLy$(), E As Dictionary
Set E = ExprDic(SqLy)
NoExprLinSqLy = RmvExprLin(SqLy)
Dim O As SqlRslt
    Select Case Ty
    Case eUpdStmt: O = SqlRsltUpd(NoExprLinSqLy, E)
    Case eDrpStmt: O = SqlRsltDrp(NoExprLinSqLy)
    Case eSelStmt: O = SqlRsltSel(NoExprLinSqLy, E)
    Case Else: Stop
    End Select
SqlRsltzSqLy = O
End Function

Private Function SqlRsltDrp(SqLy$()) As SqlRslt
End Function

Private Function SqlRsltSel(SqLy$(), ExprDic As Dictionary) As SqlRslt
Dim Er$()
Dim Sel$(), Into$(), Fm$(), Jn$(), Wh$(), Gp$()
#If False Then
    Dim E As Dictionary
    Set E = ExprDic
    Dim I, J%, B$(), L$
    B = AyReverseI(A)
    PushI O, XSel(Pop(B), E)
'    PushI O, QSqpInto_T(RmvT1(Pop(B)))
'    PushI O, SqpFm(RmvT1(Pop(B)))
    PushIAy O, XJnOrLeftJn(PopJnOrLeftJn(B), E)
    L = PopWh(B)
    If L <> "" Then
        PushI O, XWh(L, E)
        PushIAy O, XAnd(PopAnd(B), E)
    End If
    PushI O, XGp(PopGp(B), E)
#End If
Dim O$(): O = AyAddAp(Sel, Into, Fm, Jn, Wh, Gp)
SqlRsltSel = SqlRsltzEr(JnCrLf(O), Er)
End Function
Private Function RmvExprLin(SqLy$()) As String()

End Function

Private Function SqlRsltUpd(A$(), E As Dictionary) As SqlRslt

End Function

Private Function Fny_WhActive(A$()) As String()
Dim F
For Each F In A
    If FldSw.Exists(F) Then PushI Fny_WhActive, F
Next
End Function

Private Function ExprDic(A) As Dictionary
Dim Expr$(), M As AyAB
'M = AyBrk_BY_ELE(A, "$")
'Set ExprDic = LyDic(CvSy(M.B))
End Function


Private Function StmtTy(SqLy$()) As eStmtTy
Dim L$
L = UCase(RmvPfx(T1(SqLy(0)), "?"))
Select Case L
Case "SEL": StmtTy = eSelStmt
Case "UPD": StmtTy = eUpdStmt
Case "DRP": StmtTy = eDrpStmt
Case Else: Stop
End Select
End Function

Private Function StmtSwKey$(SqLy$(), Ty As eStmtTy)
Stop
Select Case Ty
Case eStmtTy.eSelStmt: StmtSwKey = StmtSwKey_SEL(SqLy)
Case eStmtTy.eUpdStmt: StmtSwKey = StmtSwKey_UPD(SqLy)
Case Else: Stop
End Select
End Function

Private Function StmtSwKey_SEL$(SqLy$())
StmtSwKey_SEL = FstEleRmvT1(SqLy, "FM")
End Function

Private Function StmtSwKey_UPD$(SqLy)
Dim Lin1$
    Lin1 = SqLy(0)
If RmvPfx(ShfTerm(Lin1), "?") <> "upd" Then Stop
StmtSwKey_UPD = Lin1
End Function

Private Function FndVy(K, E As Dictionary, OVy$(), OQ$)
'Return true if not found
End Function

Private Function FndValPair(K, E As Dictionary, OV1, OV2)
'Return true if not found
End Function

Private Function IsXXX(A$(), XXX$) As Boolean
IsXXX = UCase(T1(A(UB(A)))) = XXX
End Function


Private Property Get SampExprDic() As Dictionary
Dim O$()
PushI O, "A XX"
PushI O, "B BB"
PushI O, "C DD"
PushI O, "E FF"
'Set SampExprDic = LyDic(O)
End Property

Private Property Get SampSqLnxAy() As Lnx()
Dim O$()
PushI O, "sel ?MbrCnt RecCnt TxCnt Qty Amt"
PushI O, "into #Cnt"
PushI O, "fm   #Tx"
PushI O, "wh   RecCnt bet @XX @XX"
PushI O, "and  RecCnt bet @XX @XX"

PushI O, "$"
PushI O, "?MbrCnt ?Count(Distinct Mbr)"
PushI O, "RecCnt  Count(*)"
PushI O, "TxCnt   Sum(TxCnt)"
PushI O, "Qty     Sum(Qty)"
PushI O, "Amt     Sum(Amt)"
SampSqLnxAy = LnxAy(O)
End Property
Private Function SqLyRslt(A As Gp) As LyRslt

End Function

Private Function PushSqlRslt(A As SqyRslt, B As SqlRslt) As SqyRslt
Dim O As SqyRslt
O = A

End Function
Private Function MsgAndLinOp_ShouldBe_BetOrIn(A)

End Function
Private Function XAnd(A$(), E As Dictionary)
'and f bet xx xx
'and f in xx
Dim F$, I, L$, Ix%, M As Lnx
For Each I In Itr(A)
    Set M = I
    LnxAsg M, L, Ix
    If ShfTerm(L) <> "and" Then Stop
    F = ShfTerm(L)
    Select Case ShfTerm(L)
    Case "bet":
    Case "in"
    Case Else: Stop
    End Select
Next
End Function

Private Function XGp$(L$, E As Dictionary)
If L = "" Then Exit Function
Dim ExprAy$(), Ay$()
Stop
'    ExprAy = DicSelIntoSy(EDic, Ay)
'XGp = SqpGp(ExprAy)
End Function

Private Function XJnOrLeftJn(A$(), E As Dictionary) As String()

End Function

Private Function PopJnOrLeftJn(A$()) As String()
PopJnOrLeftJn = PopMulXorYOpt(A, U_Jn, U_LeftJn)
End Function

Private Function PopXXXOpt$(A$(), XXX$)
'Return No-Q-T1-of-LasEle-of-A$() if No-Q-T1-of it = XXX else return ''
If Sz(A) = 0 Then Exit Function
PopXXXOpt = PopXXX(A, XXX)
End Function

Private Function PopXXX$(A$(), XXX$)
'Return No-Q-T1-of-LasEle-of-A$() if No-Q-T1-of it = XXX else Stop
Dim L$: L = A(UB(A))
If RmvPfx(T1(L), "?") = XXX Then
    PopXXX = RmvT1(L)
    Pop A
End If
End Function

Private Function PopGp$(A$())
PopGp = PopXXXOpt(A, U_Gp)
End Function

Private Function PopWh$(A$())
PopWh = PopXXXOpt(A, U_Wh)
End Function

Private Function PopAnd(A$()) As String()
PopAnd = PopMulXXX(A, U_And)
End Function

Private Function PopXorYOpt$(A$(), X$, Y$)
Dim L$
L = PopXXXOpt(A, X): If L <> "" Then PopXorYOpt = L: Exit Function
PopXorYOpt = PopXXXOpt(A, Y)
End Function

Private Function PopMulXorYOpt(A$(), X$, Y$) As String()
Dim J%, L$
While Sz(A) > 0
    J = J + 1: If J > 1000 Then Stop
    L = PopXorYOpt(A, X, Y)
    If L = "" Then Exit Function
    PushI PopMulXorYOpt, L
Wend
End Function

Private Function PopMulXXX(A$(), XXX$) As String()
Dim J%
While Sz(A) > 0
    J = J + 1: If J > 1000 Then Stop
    If Not IsXXX(A, XXX) Then Exit Function
    PushObj PopMulXXX, Pop(A)
Wend
End Function

Private Function XSel$(A$, E As Dictionary)
Dim Fny$()
    Dim T1$, L$
    L = A
    T1 = RmvPfx(ShfTerm(L), "?")
    Fny = XSelFny(SySsl(L), FldSw)
Select Case T1
'Case U_Sel:    XSel = X.Sel_Fny_EDic(Fny, E)
'Case U_SelDis: XSel = X.Sel_Fny_EDic(Fny, E, IsDis:=True)
Case Else: Stop
End Select
End Function
Private Function XSelFny(Fny$(), FldSw As Dictionary) As String()
Dim F
For Each F In Fny
    If FstChr(F) = "?" Then
        If Not FldSw.Exists(F) Then Stop
        If FldSw(F) Then PushI XSelFny, F
    Else
        PushI XSelFny, F
    End If
Next
End Function

Private Function XSet(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XUpd(A As Lnx, E As Dictionary, OEr$())

End Function
Private Function XWh$(L$, E As Dictionary)
'L is following
'  ?Fld in @ValLis  -
'  ?Fld bet @V1 @V2
Dim F$, Vy$(), V1, V2, IsBet As Boolean
If IsBet Then
    If Not FndValPair(F, E, V1, V2) Then Exit Function
    'XWh = SqpWhBet(F, V1, V2)
    Exit Function
End If
'If Not FndVy(F, E, Vy, Q) Then Exit Function
'XWh = SqpWhFldInVy_Str(F, Vy)
End Function

Private Function XWhBetNbr$(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XWhExpr(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XWhInNbrLis$(A As Lnx, E As Dictionary, OEr$())

End Function

Private Sub Z_SqlRsltSel()
Dim E As Dictionary, Ly$(), Act As SqlRslt
'---
Erase Ly
    Push Ly, "?XX Fld-XX"
    Push Ly, "BB Fld-BB-LINE-1"
'    Push Ly, "BB Fld-BB-LINE-2"
'    Set E = LyDic(Ly)           '<== Set ExprDic
Erase Ly
    Set FldSw = New Dictionary
    FldSw.Add "?XX", False       '<=== Set FldSw
Erase Ly
    Erase Ly
    PushI Ly, "sel ?XX BB CC"
    PushI Ly, "into #AA"
    PushI Ly, "fm   #AA"
    PushI Ly, "jn   #AA"
    PushI Ly, "jn   #AA"
    PushI Ly, "wh   A bet $a $b"
    PushI Ly, "and  B in $c"
    PushI Ly, "gp   D C"        '<== LySq
GoSub Tst
Exit Sub
Tst:
    Act = SqlRsltSel(Ly, E)
    C
    Return
End Sub

Private Sub Z_ExprDic()
Dim Ly$()
Dim D As New Dictionary
'-----

Erase Ly
PushI Ly, "aaa bbb"
PushI Ly, "111 222"
PushI Ly, "$"
PushI Ly, "A B0"
PushI Ly, "A B1"
PushI Ly, "A B2"
PushI Ly, "B B0"
D.RemoveAll
    D.Add "A", JnCrLf(SySsl("B0 B1 B2"))
    D.Add "B", "B0"
    Set Ept = D
GoSub Tst
Exit Sub
Tst:
    Set Act = ExprDic(Ly)
    Ass IsEqDic(CvDic(Act), CvDic(Ept))
    
    Return
End Sub

Private Sub Z_SqyRslt()
Dim A() As Gp, Pm As Dictionary, StmtSw As Dictionary, FldSw As Dictionary, Er$()
GoSub Dta1
GoSub Tst
Return
Tst:
    Dim Act As SqyRslt
    Act = SqyRsltzGpAy(A, Pm, StmtSw, FldSw)
    C
    Return
Dta1:
    Return
End Sub

Private Sub Z_Sel()
Dim A$, E As Dictionary
A = "dsklfj"
Set E = SampExprDic
GoSub Tst
Exit Sub
Tst:
    Act = XSel(A, E)
    C
    Return
End Sub

Private Sub Z_StmtSwKey()
Dim Ly$(), Ty As eStmtTy
'---
PushI Ly, "sel sdflk"
PushI Ly, "fm AA BB"
Ept = "AA BB"
Ty = eSelStmt
GoSub Tst
'---
Erase Ly
PushI Ly, "?upd XX BB"
PushI Ly, "fm dsklf dsfl"
Ept = "XX BB"
Ty = eUpdStmt
GoSub Tst
Exit Sub
Tst:
    Act = StmtSwKey(Ly, Ty)
    C
    Return
End Sub


Private Sub Z()
Z_SqlRsltSel
Z_ExprDic
Z_SqyRslt
Z_StmtSwKey
Z_Sel
MTp_Sq_Sq:
End Sub



