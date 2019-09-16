Attribute VB_Name = "MxInspStmt"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxInspStmt."
Private Function InspExprLiszPm$(Pm$)
If Pm = "" Then Exit Function
Dim Ay$(): Ay = Split(Pm, ", ")
Dim O$(), P
For Each P In Ay
    Dim L$: L = P
    Dim V$: V = ShfNm(L)
    Dim S$: S = L
    PushI O, InspExprzDclSfx(V, S)
Next
InspExprLiszPm = JnCommaSpc(O)
End Function

Private Sub Z_InspStmtzL()
Dim A As Drs: A = DoMthzM(CMd)
Dim B$(): B = StrCol(A, "MthLin")
Dim L, ODy()
For Each L In B
    PushI ODy, Array(L, InspStmtzL(L, "Md"))
Next
Dim C As Drs: C = DrszFF("MthLin InspStmt", ODy)
Brw LinzDrsR(C)
End Sub

Function InspStmtzL$(MthLin, Mdn$)
With MthLinRec(MthLin)
    If .Pm = "" And Not .IsRetVal Then Exit Function
    Dim Nn$: Nn = JnSpc(ArgNyzPm(.Pm))
    Dim Ee$: Ee = InspExprLiszPm(.Pm)
    Dim IsN0$: IsN0 = XIsN0(.IsRetVal, .NM)  '#Insp-Nm-0.
    Dim IsE0$: IsE0 = XIsE0(.IsRetVal, .NM, .TyChr, .RetTy) '#Insp-Expr-0
    Nn = IsN0 & Nn
    Ee = IsE0 & Ee
    InspStmtzL = InspStmt(Nn, Ee, Mdn, .NM)
End With
End Function

Function InspStmt$(Varnn$, ExprLis$, Mdn$, Mthn$)
Const C$ = "Insp ""?.?"", ""Inspect"", ""?"", ?"
InspStmt = FmtQQ(C, Mdn, Mthn, Varnn, ExprLis)
End Function

Private Function InspExpr$(V, VSfx As Dictionary)
If Not VSfx.Exists(V) Then
    InspExpr = FmtQQ("""V(?)-NFnd""", V)
    Exit Function
End If
InspExpr = InspExprzDclSfx(V, VSfx(V))
End Function

Private Function InspExprzDclSfx$(V, DclSfx$)
Dim O$, S$
S = RmvPfx(DclSfx, " As ")
Select Case S
Case "DoLTDH": O = FmtQQ("FmtCellDrs(?.D)", V)
Case "Drs":    O = FmtQQ("FmtCellDrs(?)", V)
Case "S12s":   O = FmtQQ("FmtS12s(?)", V)
Case "CodeModule": O = FmtQQ("Mdn(?)", V)
Case "Dictionary", "Byte", "Boolean", "String", "Integer": O = V
Case "", "String()", "Integer()", "Long()", "Byte()":      O = V
Case "", "$", "$()", "#", "@", "%", "&", "%()", "&()", "#()", "@()", "$()": O = V
Case Else: O = """NoFmtr(" & S & ")"""
End Select
InspExprzDclSfx = O
End Function

Function InspExprLis$(Varnn$, DiVarqDclSfx As Dictionary)
Dim O$()
    Dim V: For Each V In Itr(SyzSS(Varnn))
        PushI O, InspExpr(V, DiVarqDclSfx)
    Next
InspExprLis = JnCommaSpc(O)
End Function

Private Function XIsN0$(IsRetVal As Boolean, Mthn$)
If Not IsRetVal Then Exit Function
XIsN0 = "Oup(" & Mthn & ") "
End Function

Private Function XIsE0$(IsRetVal As Boolean, V, TyChr$, RetTy$)
If Not IsRetVal Then Exit Function
XIsE0 = InspExprzDclSfx(V, TyChr & RetTy) & ", "
End Function

Function InspStmtzDi$(Varnn$, DiVarnnqDclSfx As Dictionary, Mdn$, Mthn$)
Dim E$: E = InspExprLis(Varnn, DiVarnnqDclSfx)
InspStmtzDi = InspStmt(Varnn, E, Mdn, Mthn)
End Function