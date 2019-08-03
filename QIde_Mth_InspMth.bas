Attribute VB_Name = "QIde_Mth_InspMth"
Option Explicit
Option Compare Text
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

Private Sub Z_InspStmtzMthLin()
Dim A As Drs: A = DoMthzM(CMd)
Dim B$(): B = StrCol(A, "MthLin")
Dim L, ODy()
For Each L In B
    PushI ODy, Array(L, InspStmtzMthLin(L, "Md"))
Next
Dim C As Drs: C = DrszFF("MthLin InspStmt", ODy)
Brw LinzDrsR(C)
End Sub

Function InspStmtzMthLin$(MthLin, Mdn$)
With MthLinRec(MthLin)
    If .Pm = "" And Not .IsRetVal Then Exit Function
    Dim NN$: NN = JnSpc(ArgNyzPm(.Pm))
    Dim Ee$: Ee = InspExprLiszPm(.Pm)
    Dim IsN0$: IsN0 = XIsN0(.IsRetVal, .Nm)  '#Insp-Nm-0.
    Dim IsE0$: IsE0 = XIsE0(.IsRetVal, .Nm, .TyChr, .RetTy) '#Insp-Expr-0
    NN = IsN0 & NN
    Ee = IsE0 & Ee
    InspStmtzMthLin = InspStmt(NN, Ee, Mdn, .Nm)
End With
End Function

Function InspStmt$(NN$, ExprLis$, Mdn$, Mthn$)
Const C$ = "Insp ""?.?"", ""Inspect"", ""?"", ?"
InspStmt = FmtQQ(C, Mdn, Mthn, NN, ExprLis)
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
Case "DoLTDH": O = FmtQQ("FmtDrs(?.D)", V)
Case "Drs": O = FmtQQ("FmtDrs(?)", V)
Case "S12s": O = FmtQQ("FmtS12s(?)", V)
Case "CodeModule": O = FmtQQ("Mdn(?)", V)
Case "", "$", "$()", "#", "@", "%", "&", "%()", "&()", "#()", "@()", "$()": O = V
Case "Dictionary", "Byte", "Boolean", "String", "Integer": O = V
Case "", "String()", "Integer()", "Long()", "Byte()": O = V
Case Else: O = """NoFmtr(" & S & ")"""
End Select
InspExprzDclSfx = O
Exit Function
X: InspExprzDclSfx = FmtQQ(Q, V)
End Function
Function InspExprLis$(PP$, VSfx As Dictionary)
InspExprLis = Join(InspExprs(PP, VSfx), ", ")
End Function

Private Function InspExprs(PP$, VSfx As Dictionary) As String()
Dim V
For Each V In Itr(SyzSS(PP))
    PushI InspExprs, InspExpr(V, VSfx)
Next
End Function

Private Function XIsN0$(IsRetVal As Boolean, Mthn$)
If Not IsRetVal Then Exit Function
XIsN0 = "Oup(" & Mthn & ") "
End Function

Private Function XIsE0$(IsRetVal As Boolean, V, TyChr$, RetTy$)
If Not IsRetVal Then Exit Function
XIsE0 = InspExprzDclSfx(V, TyChr & RetTy) & ", "
End Function


