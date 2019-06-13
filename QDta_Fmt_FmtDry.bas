Attribute VB_Name = "QDta_Fmt_FmtDry"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Fmt_Dry_Fun."
Private Const Asm$ = "QDta"
Type Syay
    Syay() As Variant ' Each element is Sy
End Type

Function AlignQuoteSq(Fny$()) As String()
AlignQuoteSq = AlignLzAy(QuoteSqzAy(Fny))
End Function

Function AlignLzAy(Sy$(), Optional W0%) As String()
Dim W%
If W0 <= 0 Then W = WdtzAy(Sy) Else W = W0
Dim I, S$
For Each I In Itr(Sy)
    S = I
    PushI AlignLzAy, AlignL(S, W)
Next
End Function

Function AlignRzAy(Ay) As String() 'Fmt-Dr-ToWdt
Dim W%: W = WdtzAy(Ay)
Dim I
For Each I In Itr(Ay)
    PushI AlignRzAy, AlignR(I, W)
Next
End Function

Function AlignzDrvW(Drv, WdtAy%()) As String() 'Fmt-Dr-ToWdt
Dim J%, W%, Cell$, I
For Each I In ResiMax(Drv, WdtAy)
    Cell = I
    W = WdtAy(J)
    PushI AlignzDrvW, AlignL(Cell, W)
    J = J + 1
Next
End Function

Function DryzSyzBySepSy(Sy$(), SepSy$()) As Variant()
Dim I, Lin
For Each I In Itr(Sy)
    Lin = I
    PushI DryzSyzBySepSy, BrkLin(Lin, SepSy)
Next
End Function

Function SslSyzDry(Dry()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    Push SslSyzDry, SslzDrv(Drv) ' Fmtss(X)
Next
End Function

Private Sub Z_BrkLin()
Dim Lin, SepSy$()
SepSy = SyzSS(". . . . . .")
Lin = "AStkShpCst_Rpt.OupFx.Fun."
Ept = Sy("AStkShpCst_Rpt", ".OupFx", ".Fun", ".")
GoSub Tst
Tst:
    Act = BrkLin(Lin, SepSy)
    C
End Sub

Function BrkLin(Lin, SepSy$()) As String()
Dim L$: L = Lin
Dim I, Sep$
For Each I In SepSy
    Sep = I
    PushI BrkLin, ShfBef(L, Sep)
Next
PushI BrkLin, L
End Function
Function JnDry(Dry()) As String()
Dim Drv
For Each Drv In Itr(Dry)
    PushI JnDry, Join(Drv)
Next
End Function

Function JnAlignDry(Dry()) As String()
JnAlignDry = JnDry(AlignDry(Dry))
End Function

Function AlignzBySepss(Sy$(), Sepss$) As String()
AlignzBySepss = JnAlignDry(DryzSyzBySepSy(Sy, SyzSS(Sepss)))
End Function

'=======================
Function AlignDryzW(Dry(), WdtAy%()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI AlignDryzW, AlignzDrvW(Drv, WdtAy)
Next
End Function
Function AlignDry(Dry()) As Variant()
AlignDry = AlignDryzW(Dry, WdtAyzDry(Dry))
End Function

Private Function CellgDry(Dry(), ShwZer As Boolean, MaxColWdt%) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
   Push CellgDry, CellgDr(Dr, ShwZer, MaxColWdt)
Next
End Function

Private Function CellgDr(Dr, ShwZer As Boolean, MaxWdt%) As String()
Dim V
For Each V In Itr(Dr)
    PushI CellgDr, CellgV(V, ShwZer, MaxWdt)
Next
End Function
Private Function CellgNum$(N, MaxW%, ShwZer As Boolean)
If N = 0 Then
    If ShwZer Then
        CellgNum = "0"
    End If
Else
    CellgNum = N
End If
End Function
Private Function CellgStr$(S, MaxW%)
CellgStr = SlashCrLf(Left(S, MaxW))
End Function
Private Function CellgAy$(Ay, MaxW%)
Dim N&: N = Si(Ay)
If N = 0 Then
    CellgAy = "*[0]"
Else
    CellgAy = "*[" & N & "]"
End If
End Function
Function CellgV$(V, Optional ShwZer As Boolean, Optional MaxWdt0% = 30) ' Convert V into a string fit in a cell
Dim O$, MaxWdt%
MaxWdt = EnsBet(MaxWdt0, 1, 1000)
Select Case True
Case IsStr(V):     O = CellgStr(V, MaxWdt)
Case IsBool(V):    O = IIf(V, "True", "")
Case IsNumeric(V): O = CellgNum(V, MaxWdt, ShwZer)
Case IsEmp(V):     O = "#Emp#"
Case IsNull(V):    O = "#Null#"
Case IsArray(V):   O = CellgAy(V, MaxWdt)
Case IsObject(V):  O = "#O:" & TypeName(V)
Case IsErObj(V)
Case Else:         O = V
End Select
CellgV = O
End Function

Function IsEqAyzIxy(A, B, Ixy&()) As Boolean
Dim J%
For J = 0 To UB(Ixy)
    If A(Ixy(J)) <> B(Ixy(J)) Then Exit Function
Next
IsEqAyzIxy = True
End Function

Function A_Ins(Nobrk As Boolean, IsBrk() As Boolean, JnD$(), SepLin$) As String()
If Nobrk Then A_Ins = JnD: Exit Function
Dim L, J&: For Each L In JnD
    If IsBrk(J) Then PushI A_Ins, SepLin
    PushI A_Ins, L
    J = J + 1
Next
End Function

Function WdtAyzDry(Dry()) As Integer()
Dim J&
For J = 0 To NColzDry(Dry) - 1
    Push WdtAyzDry, WdtzAy(StrColzDry(Dry, J))
Next
End Function

Function JnDr$(Dr, Sep$, QuoteStr$)
JnDr = Quote(Jn(Dr, Sep), QuoteStr)
End Function
Function FmtDr(Dr, A As DrSepr, Optional IsLin As Boolean)
If IsLin Then
    FmtDr = JnDr(Dr, A.LinSep, A.LinQuote)
Else
    FmtDr = JnDr(Dr, A.DtaSep, A.DtaQuote)
End If
End Function

Function SepLin$(W%())
Dim A As DrSepr
     A = DrSeprzEmTblFmt(EiTblFmt)
SepLin = FmtDr(SepDr(W), A, IsLin:=True)
End Function

Function SepDr(W%()) As String()
Dim I
For Each I In W
    Push SepDr, Dup("-", I)
Next
End Function

Sub BrwDry(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean, Optional Fmt As EmTblFmt = EmTblFmt.EiTblFmt)
BrwAy FmtDry(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt)
End Sub

Sub BrwDryzSpc(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean)
BrwAy FmtDry(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt:=EiSSFmt)
End Sub
Private Function A_IsBrk(Nobrk As Boolean, Dry(), Ixy&()) As Boolean()
If Nobrk Then Exit Function
Dim LasK, CurK, Dr
LasK = AywIxy(Dry(0), Ixy)
For Each Dr In Itr(Dry)
    CurK = AywIxy(Dr, Ixy)
    PushI A_IsBrk, Not IsEqAy(CurK, LasK)
    LasK = CurK
Next
End Function
Function FmtDry(Dry(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt) _
As String()
If Si(Dry) = 0 Then Exit Function
Dim StrD():     StrD = CellgDry(Dry, ShwZer, MaxColWdt)
Dim W%():          W = WdtAyzDry(StrD)
Dim AlignD(): AlignD = AlignDryzW(StrD, W)

Dim Ixy&():               Ixy = CvLngAy(BrkCCIxy0)
Dim Nobrk As Boolean:   Nobrk = Si(Ixy) = 0
Dim IsBrk() As Boolean: IsBrk = A_IsBrk(Nobrk, Dry, Ixy)

Dim S$:       S = IIf(Fmt = EiSSFmt, " ", " | ")
Dim Q$:       Q = IIf(Fmt = EiSSFmt, "", "| * |")
Dim JnD$(): JnD = JnDrzDry(AlignD, S, Q)
Dim Sep$(): Sep = SepDr(W)
              S = IIf(Fmt = EiSSFmt, " ", "-|-")
              Q = IIf(Fmt = EiSSFmt, "", "|-*-|")
Dim L$:       L = JnDr(Sep, S, Q)
Dim Ins$(): Ins = A_Ins(Nobrk, IsBrk, JnD, L)
         FmtDry = Sy(L, Ins, L)
End Function

Function JnDrzDry(Dry(), Sep$, QuoteStr$) As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI JnDrzDry, JnDr(Dr, Sep, QuoteStr)
Next
End Function

Sub DmpDryzSpc(Dry(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean)
D FmtDry(Dry, MaxColWdt, BrkCCIxy0, ShwZer, Fmt:=EiSSFmt)
End Sub

Sub DmpDry(Dry(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt)
D FmtDry(Dry, MaxColWdt, BrkCCIxy0, ShwZer, Fmt)
End Sub

