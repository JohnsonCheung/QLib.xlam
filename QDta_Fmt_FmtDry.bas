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
    CellgAy = "*[" & N & "]" & Ay(0)
End If
End Function
Function CellgV$(V, Optional ShwZer As Boolean, Optional MaxWdt0% = 30) ' Convert V into a string fit in a cell
Dim O$, MaxWdt%
MaxWdt = EnsBet(MaxWdt0, 1, 1000)
Select Case True
Case IsStr(V):     O = CellgStr(V, MaxWdt)
Case IsBool(V):    O = V
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

Function InsBrk(SrtedDry(), SepDr$(), BrkCCIxy0) As Variant()
Dim BrkCCIxy&(), Drv, IsBrk As Boolean, LasDrv, J&, IsNE As Boolean
BrkCCIxy = CvLngAy(BrkCCIxy0)
           If Si(BrkCCIxy) = 0 Then InsBrk = SrtedDry: Exit Function
  LasDrv = SrtedDry(0)
           PushI InsBrk, LasDrv
           For J = 1 To UB(SrtedDry)
                 Drv = SrtedDry(J)
                IsNE = Not IsEqAyzIxy(LasDrv, Drv, BrkCCIxy)
                       If IsNE Then PushI InsBrk, SepDr
                       If IsNE Then LasDrv = Drv
                       Push InsBrk, Drv
           Next
End Function

Function WdtAyzDry(Dry()) As Integer()
Dim J&
For J = 0 To NColzDry(Dry) - 1
    Push WdtAyzDry, WdtzAy(StrColzDry(Dry, J))
Next
End Function

Function FmtDr(Dr, A As DrSepr, Optional IsLin As Boolean)
If IsLin Then
    FmtDr = Quote(Jn(Dr, A.LinSep), A.LinQuote)
Else
    FmtDr = Quote(Jn(Dr, A.DtaSep), A.DtaQuote)
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

Private Sub A_Main()
FmtDryAsJnSep:
FmtDry:
End Sub
Sub BrwDry(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean, Optional Fmt As EmTblFmt = EmTblFmt.EiTblFmt)
BrwAy FmtDry(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt)
End Sub

Sub BrwDryzSpc(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean)
BrwAy FmtDry(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt:=EiSSFmt)
End Sub

Function FmtDry(Dry(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt) _
As String()
Dim Dry2()
    Dim Dry1(), W%(), Sep$()
         If Si(Dry) = 0 Then Exit Function
  Dry1 = CellgDry(Dry, ShwZer, MaxColWdt)
     W = WdtAyzDry(Dry1)
  Dry2 = AlignDryzW(Dry1, W)
   Sep = SepDr(W)
         If IsArray(BrkCCIxy0) Then Dry2 = InsBrk(Dry2, Sep, BrkCCIxy0)

Dim L$, M$(), Sepr As DrSepr
  Sepr = DrSeprzEmTblFmt(Fmt)
     M = JnDrzDry(Dry2, Sepr)
     L = FmtDr(Sep, Sepr, IsLin:=True)
FmtDry = Sy(L, M, L)
End Function

Function JnDrzDry(Dry(), A As DrSepr) As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI JnDrzDry, FmtDr(Dr, A)
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

