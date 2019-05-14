Attribute VB_Name = "QDta_Fmt_FmtDry"
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
For Each I In ResiMax(Drv, UB(WdtAy))
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

Function ShfBef$(OLin, Sep$)
ShfBef = Bef(OLin, Sep)
OLin = RmvBef(OLin, Sep)
End Function

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
Function DryFmtCommon(Dry(), MaxColWdt%, ShwZer As Boolean) As Variant()
DryFmtCommon = AlignDry(DryMkCell(Dry, ShwZer, MaxColWdt))
End Function

Function AlignDryW(Dry(), WdtAy%()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI AlignDryW, AlignzDrvW(Drv, WdtAy)
Next
End Function
Function AlignDry(Dry()) As Variant()
AlignDry = AlignDryW(Dry, WdtAyzDry(Dry))
End Function

Function DryMkCell(Dry(), ShwZer As Boolean, MaxColWdt%) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
   Push DryMkCell, DrMkCell(Dr, ShwZer, MaxColWdt)
Next
End Function

Private Function DrMkCell(Dr, ShwZer As Boolean, MaxWdt%) As String()
Dim I
For Each I In Itr(Dr)
    PushI DrMkCell, MkCell(I, ShwZer, MaxWdt)
Next
End Function

Function MkCell$(V, Optional ShwZer As Boolean, Optional MaxWdt% = 30) ' Convert V into a string fit in a cell
Dim O$
Select Case True
Case IsStr(V)
'    If HasSubStr(V, vbCr) Then Stop
    O = EscCrLf(Left(V, MaxWdt))
Case IsNumeric(V)
    If V = 0 Then
        If ShwZer Then
            O = "0"
        End If
    Else
        O = V
    End If
Case IsEmp(V):
Case IsArray(V)
    Dim N&: N = Si(V)
    If N = 0 Then
        O = "*[0]"
    Else
        O = "*[" & N & "]" & V(0)
    End If
Case IsObject(V): O = TypeName(V)
Case Else:        O = V
End Select
MkCell = O
End Function

Function IsEqAyzIxy(A, B, Ixy&()) As Boolean
Dim J%
For J = 0 To UB(Ixy)
    If A(Ixy(J)) <> B(Ixy(J)) Then Exit Function
Next
IsEqAyzIxy = True
End Function

Function InsBrk(SrtedDry(), SepDr$(), Optional BrkCCIxy0) As Variant()
If IsMissing(BrkCCIxy0) Then InsBrk = SrtedDry: Exit Function
Dim BrkCCIxy&(): BrkCCIxy = BrkCCIxy0
Dim Drv, IsBrk As Boolean, LasDrv, J&
LasDrv = SrtedDry(0)
PushI InsBrk, LasDrv
For J = 1 To UB(SrtedDry)
    Drv = SrtedDry(J)
    If Not IsEqAyzIxy(LasDrv, Drv, BrkCCIxy) Then
        PushI InsBrk, SepDr
        LasDrv = Drv
    End If
    Push InsBrk, Drv
Next
End Function

Function WdtAyzDry(Dry()) As Integer()
Dim J&
For J = 0 To NColzDry(Dry) - 1
    Push WdtAyzDry, WdtzAy(StrColzDry(Dry, J))
Next
End Function

Function SepLin(W%(), Sep$)
SepLin = SepLinzSepDr(SepDr(W), Sep)
End Function

Function SepDr(W%()) As String()
Dim I
For Each I In W
    Push SepDr, Dup("-", I + 2)
Next
End Function

Function SepLinzSepDr$(SepDr$(), Sep$)
SepLinzSepDr = "|" & Join(SepDr, Sep) & "|"
End Function

Function JnCellzDr$(Dr, Optional Sep$ = "|", Optional Pfx$ = "|", Optional Sfx$ = "|")
JnCellzDr = Pfx & Jn(Dr, Sep) & Sfx
End Function

Private Sub A_Main()
FmtDryzAsSpcSep:
FmtDry:
End Sub
Sub BrwDry(A(), Optional MaxColWdt% = 100, Optional BrkCC, Optional ShwZer As Boolean)
BrwAy FmtDry(A, MaxColWdt, BrkCC, ShwZer)
End Sub
Sub BrwDryzSpc(A(), Optional MaxColWdt% = 100, Optional ShwZer As Boolean)
BrwAy FmtDryzAsSpcSep(A, MaxColWdt, ShwZer)
End Sub

Function FmtDry(Dry(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean) _
As String()
If Si(Dry) = 0 Then Exit Function
Dim Dry1(): Dry1 = DryMkCell(Dry, ShwZer, MaxColWdt)
Dim W%(): W = WdtAyzDry(Dry1)
Dim Dry2(): Dry2 = AlignDryW(Dry1, W)
Dim Sep$(): Sep = SepDr(W)
Dim Dry3(): Dry3 = InsBrk(Dry2, Sep, BrkCCIxy0)
Dim SepLin: SepLin = "|" & Jn(Sep, "|") & "|"
FmtDry = Sy(SepLin, FmtDryzAsSpcSep(Dry3), SepLin)
End Function

Sub DmpDryzSpcSep(Dry())
D FmtDryzAsSpcSep(Dry)
End Sub
Sub DmpDry(Dry())
D FmtDry(Dry)
End Sub

Function FmtDryzAsSpcSep(Dry(), _
Optional MaxColWdt% = 100, _
Optional ShwZer As Boolean) As String()
If Si(Dry) = 0 Then Exit Function
Dim Dr
For Each Dr In DryFmtCommon(Dry, MaxColWdt, ShwZer)
    PushI FmtDryzAsSpcSep, JnSpc(Dr)
Next
End Function



Function JnCell(Dry(), Optional Sep$ = "|", Optional Pfx$ = "|", Optional Sfx$ = "|") As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI JnCell, JnCellzDr(Dr, Sep, Pfx, Sfx)
Next
End Function

