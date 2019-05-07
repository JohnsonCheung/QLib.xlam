Attribute VB_Name = "QDta_Fmt_Dry_Fun"
Option Explicit
Private Const CMod$ = "MDta_Fmt_Dry_Fun."
Private Const Asm$ = "QDta"
Type Syay
    Syay() As Variant ' Each element is Sy
End Type
Function AlignLzSy(Sy$(), Optional W0%) As String()
Dim W%
If W0 <= 0 Then W = WdtzSy(Sy) Else W = W0
Dim I, S$
For Each I In Itr(Sy)
    S = I
    PushI AlignLzSy, AlignL(S, W)
Next
End Function

Function AlignRzSy(Sy$()) As String() 'Fmt-Dr-ToWdt
Dim W%: W = WdtzSy(Sy)
Dim I
For Each I In Itr(Sy)
    PushI AlignRzSy, AlignR(CStr(I), W)
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
Dim I, Lin$
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
Dim Lin$, SepSy$()
SepSy = SyzSsLin(". . . . . .")
Lin = "AStkShpCst_Rpt.OupFx.Fun."
Ept = Sy("AStkShpCst_Rpt", ".OupFx", ".Fun", ".")
GoSub Tst
Tst:
    Act = BrkLin(Lin, SepSy)
    C
End Sub

Function ShfBef$(OLin$, Sep$)
ShfBef = Bef(OLin, Sep)
OLin = RmvBef(OLin, Sep)
End Function

Function BrkLin(Lin$, SepSy$()) As String()
Dim L$: L = Lin
Dim I, Sep$
For Each I In SepSy
    Sep = I
    PushI BrkLin, ShftBef(L, Sep)
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
AlignzBySepss = JnAlignDry(DryzSySepss(Sy, Sepss))
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

Function IsEqAyzIxAy(A, B, IxAy&()) As Boolean
Dim J%
For J = 0 To UB(IxAy)
    If A(IxAy(J)) <> B(IxAy(J)) Then Exit Function
Next
IsEqAyzIxAy = True
End Function

Function InsBrk(SrtedDry(), SepDr$(), Optional BrkCCIxAy0) As Variant()
If IsMissing(BrkCCIxAy0) Then InsBrk = SrtedDry: Exit Function
Dim BrkCCIxAy&(): BrkCCIxAy = BrkCCIxAy0
Dim Drv, IsBrk As Boolean, LasDrv, J&
LasDrv = SrtedDry(0)
PushI InsBrk, LasDrv
For J = 1 To UB(SrtedDry)
    Drv = SrtedDry(J)
    If Not IsEqAyzIxAy(LasDrv, Drv, BrkCCIxAy) Then
        PushI InsBrk, SepDr
        LasDrv = Drv
    End If
    Push InsBrk, Drv
Next
End Function

Function WdtAyzDry(Dry()) As Integer()
Dim J&
For J = 0 To NColzDry(Dry) - 1
    Push WdtAyzDry, WdtzSy(StrColzDry(Dry, J))
Next
End Function

Function SepLin$(W%(), Sep$)
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

Function LinFmDrByJnCell$(Dr, Optional Sep$ = "|", Optional Pfx$ = "|", Optional Sfx$ = "|")
LinFmDrByJnCell = Pfx & Jn(Dr, Sep) & Sfx
End Function

Function FmtDryByJnCell(Dry(), Optional Sep$ = "|", Optional Pfx$ = "|", Optional Sfx$ = "|") As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI FmtDryByJnCell, LinFmDrByJnCell(Dr, Sep, Pfx, Sfx)
Next
End Function
