Attribute VB_Name = "MDta_Fmt_Dry_Fun"
Option Explicit
Type Syay
    Syay() As Variant ' Each element is Sy
End Type

Function AyExtend(Ay, ToSamSzAsThisAy)
Dim Sz&, SzNew&
Sz = Si(Ay)
SzNew = Si(ToSamSzAsThisAy)
Select Case True
Case Sz > SzNew: Thw CSub, "Ay-Sz cannot > The-Given-New-AySz", "Ay-Sz NewAySz", Sz, SzNew
Case Sz < SzNew
    Dim O: O = Ay
    ReDim Preserve O(SzNew - 1)
    AyExtend = O
Case Else
    AyExtend = Ay
End Select
End Function

Function AyFmtToWdtAy(Dr, ToWdt%()) As String() 'Fmt-Dr-ToWdt
Dim J%, W%, Cell
For Each Cell In AyExtend(Dr, ToWdt)
    W = ToWdt(J)
    PushI AyFmtToWdtAy, AlignL(Cell(J), W)
    J = J + 1
Next
End Function

Function DryzSySepss(Sy$(), Sepss$) As Variant()
Dim I, Lin$, SepAy$()
SepAy = TermAy(Sepss)
For Each I In Itr(Sy)
    Lin = I
    PushI DryzSySepss, SyzLinSepAy(Lin, SepAy)
Next
End Function

Function SyAyzDry(Dry()) As Variant()
Dim I, Dr()
For Each I In Itr(Dry)
    Dr = I
    Push SyAyzDry, SyzDrAsStrSep(Dr) ' Fmtss(X)
Next
End Function

Function SyzDrAsStrSep(Dr()) As String()
Dim J&, U&, O$()
U = UB(Dr)
If U < 0 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = SpcSepStr(CStr(Dr(J)))
Next
SyzDrAsStrSep = O
End Function
Private Sub Z_SyzLinSepAy()
Dim Lin$, SepAy$()
SepAy = SySsl(". . . . . .")
Lin = "AStkShpCst_Rpt.OupFx.Fun."
Ept = Sy("AStkShpCst_Rpt", ".OupFx", ".Fun", ".")
GoSub Tst
Tst:
    Act = SyzLinSepAy(Lin, SepAy)
    C
End Sub

Function ShfTermFmSep$(OLin$, FmPos%, Sep$)
Dim P%: P = InStr(FmPos, OLin, Sep)
If P = 0 Then ShfTermFmSep = OLin: OLin = "": Exit Function
ShfTermFmSep = Left(OLin, P - 1)
OLin = Mid(OLin, P)
End Function

Function SyzLinSepAy(Lin$, SepAy$()) As String()
Dim FmPos%: FmPos = 1
Dim L$: L = Lin
Dim J%
For J = 0 To UB(SepAy)
    PushI SyzLinSepAy, ShfTermFmSep(L, FmPos, SepAy(J))
    If L = "" Then Exit Function
    If J > 0 Then FmPos = Len(SepAy(J - 1)) + 1
Next
If L <> "" Then PushI SyzLinSepAy, L
End Function
Function BrkSyBySepss(Sy$(), Sepss$) As Syay

End Function
Function AlignSyAy(A As Syay) As Syay

End Function
Function LyzSyAy(A As Syay) As String()

End Function
Function FmtSyBySepss(Sy$(), Sepss$) As String()
FmtSyBySepss = LyzSyAy(AlignSyAy(BrkSyBySepss(Sy, Sepss)))
End Function

'=======================
Function DryFmtCommon(Dry(), MaxColWdt%, ShwZer As Boolean) As Variant()
DryFmtCommon = DryAlignCol(DryMkCell(Dry, ShwZer, MaxColWdt))
End Function
Function DrAlign(Dr, W%()) As String()
Dim Cell, J%
For Each Cell In Itr(Dr)
    PushI DrAlign, " " & AlignL(Cell, W(J)) & " "
    J = J + 1
Next
End Function
Function DryAlignColzWdt(Dry(), W%()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DryAlignColzWdt, DrAlign(Dr, W)
Next
End Function
Function DryAlignCol(Dry()) As Variant()
DryAlignCol = DryAlignColzWdt(Dry, WdtAyzDry(Dry))
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

Function IsEqDrCC(Dr1, Dr2, CC%()) As Boolean
Dim J%
For J = 0 To UB(CC)
    If Dr1(CC(J)) <> Dr2(CC(J)) Then Exit Function
Next
IsEqDrCC = True
End Function

Function DryInsBrk(SrtedDry, BrkCC, SepDr$()) As Variant()
Dim CC%(): CC = IxAyzCC(BrkCC)
If Si(CC) = 0 Then DryInsBrk = SrtedDry: Exit Function
Dim Dr, IsBrk As Boolean, LasDr, J&
LasDr = SrtedDry(0)
PushI DryInsBrk, LasDr
For J = 1 To UB(SrtedDry)
    If Not IsEqDrCC(LasDr, SrtedDry(J), CC) Then
        PushI DryInsBrk, SepDr
        LasDr = SrtedDry(J)
    End If
    Push DryInsBrk, SrtedDry(J)
Next
End Function

Function WdtAyzDry(Dry()) As Integer()
Dim J&
For J = 0 To NColzDry(Dry) - 1
    Push WdtAyzDry, AyWdt(ColzDry(Dry, J))
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
