Attribute VB_Name = "MDta_Fmt_Dry_Fun"
Option Explicit
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

Function DryzAySepSS(Ay, SepSS$) As Variant()
Dim Lin, SepAy$()
SepAy = TermAy(SepSS)
For Each Lin In Itr(Ay)
    PushI DryzAySepSS, DrzLinSepAy(Lin, SepAy)
Next
End Function

Function DryFmtCellAsStr(A()) As Variant()
Dim Dr
For Each Dr In Itr(A)
    Push DryFmtCellAsStr, DrFmtCellSpcSep(Dr) ' Fmtss(X)
Next
End Function

Function DrFmtCellSpcSep(Dr) As String()
Dim J&, U&, O$()
U = UB(Dr)
If U < 0 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = SpcSepStr(Dr(J))
Next
DrFmtCellSpcSep = O
End Function
Private Sub Z_DrzLinSepAy()
Dim Lin$, SepAy$()
SepAy = SySsl(". . . . . .")
Lin = "AStkShpCst_Rpt.OupFx.Fun."
Ept = Sy("AStkShpCst_Rpt", ".OupFx", ".Fun", ".")
GoSub Tst
Tst:
    Act = DrzLinSepAy(Lin, SepAy)
    C
End Sub
Function ShfTermFmSep$(OLin, FmPos%, Sep$)
Dim P%: P = InStr(FmPos, OLin, Sep)
If P = 0 Then ShfTermFmSep = OLin: OLin = "": Exit Function
ShfTermFmSep = Left(OLin, P - 1)
OLin = Mid(OLin, P)
End Function
Function DrzLinSepAy(Lin, SepAy$()) As String()
Dim FmPos%: FmPos = 1
Dim L$: L = Lin
Dim J%
For J = 0 To UB(SepAy)
    PushI DrzLinSepAy, ShfTermFmSep(L, FmPos, SepAy(J))
    If L = "" Then Exit Function
    If J > 0 Then FmPos = Len(SepAy(J - 1)) + 1
Next
If L <> "" Then PushI DrzLinSepAy, L
End Function

Function FmtAyzSepSS(Ay, SepSS$) As String()
FmtAyzSepSS = FmtDryAsSpcSep(DryzAySepSS(Ay, SepSS))
End Function

'=======================
Function DryFmtCommon(Dry(), MaxColWdt%, ShwZer As Boolean) As Variant()
DryFmtCommon = DryAlignCol(DryStrfy(Dry, ShwZer, MaxColWdt))
End Function
Function DrAlign(Dr, W%()) As String()
Dim Cell, J%
For Each Cell In Itr(Dr)
    PushI DrAlign, AlignL(Cell, W(J))
    J = J + 1
Next
End Function
Function DryAlignCol(Dry()) As Variant()
Dim W%(): W = WdtAyzDry(Dry)
Dim Dr
For Each Dr In Itr(Dry)
    PushI DryAlignCol, DrAlign(Dr, W)
Next
End Function
Private Function DryStrfy(Dry(), ShwZer As Boolean, MaxColWdt%) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
   Push DryStrfy, DrStrfy(Dr, ShwZer, MaxColWdt)
Next
End Function

Private Function DrStrfy(Dr, ShwZer As Boolean, MaxWdt%) As String()
Dim I
For Each I In Itr(Dr)
    PushI DrStrfy, Left(StrzVal(I, ShwZer), MaxWdt)
Next
End Function
Private Function StrzVal$(V, Optional ShwZer As Boolean) ' Convert V into a string in a cell
'SpcSepStr is a string can be displayed in a cell
Select Case True
Case IsNumeric(V)
    If V = 0 Then
        If ShwZer Then
            StrzVal = "0"
        End If
    Else
        StrzVal = V
    End If
Case IsEmp(V):
Case IsArray(V)
    Dim N&: N = Si(V)
    If N = 0 Then
        StrzVal = "*[0]"
    Else
        StrzVal = "*[" & N & "]" & V(0)
    End If
Case IsObject(V): StrzVal = TypeName(V)
Case Else:        StrzVal = V
End Select
End Function

Function DryInsSep(Dry, BrkColIxAy, SepDr$()) As Variant()
If Si(BrkColIxAy) = 0 Then DryInsSep = Dry: Exit Function
Dim Dr, IsBrk As Boolean, LasCell$, Fst As Boolean
'LasCell = Dry(0)(BrkColIxAy)
For Each Dr In Dry
    If Fst Then
        Fst = False
    Else
'        IsBrk = LasCell = Dr(BrkColIx)
    End If
    If IsBrk Then
        PushI DryInsSep, SepDr
'        LasCell = Dr(BrkColIx)
    End If
    Push DryInsSep, Dr
Next
End Function

Private Function WdtAyzDry(A()) As Integer()
Dim J%
For J = 0 To NColzDry(A) - 1
    Push WdtAyzDry, WdtzAy(ColzDry(A, J))
Next
End Function


Function SepLin$(W%(), Sep$)
SepLin = SepLinzSepDr(SepDr(W), Sep)
End Function

Function SepDr(W%()) As String()
Dim I
For Each I In W
    Push SepDr, Dup("-", I)
Next
End Function

Function SepLinzSepDr$(SepDr$(), Sep$)
SepLinzSepDr = "|" & Join(SepDr, Sep) & "|"
End Function

Function LinFmDrByJnCell$(Dr, Optional Sep$ = " | ", Optional Pfx$ = "| ", Optional Sfx$ = " |")
LinFmDrByJnCell = Pfx & Jn(Dr, Sep) & Sfx
End Function

Function FmtDryByJnCell(Dry(), Optional Sep$ = " | ", Optional Pfx$ = "| ", Optional Sfx$ = " |") As String()
Dim Dr
For Each Dr In Itr(Dry)
    PushI FmtDryByJnCell, LinFmDrByJnCell(Dr, Sep, Pfx, Sfx)
Next
End Function
