Attribute VB_Name = "MDta_Fmt_FmtSpcSep"
Option Explicit
Private Function DrzExtendU(Dr, Wdt%(), ExtendToU%)
Dim O: O = Dr
ReDim Preserve O(UB(Wdt))
Dim J%
For J = UB(Dr) + 1 To UB(Wdt)
    O(J) = Space(Wdt(J))
Next
DrzExtendU = O
End Function
Function FmtDr$(Dr, Wdt%(), Fmt As eDryFmt)
Dim Dr1$(): Dr1 = DrzExtendU(Dr, Wdt, UB(Wdt))
Dim O$(), J%
    Dim W
    For Each W In Wdt
        PushI O, AlignL(Dr1(J), W)
        J = J + 1
    Next
FmtDr = "|" & Join(O, SepChrzDryFmt(Fmt)) & "|"
End Function

Function DryzAySepSS(Ay, SepSS$) As Variant()
Dim Lin, SepAy$()
SepAy = TermAy(SepSS)
For Each Lin In Itr(Ay)
    PushI DryzAySepSS, DrzLinSepAy(Lin, SepAy)
Next
End Function

Function DryFmtCellSpcSep(A()) As Variant()
Dim Dr
For Each Dr In Itr(A)
    Push DryFmtCellSpcSep, DrFmtCellSpcSep(Dr) ' Fmtss(X)
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
FmtAyzSepSS = FmtDry(DryzAySepSS(Ay, SepSS), Fmt:=eSpcSep)
End Function
