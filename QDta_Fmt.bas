Attribute VB_Name = "QDta_Fmt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Fmt."
Private Const Asm$ = "QDta"

Sub VcDrs(A As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1, _
Optional Fmt As EmTblFmt, _
Optional FnPfx$, Optional UseVc As Boolean)
BrwDrs A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, FnPfx, UseVc
End Sub

Sub BrwDrs2(A As Drs, B As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, Optional NN$, Optional Tit$ = "Brw 2 Drs", _
Optional FnPfx$, Optional UseVc As Boolean)
Dim Ay$(), AyA$(), AyB$(), N1$, N2$, T$()
N1 = DftStr(BefSpc(NN), "Drs-A")
N2 = DftStr(AftSpc(NN), " Drs-B")
AyA = LinzDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N1)
AyB = LinzDrs(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N2)
T = Sy(Tit, UnderLinDbl(Tit))
Ay = Sy(T, AyA, AyB)
Brw Ay, FnPfx, UseVc
End Sub

Sub BrwDrs3(A As Drs, B As Drs, C As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EmIxCol.EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, Optional NN$, Optional Tit$ = "Brw 3 Drs", _
Optional FnPfx$, Optional UseVc As Boolean)
Dim Ay$(), AyA$(), AyB$(), AyC$(), N1$, N2$, N3$, T$()
N1 = DftStr(T1(NN), "Drs-A")
N2 = DftStr(T2(NN), " Drs-B")
N3 = DftStr(RmvTT(NN), " Drs-C")
AyA = LinzDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N1)
AyB = LinzDrs(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N2)
AyC = LinzDrs(C, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N3)
T = Sy(Tit, UnderLinDbl(Tit))
Ay = Sy(T, AyA, AyB, AyC)
Brw Ay, FnPfx, UseVc
End Sub

Sub BrwDrs4(A As Drs, B As Drs, C As Drs, D As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, _
Optional FnPfx$, Optional UseVc As Boolean)
Dim Ay$(), AyA$(), AyB$(), AyC$(), AyD$()
AyA = LinzDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyB = LinzDrs(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyC = LinzDrs(C, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyD = LinzDrs(D, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
Ay = Sy(AyA, AyB, AyC, AyD)
Brw Ay, FnPfx, UseVc
End Sub

Sub BrwDrs(A As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EmIxCol.EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, _
Optional FnPfx$, Optional UseVc As Boolean)
Dim Ay$()
Ay = LinzDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
BrwAy Ay, FnPfx, UseVc
End Sub
Function DrszFmtg(DrsFmtg$()) As Drs
Dim TitLin$: TitLin = DrsFmtg(1)
Dim Fny$(): Fny = AyeFstLas(SyzTrim(Split(TitLin, "|")))
Dim Dry()
    Dim J&
    For J = 3 To UB(DrsFmtg) - 1
        PushI Dry, AvzAy(AyeFstLas(RmvFstChrzAy(RSyzTrim(Split(DrsFmtg(J), "|")))))
    Next
DrszFmtg = Drs(Fny, Dry)
End Function

Function LinzDrsR(A As Drs, Optional Nm$) As String()
If NoReczDrs(A) Then Exit Function
Dim AFny$(): AFny = Sy("#", AlignzAy(A.Fny))

Dim Ly$(), Lixy&()
    Dim N&: N = Si(A.Dry)
    Dim Dr, J&: For Each Dr In Itr(A.Dry)
        J = J + 1
        PushI Ly, Empty
        PushI Lixy, UB(Ly)
        Dim I$: I = J & " of " & N
        Dim Av(): Av = AyzAdd(Array(I), Dr)
        PushIAy Ly, LyzNyAv(AFny, Av)
    Next
Dim Align$(): Align = AlignzAy(Ly)
Dim Q$()
    Dim L: For Each L In Align
        Push Q, "| " & L & " |"
    Next
'== Oup ===
Dim O$(): O = Q
Dim W%:   W = Len(Align(0))
Dim Lin$:   Lin = "|-" & Dup("-", W) & "-|"
Dim Ix: For Each Ix In Itr(Lixy)
    O(Ix) = Lin
Next
PushI O, Lin
LinzDrsR = O
End Function

Function LinzDrs(A As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EmIxCol.EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, Optional Nm$) As String() ' _
If BrkColNm changed, insert a break line if BrkColNm is given
Dim NmBox$(): If Nm <> "" Then NmBox = Box(Nm)

If NoReczDrs(A) Then
    Dim S$: S = JnSpc(A.Fny)
    If S = "" Then S = " (No Fny)"
    Dim Lin$: Lin = "(NoRec) " & S
    LinzDrs = Sy(NmBox, Lin)
    Exit Function
End If
      
Dim IxD As Drs:      IxD = DrsAddIxCol(A, IxCol)
Dim IxyB&():        IxyB = Ixy(IxD.Fny, TermAy(BrkColnn))
Dim WiFnyDry(): WiFnyDry = AddEle(IxD.Dry, IxD.Fny)
Dim Ly$():            Ly = FmtDry(WiFnyDry, MaxColWdt, IxyB, ShwZer, Fmt)
Dim H$:                H = LasSndEle(Ly)
Dim L$:                L = LasEle(Ly)
Dim Ly1$():          Ly1 = AyeLasNEle(Ly, 2)
                  LinzDrs = Sy(NmBox, L, H, Ly1, L)
End Function

Function FmtDt(A As Dt, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1) As String()
PushI FmtDt, "*Tbl " & A.DtNm
PushIAy FmtDt, LinzDrs(DrszDt(A), MaxColWdt, BrkColNm, ShwZer, IxCol)
End Function

Private Sub Z_LinzDrs()
Dim A As Drs, MaxColWdt%, BrkColVbl$, ShwZer As Boolean, IxCol As EmIxCol
A = SampDrs
GoSub Tst
Exit Sub
Tst:
    Act = LinzDrs(A, MaxColWdt, BrkColVbl, ShwZer, IxCol)
    Brw Act: Stop
    C
    Return
End Sub

Private Sub Z_FmtDt()
Dim A As Dt, MaxColWdt%, BrkColNm$, ShwZer As Boolean
'--
A = SampDt1
'Ept = Z_DteTimStrpt1
GoSub Tst
'--
Exit Sub
Tst:
    Act = FmtDt(A, MaxColWdt, BrkColNm, ShwZer)
    C
    Return
End Sub

Private Sub Z()
Z_LinzDrs
'Z_FmtDt
End Sub
