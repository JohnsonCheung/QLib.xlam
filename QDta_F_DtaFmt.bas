Attribute VB_Name = "QDta_F_DtaFmt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Fmt."
Private Const Asm$ = "QDta"
Type Syay
    Syay() As Variant ' Each element is Sy
End Type

Function AlignQteSq(Fny$()) As String()
AlignQteSq = AlignAy(SyzQteSq(Fny))
End Function

Function AlignDrWyAsLin(Ay, WdtAy%()) As String()
Dim S, J&: For Each S In Ay
    PushI AlignDrWyAsLin, Align(S, WdtAy(J))
    J = J + 1
Next
End Function
Function AlignAy(Ay, Optional W0%) As String()
Dim W%: If W0 <= 0 Then W = WdtzAy(Ay) Else W = W0
Dim S: For Each S In Itr(Ay)
    PushI AlignAy, AlignL(S, W)
Next
End Function

Function AlignRzAy(Ay, Optional W0%) As String() 'Fmt-Dr-ToWdt
Dim W%: If W0 <= 0 Then W = WdtzAy(Ay) Else W = W0
Dim I
For Each I In Itr(Ay)
    PushI AlignRzAy, AlignR(I, W)
Next
End Function

Function AlignDrW(Dr, WdtAy%()) As String() 'Fmt-Dr-ToWdt
Dim J%, W%, Cell$, I
For Each I In ResiMax(Dr, WdtAy)
    Cell = I
    W = WdtAy(J)
    PushI AlignDrW, AlignL(Cell, W)
    J = J + 1
Next
End Function

Function DyoSySepss(Ly$(), SepSS$) As Variant()
DyoSySepss = DyoSySep(Sy, SyzSS(SepSS))
End Function

Function DyoSySep(Sy$(), Sep$()) As Variant()
'Ret : a dry wi each rec as a sy of brkg one lin of @Sy.  Each lin is brk by @Sep using fun-BrkLin @@
Dim I, Lin
For Each I In Itr(Sy)
    Lin = I
    PushI DyoSySep, BrkLin(Lin, Sep)
Next
End Function

Function SslSyzDy(Dy()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    Push SslSyzDy, SslzDr(Dr) ' Fmtss(X)
Next
End Function

Private Sub Z_DyoSySepss()
Dim Ly$(), Sep$
GoSub T0
Exit Sub
T0:
    Sep = ". . . . . ."
    Ly = Sy("AStkShpCst_Rpt.OupFx.Fun.")
    Ept = Sy("AStkShpCst_Rpt", ".OupFx", ".Fun", ".")
    GoTo Tst
Tst:
    BrwDy DyoSySepss(Sy, Sep)
    C
    Return
End Sub

Function BrkLin(Lin, Sep$(), Optional IsRmvSep As Boolean) As String()
'Ret : seg ay of a lin sep by @Sep.  Si of seg ret = si of @sep + 1.  Each will have its own sep, expt fst.
'      Segs are not trim and wi/wo by @IsRmvSep.  If not @IsRmvSep, Jn(@Rslt) will eq @Lin @@
Dim L$: L = Lin
Dim O$()
Dim S: For Each S In Sep
    PushI O, ShfBef(L, S)
Next
PushI O, L
If IsRmvSep Then
    Dim J&, Seg: For Each Seg In O
        PushI BrkLin, RmvPfx(Seg, Sep(J))
        J = J + 1
    Next
Else
    BrkLin = O
End If
End Function

Function JnDy(Dy(), Optional QteStr$, Optional Sep$ = " ") As String()
'Ret: :Ly by joining each :Dr in @Dy by @Sep
Dim Dr: For Each Dr In Itr(Dy)
    PushI JnDy, QteJn(Dr, Sep, QteStr)
Next
End Function

Function AlignDyAsLy(Dy()) As String()
AlignDyAsLy = JnDy(AlignDy(Dy))
End Function

Function AlignSepss(Sy$(), SepSS$) As String()
AlignSepss = AlignDyAsLy(DyoSySep(Sy, SyzSS(SepSS)))
End Function

'=======================
Private Function AlignDyzW(Dy(), WdtAy%()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI AlignDyzW, AlignDrW(Dr, WdtAy)
Next
End Function

Function AlignDy(Dy()) As Variant()
AlignDy = AlignDyzW(Dy, WdtAyzDy(Dy))
End Function

Private Function FmtDy__CellDy(Dy(), ShwZer As Boolean, MaxColWdt%) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
   Push FmtDy__CellDy, CellgzDr(Dr, ShwZer, MaxColWdt)
Next
End Function

Private Function CellgzDr(Dr, ShwZer As Boolean, MaxWdt%) As String()
Dim V
For Each V In Itr(Dr)
    PushI CellgzDr, Cellg(V, ShwZer, MaxWdt)
Next
End Function
Private Function CellgzN$(N, MaxW%, ShwZer As Boolean)
If N = 0 Then
    If ShwZer Then
        CellgzN = "0"
    End If
Else
    CellgzN = N
End If
End Function
Private Function CellgzS$(S, MaxW%)
CellgzS = SlashCrLf(Left(S, MaxW))
End Function
Private Function CellgzAy$(Ay, MaxW%)
Dim N&: N = Si(Ay)
If N = 0 Then
    CellgzAy = "*[0]"
Else
    CellgzAy = "*[" & N & "]"
End If
End Function
Function Cellg$(V, Optional ShwZer As Boolean, Optional MaxWdt0% = 30) ' Convert V into a string fit in a cell
Dim O$, MaxWdt%
MaxWdt = EnsBet(MaxWdt0, 1, 1000)
Select Case True
Case IsStr(V):     O = CellgzS(V, MaxWdt)
Case IsBool(V):    O = IIf(V, "True", "")
Case IsNumeric(V): O = CellgzN(V, MaxWdt, ShwZer)
Case IsEmp(V):     O = "#Emp#"
Case IsNull(V):    O = "#Null#"
Case IsArray(V):   O = CellgzAy(V, MaxWdt)
Case IsObject(V):  O = "#O:" & TypeName(V)
Case IsErObj(V)
Case Else:         O = V
End Select
Cellg = O
End Function

Function IsEqAyzIxy(A, B, Ixy&()) As Boolean
Dim J%
For J = 0 To UB(Ixy)
    If A(Ixy(J)) <> B(Ixy(J)) Then Exit Function
Next
IsEqAyzIxy = True
End Function

Function FmtDy__InsSep(Bdy$(), IsBrk() As Boolean, LinzSep$) As String()
If Si(IsBrk) = 0 Then FmtDy__InsSep = Bdy: Exit Function
Dim L, J&: For Each L In Bdy
    If IsBrk(J) Then PushI FmtDy__InsSep, LinzSep
    PushI FmtDy__InsSep, L
    J = J + 1
Next
End Function

Function WdtAyzDy(Dy()) As Integer()
Dim J&
For J = 0 To NColzDy(Dy) - 1
    Push WdtAyzDy, WdtzAy(StrColzDy(Dy, J))
Next
End Function

Function LinzDr(Dr, Optional Sep$ = " ", Optional QteStr$)
'Ret : ret a lin from Dr-QteStr-Sep
LinzDr = Qte(Jn(Dr, Sep), QteStr)
End Function

Function LinzSep$(W%())
LinzSep = LinzDr(DupzWdtAy(W), "-|-", "|-*-|")
End Function

Sub BrwDy(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean, Optional Fmt As EmTblFmt = EmTblFmt.EiTblFmt)
BrwAy FmtDy(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt)
End Sub

Sub BrwDyoSpc(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean)
BrwAy FmtDy(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt:=EiSSFmt)
End Sub
Private Function FmtDy__IsBrkAy(Dy(), BrkCCIxy0) As Boolean()
Dim Ixy&(): Ixy = CvLngAy(BrkCCIxy0)
If Si(Ixy) = 0 Then Exit Function
Dim LasK, CurK, Dr
LasK = AwIxy(Dy(0), Ixy)
For Each Dr In Itr(Dy)
    CurK = AwIxy(Dr, Ixy)
    PushI FmtDy__IsBrkAy, Not IsEqAy(CurK, LasK)
    LasK = CurK
Next
'Insp "QDta_F_DtaFmt.XIsBrk", "Inspect", "Oup(XIsBrk) NoBrk Dy Ixy", "NoFmtr(Boolean())", NoBrk, "NoFmtr(())", Ixy: Stop
End Function

Function DyoInsIx(Dy()) As Variant()
' Ret Dy with each row has ix run from 0..{N-1} in front
Dim Ix&, Dr: For Each Dr In Itr(Dy)
    Dr = InsEle(Dr, Ix)
    PushI DyoInsIx, Dr
    Ix = Ix + 1
Next
End Function
Function DupzWdtAy(W%(), Optional C$ = "-") As String()
Dim I: For Each I In W
    Push DupzWdtAy, Dup(C, I)
Next
End Function

Function FmtDy(Dy(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt) _
As String()
If Si(Dy) = 0 Then Exit Function
Dim CellDy(): CellDy = FmtDy__CellDy(Dy, ShwZer, MaxColWdt)

Dim W%(): W = WdtAyzDy(CellDy)

Dim AlignDy(): AlignDy = AlignDyzW(CellDy, W)
Dim S$:              S = IIf(Fmt = EiSSFmt, " ", " | ")
Dim Q$:              Q = IIf(Fmt = EiSSFmt, "", "| * |")
Dim Bdy$():        Bdy = FmtDyoSepQte(AlignDy, S, Q)

Dim SepDr$(): SepDr = DupzWdtAy(W)
                  S = IIf(Fmt = EiSSFmt, " ", "-|-")
                  Q = IIf(Fmt = EiSSFmt, "", "|-*-|")
Dim Sep$:       Sep = LinzDr(SepDr, S, Q)

Dim IsBrk() As Boolean:  IsBrk = FmtDy__IsBrkAy(Dy, BrkCCIxy0)
Dim BdyBrk$():          BdyBrk = FmtDy__InsSep(Bdy, IsBrk, Sep)
                         FmtDy = Sy(Sep, BdyBrk, Sep)
End Function

Function FmtDyoSepQte(Dy(), Sep$, QteStr$) As String()
Dim Dr
For Each Dr In Itr(Dy)
    PushI FmtDyoSepQte, LinzDr(Dr, Sep, QteStr)
Next
End Function

Sub DmpDyoSpc(Dy(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean)
D FmtDy(Dy, MaxColWdt, BrkCCIxy0, ShwZer, Fmt:=EiSSFmt)
End Sub

Sub DmpDy(Dy(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt)
D FmtDy(Dy, MaxColWdt, BrkCCIxy0, ShwZer, Fmt)
End Sub


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
AyA = FmtDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N1)
AyB = FmtDrs(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N2)
T = Sy(Tit, ULinDbl(Tit))
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
AyA = FmtDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N1)
AyB = FmtDrs(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N2)
AyC = FmtDrs(C, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt, Nm:=N3)
T = Sy(Tit, ULinDbl(Tit))
Ay = Sy(T, AyA, AyB, AyC)
Brw Ay, FnPfx, UseVc
End Sub

Sub BrwDrs4(A As Drs, B As Drs, C As Drs, D As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, _
Optional FnPfx$, Optional UseVc As Boolean)
Dim Ay$(), AyA$(), AyB$(), AyC$(), AyD$()
AyA = FmtDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyB = FmtDrs(B, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyC = FmtDrs(C, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
AyD = FmtDrs(D, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt)
Ay = Sy(AyA, AyB, AyC, AyD)
Brw Ay, FnPfx, UseVc
End Sub

Sub BrwDrs(A As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EmIxCol.EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, _
Optional FnPfx$, Optional UseVc As Boolean)
BrwAy FmtDrs(A, MaxColWdt, BrkColnn, ShwZer, IxCol, Fmt), FnPfx, UseVc
End Sub
Function DrszFmtg(DrsFmtg$()) As Drs
Dim TitLin$: TitLin = DrsFmtg(1)
Dim Fny$(): Fny = AeFstLas(SyzTrim(Split(TitLin, "|")))
Dim Dy()
    Dim J&
    For J = 3 To UB(DrsFmtg) - 1
        PushI Dy, AvzAy(AeFstLas(RmvFstChrzAy(RSyzTrim(Split(DrsFmtg(J), "|")))))
    Next
DrszFmtg = Drs(Fny, Dy)
End Function

Function LinzDrsR(A As Drs, Optional Nm$) As String()
If NoReczDrs(A) Then Exit Function
Dim AFny$(): AFny = Sy("#", AlignAy(A.Fny))

Dim Ly$(), Lixy&()
    Dim N&: N = Si(A.Dy)
    Dim Dr, J&: For Each Dr In Itr(A.Dy)
        J = J + 1
        PushI Ly, Empty
        PushI Lixy, UB(Ly)
        Dim I$: I = J & " of " & N
        Dim Av(): Av = AddAy(Array(I), Dr)
        PushIAy Ly, LyzNyAv(AFny, Av)
    Next
Dim Align$(): Align = AlignAy(Ly)
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

Private Function FmtDrs__SumDr(D As Drs, IsSum As Boolean, SumFF$) As Variant()
If Not IsSum Then Exit Function
End Function

Private Function FmtDrs__NoRec(D As Drs, NmBox$()) As String()
Dim S$:        S = JnSpc(D.Fny)
Dim S1$:           If S1 = "" Then S1 = " (No Fny)"
Dim Lin$:    Lin = "(NoRec) " & S1
          FmtDrs__NoRec = Sy(NmBox, Lin)
End Function

Function FmtDrs(D As Drs, _
Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EmIxCol.EiBeg1, _
Optional Fmt As EmTblFmt = EiTblFmt, Optional Nm$, Optional IsSum As Boolean, Optional SumFF$) As String()
'Fm IsSum    : If true all num col will have a sum as las lin in the fmt
'Fm BrkColnn : if changed, insert a break line if BrkColNm is given
Dim NmBox$(): If Nm <> "" Then NmBox = Box(Nm)
If NoReczDrs(D) Then FmtDrs = FmtDrs__NoRec(D, NmBox): Exit Function
Dim IxD As Drs:    IxD = AddColzIx(D, IxCol)                     ' Add Col-Ix
Dim IxyB&():      IxyB = Ixy(IxD.Fny, TermAy(BrkColnn))          ' Ixy-Of-BrkCol
Dim Dy():           Dy = AddEle(IxD.Dy, IxD.Fny)                 ' Dy-With-Fny
Dim SumDr():     SumDr = FmtDrs__SumDr(D, IsSum, SumFF)          '              Sam-ele-as-Col-or-no-ele.  Each ele is Sum of the num-col or emp
                         If IsSum Then PushI Dy, SumDr
Dim Bdy$():        Bdy = FmtDy(Dy, MaxColWdt, IxyB, ShwZer, Fmt) ' Ly-For-Dy
Dim Sep$:          Sep = Pop(Bdy)                              ' Sep-Lin
Dim Hdr$:          Hdr = Pop(Bdy)                           ' Hdr-Lin
Dim Sum$:          Sum = Pop(Bdy)
Dim O$():            O = Sy(NmBox, Sep, Hdr, Bdy, Sep)
:                        If IsSum Then PushI O, Sum
                FmtDrs = O
End Function

Function FmtDt(A As DT, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional IxCol As EmIxCol = EiBeg1) As String()
PushI FmtDt, "*Tbl " & A.DtNm
PushIAy FmtDt, FmtDrs(DrszDt(A), MaxColWdt, BrkColNm, ShwZer, IxCol)
End Function

Private Sub Z_FmtDrs()
Dim A As Drs, MaxColWdt%, BrkColVbl$, ShwZer As Boolean, IxCol As EmIxCol
A = SampDrs
GoSub Tst
Exit Sub
Tst:
    Act = FmtDrs(A, MaxColWdt, BrkColVbl, ShwZer, IxCol)
    Brw Act: Stop
    C
    Return
End Sub

Private Sub Z_FmtDt()
Dim A As DT, MaxColWdt%, BrkColNm$, ShwZer As Boolean
'--
A = SampDt1
'Ept = Z_TimStrpt1
GoSub Tst
'--
Exit Sub
Tst:
    Act = FmtDt(A, MaxColWdt, BrkColNm, ShwZer)
    C
    Return
End Sub

Private Sub Z()
Z_FmtDrs
'Z_FmtDt
End Sub
