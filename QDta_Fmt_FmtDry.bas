Attribute VB_Name = "QDta_Fmt_FmtDry"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Fmt_Dy_Fun."
Private Const Asm$ = "QDta"
Type Syay
    Syay() As Variant ' Each element is Sy
End Type

Function AlignQteSq(Fny$()) As String()
AlignQteSq = AlignzAy(SyzQteSq(Fny))
End Function

Function AlignzDrWyAsLin(Ay, WdtAy%()) As String()
Dim S, J&: For Each S In Ay
    PushI AlignzDrWyAsLin, Align(S, WdtAy(J))
    J = J + 1
Next
End Function
Function AlignzAy(Ay, Optional W0%) As String()
Dim W%: If W0 <= 0 Then W = WdtzAy(Ay) Else W = W0
Dim S: For Each S In Itr(Ay)
    PushI AlignzAy, AlignL(S, W)
Next
End Function

Function AlignRzAy(Ay, Optional W0%) As String() 'Fmt-Dr-ToWdt
Dim W%: If W0 <= 0 Then W = WdtzAy(Ay) Else W = W0
Dim I
For Each I In Itr(Ay)
    PushI AlignRzAy, AlignR(I, W)
Next
End Function

Function AlignzDrW(Dr, WdtAy%()) As String() 'Fmt-Dr-ToWdt
Dim J%, W%, Cell$, I
For Each I In ResiMax(Dr, WdtAy)
    Cell = I
    W = WdtAy(J)
    PushI AlignzDrW, AlignL(Cell, W)
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

Function AlignzDyAsLy(Dy()) As String()
AlignzDyAsLy = JnDy(AlignzDy(Dy))
End Function

Function AlignzSepss(Sy$(), SepSS$) As String()
AlignzSepss = AlignzDyAsLy(DyoSySep(Sy, SyzSS(SepSS)))
End Function

'=======================
Function AlignzDyW(Dy(), WdtAy%()) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    PushI AlignzDyW, AlignzDrW(Dr, WdtAy)
Next
End Function

Function AlignzDy(Dy()) As Variant()
AlignzDy = AlignzDyW(Dy, WdtAyzDy(Dy))
End Function

Private Function CellgzDy(Dy(), ShwZer As Boolean, MaxColWdt%) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
   Push CellgzDy, CellgzDr(Dr, ShwZer, MaxColWdt)
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

Function A_Ins(NoBrk As Boolean, IsBrk() As Boolean, JnD$(), LinzSep$) As String()
If NoBrk Then A_Ins = JnD: Exit Function
Dim L, J&: For Each L In JnD
    If IsBrk(J) Then PushI A_Ins, LinzSep
    PushI A_Ins, L
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
LinzSep = LinzDr(DupzWy(W), "-|-", "|-*-|")
End Function

Sub BrwDy(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean, Optional Fmt As EmTblFmt = EmTblFmt.EiTblFmt)
BrwAy FmtDy(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt)
End Sub

Sub BrwDyoSpc(A(), Optional MaxColWdt% = 100, Optional BrkCCIxy, Optional ShwZer As Boolean)
BrwAy FmtDy(A, MaxColWdt, BrkCCIxy, ShwZer, Fmt:=EiSSFmt)
End Sub
Private Function XIsBrk(NoBrk As Boolean, Dy(), Ixy&()) As Boolean()
If NoBrk Then Exit Function
Dim LasK, CurK, Dr
LasK = AwIxy(Dy(0), Ixy)
For Each Dr In Itr(Dy)
    CurK = AwIxy(Dr, Ixy)
    PushI XIsBrk, Not IsEqAy(CurK, LasK)
    LasK = CurK
Next
End Function

Function DyoInsIx(Dy()) As Variant()
' Ret Dy with each row has ix run from 0..{N-1} in front
Dim Ix&, Dr: For Each Dr In Itr(Dy)
    Dr = InsEle(Dr, Ix)
    PushI DyoInsIx, Dr
    Ix = Ix + 1
Next
End Function

Function FmtDy(Dy(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt) _
As String()
If Si(Dy) = 0 Then Exit Function
Dim StrD():     StrD = CellgzDy(Dy, ShwZer, MaxColWdt)
Dim W%():          W = WdtAyzDy(StrD)
Dim AlignD(): AlignD = AlignzDyW(StrD, W)

Dim Ixy&():               Ixy = CvLngAy(BrkCCIxy0)
Dim NoBrk As Boolean:   NoBrk = Si(Ixy) = 0
Dim IsBrk() As Boolean: IsBrk = XIsBrk(NoBrk, Dy, Ixy)

Dim S$:       S = IIf(Fmt = EiSSFmt, " ", " | ")
Dim Q$:       Q = IIf(Fmt = EiSSFmt, "", "| * |")
Dim JnD$(): JnD = FmtDyoSepQte(AlignD, S, Q)
Dim Sep$(): Sep = DupzWy(W)
              S = IIf(Fmt = EiSSFmt, " ", "-|-")
              Q = IIf(Fmt = EiSSFmt, "", "|-*-|")
Dim L$:       L = LinzDr(Sep, S, Q)
Dim Ins$(): Ins = A_Ins(NoBrk, IsBrk, JnD, L)
         FmtDy = Sy(L, Ins, L)
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

