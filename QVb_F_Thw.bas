Attribute VB_Name = "QVb_F_Thw"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Thw."
Type CfgInf
    ShwInf As Boolean
    ShwTim As Boolean
End Type
Type CfgSql
    FmtSql As Boolean
End Type
Type Cfg
    Inf As CfgInf
    Sql As CfgSql
End Type

Public Property Get Cfg() As Cfg
Static X As Boolean, Y As Cfg
If Not X Then
    X = True
    Y.Sql.FmtSql = True
    Y.Inf.ShwInf = True
    Y.Inf.ShwTim = True
End If
Cfg = Y
End Property

Sub ThwIf_NegEle(Ay, Fun$)
Const CSub$ = CMod & "ThwIf_NEgEle"
Dim I, J&, O$()
For Each I In Itr(Ay)
    If I < 0 Then
        PushI O, J & ": " & I
        J = J + 1
    End If
Next
If Si(O) > 0 Then
    Thw CSub, "In [Ay], there are [negative-element (Ix Ele)]", Ay, O
End If
End Sub

Sub ThwIf_AyabNE(A, B, Optional N1$ = "Exp", Optional N2$ = "Act")
ThwIf_Er ChkEqAy(A, B, N1, N2), CSub
End Sub

Sub ThwIf_NE(A, B, Optional N1$ = "A", Optional N2$ = "B")
Const CSub$ = CMod & "ThwIf_NE"
ThwIf_DifTy A, B, N1, N2
Dim IsLinesA As Boolean, IsLinesB As Boolean
IsLinesA = IsLines(A)
IsLinesB = IsLines(B)
Select Case True
Case IsLinesA Or IsLinesB: If A <> B Then CmprLines CStr(A), CStr(B), Hdr:=FmtQQ("Lines [?] [?] not eq.", N1, N2): Stop: Exit Sub
Case IsStr(A):    If A <> B Then CmprStr CStr(A), CStr(B), Hdr:=FmtQQ("String [?] [?] not eq.", N1, N2): Stop: Exit Sub
Case IsDic(A):    If Not IsEqDic(CvDic(A), CvDic(B)) Then BrwCmpgDicAB CvDic(A), CvDic(B): Stop: Exit Sub
Case IsArray(A):  ThwIf_DifAy A, B, N1, N2
Case IsObject(A): If ObjPtr(A) <> ObjPtr(B) Then Thw CSub, "Two object are diff", FmtQQ("Ty-? Ty-?", N1, N2), TypeName(A), TypeName(B)
Case Else:
    If A <> B Then
        Thw CSub, "A B NE", "A B", A, B
        Exit Sub
    End If
End Select
End Sub

Private Sub ThwIf_DifAy(AyA, AyB, N1$, N2$)
ThwIf_DifSi AyA, AyB, CSub
ThwIf_DifTy AyA, AyB, N1, N2
Dim J&, A
For Each A In Itr(AyA)
    If Not IsEq(A, AyB(J)) Then
        Dim NN$: NN = FmtQQ("AyTy AySi Dif-At Ay-?-Ele-?-Ty Ay-?-Ele-?-Ty Ay-?-Ele-Val Ay-?-Ele-Val Ay-? Ay-?", N1, J, N2, J, N1, N2, N1, N2)
        Thw CSub, "There is ele in 2 Ay are diff", NN, TypeName(AyA), Si(AyA), J, TypeName(A), TypeName(AyB(J)), A, AyB(J), AyA, AyB
        Exit Sub
    End If
    J = J + 1
Next
End Sub

Sub ThwIf_DifTy(A, B, Optional N1$ = "A", Optional N2$ = "B")
If TypeName(A) = TypeName(B) Then Exit Sub
Dim NN$
NN = FmtQQ("?-TyNm ?-TyNm", N1, N2)
Thw CSub, "Type Diff", NN, TypeName(A), TypeName(B)
End Sub

Sub ThwIf_DifSi(A, B, Fun$)
If Si(A) <> Si(B) Then Thw Fun, "Si-A <> Si-B", "Si-A Si-B", Si(A), Si(B)
End Sub

Sub ThwIf_DifFF(A As Drs, FF$, Fun$)
If JnSpc(A.Fny) <> FF Then Thw Fun, "Drs-FF <> FF", "Drs-FF FF", JnSpc(A.Fny), FF
End Sub

Sub ThwIf_ObjNE(A, B, Fun$, Msg$, Nav())
If IsEqObj(A, B) Then ThwNav Fun, Msg, Nav
End Sub

Sub ThwIf_NoSrt(Ay, Fun$)
If IsSrtedzAy(Ay) Then Thw Fun, "Array should be sorted", "Ay-Ty Ay", TypeName(Ay), Ay
End Sub

Sub Insp(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Dim F$: If Fun <> "" Then F = " (@" & Fun & ")"
Dim A$(): A = BoxS("Insp: " & Msg & F)
BrwAy Sy(A, LyzNav(Nav))
End Sub

Sub Thw(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Dim A$(): A = BoxS("Program error")
ThwNav Fun, JnCrLf(Sy(A, Msg)), Nav
End Sub

Sub ThwNav(Fun$, Msg$, Nav())
BrwAy LyzFunMsgNav(Fun, Msg, Nav)
Halt
End Sub

Sub Ass(A As Boolean)
Debug.Assert A
End Sub

Sub ThwIf_Nothing(A, VarNm$, Fun$)
If Not IsNothing(A) Then Exit Sub
Thw Fun, FmtQQ("Given[?] is nothing", VarNm)
End Sub
Sub ThwIf_NotAy(A, Fun$)
If IsArray(A) Then Exit Sub
Thw Fun, "Given parameter should be array, but now TypeName=" & TypeName(A)
End Sub
Sub ThwIf_NotStr(A, Fun$)
If IsStr(A) Then Exit Sub
Thw Fun, "Given parameter should be str, but now TypeName=" & TypeName(A)
End Sub
Sub ThwIf_Never(Fun$, Optional Msg$ = "Program should not reach here")
Thw Fun, Msg
End Sub

Sub Halt(Optional Fun$)
Err.Raise -1, Fun, "Please check messages opened in notepad"
End Sub

Sub Done()
MsgBox "Done"
End Sub

Function AddNNAv(Nav(), NN$, Av()) As Variant()
Dim O(): O = Nav
If Si(O) = 0 Then
    PushI O, NN
Else
    O(0) = O(0) & " " & NN
End If
PushAy O, Av
AddNNAv = O
End Function

Function AddNmV(Nav(), Nm$, V) As Variant()
AddNmV = AddNNAv(Nav, Nm, Av(V))
End Function

Sub ThwIf_ErMsg(Er$(), Fun$, Msg$, ParamArray Nap())
If Si(Er) = 0 Then Exit Sub
Dim Nav(): Nav = Nap
ThwNav Fun, Msg, AddNmV(Nav, "Er", Er)
End Sub

Sub ThwIf_Er(Er$(), Fun$)
If Si(Er) = 0 Then Exit Sub
BrwAy AddSy(LyzFunMsgNap(Fun, ""), Er)
Halt
End Sub

Sub ThwLoopingTooMuch(Fun$)
Thw Fun, "Looping too much"
End Sub

Sub ThwPmEr(VzPm, Fun$, Optional MsgWhyPmEr$ = "Invalid value")
Thw Fun, "Parameter error: " & MsgWhyPmEr, "Pm-Type Pm-Val", TypeName(VzPm), FmtV(VzPm)
End Sub

Sub D(Optional V, Optional OupTy As EmOupTy)
Dim A$(): A = FmtV(V)
Select Case True
Case OupTy = EiOtDmp: DmpAy A
Case OupTy = EiOtVc: VcAy A
Case OupTy = EiOtBrw:  BrwAy A
Case Else: BrwAy A
End Select
End Sub

Sub Dmp(A, Optional OupTy As EmOupTy)
D A, OupTy
End Sub

Sub DmpTy(A)
Debug.Print TypeName(A)
End Sub

Sub DmpAyWithIx(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print J; ": "; Ay(J)
Next
End Sub

Sub DmpAy(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print Ay(J)
Next
End Sub

Sub InfLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
D LinzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub InfNav(Fun$, Msg$, Nav())
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub Inf(Fun$, Msg$, ParamArray Nap())
If Not Cfg.Inf.ShwInf Then Exit Sub
Dim Nav(): Nav = Nap
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub WarnLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Debug.Print LinzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub Warn(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub

Private Sub Z_LyzObjPP()
Dim Obj As Object, PP$
GoSub T0
Exit Sub
T0:
    Set Obj = New Dao.Field
    PP = "Name Type Size"
    GoTo Tst
Tst:
    Act = LyzObjPP(Obj, PP)
    C
    Return
End Sub

Private Sub Z()
Dim A$
Dim B()
Dim C
Dim D%
Dim F$()
Dim XX
End Sub

Sub StopEr(Er$())
If Si(Er) = 0 Then Exit Sub
BrwAy Er
Stop
End Sub

Sub ThwEqObj(A, B, Fun$, Optional Msg$ = "Two given object cannot be same")
If IsEqObj(A, B) Then Thw Fun, Msg
End Sub

Function VblzLines$(Lines$)
VblzLines = Replace(RmvCr(Lines), vbLf, "|")
End Function

Function LinzFunMsg$(Fun$, Msg$)
Dim F$: F = IIf(Fun = "", "", " | @" & Fun)
Dim A$: A = Msg & F
If Cfg.Inf.ShwTim Then
    LinzFunMsg = NowStr & " | " & A
Else
    LinzFunMsg = A
End If
End Function

Function LyzFunMsgNav(Fun$, Msg$, Nav()) As String()
Dim A$(): A = LyzFunMsg(Fun, Msg)
Dim B$(): B = IndentSy(LyzNav(Nav))
LyzFunMsgNav = AddAy(A, B)
End Function

Function LyzFunMsgNap(Fun$, Msg$, ParamArray Nap()) As String()
Dim Nav(): Nav = Nap
If Fun = "" And Msg = "" And Si(Nav) = 0 Then Exit Function
LyzFunMsgNap = LyzFunMsgNav(Fun, Msg, Nav)
End Function

Function LyzFunMsgObjPP(Fun$, Msg$, Obj As Object, PP$) As String()
LyzFunMsgObjPP = AddAy(LyzFunMsg(Fun, Msg), LyzObjPP(Obj, PP))
End Function

Function LyzFunMsgNyAv(Fun$, Msg$, Ny$(), Av()) As String()
LyzFunMsgNyAv = AddAy(LyzFunMsg(Fun, Msg), IndentSy(LyzNyAv(Ny, Av)))
End Function

Function LyzNv(Nm$, V, Optional Sep$ = ": ") As String()
Dim Ly$(): Ly = FmtV(V)
Dim J%, S$
If Si(Ly) = 0 Then
    PushI LyzNv, Nm & Sep
Else
    PushI LyzNv, Nm & Sep & Ly(0)
End If
S = Space(Len(Nm) + Len(Sep))
For J = 1 To UB(Ly)
    PushI LyzNv, S & Ly(J)
Next
End Function

Function LinzNv$(Nm$, V)
LinzNv = Nm & "=[" & Cell(V) & "]"
End Function

Function LyzMsgNap(Msg$, ParamArray Nap()) As String()
Dim Nav(): Nav = Nap
LyzMsgNap = LyzMsgNav(Msg, Nav)
End Function

Function LyzNmDrs(Nm$, A As Drs, Optional MaxColWdt% = 100) As String()
LyzNmDrs = LyzNmLy(Nm, FmtDrs(A, MaxColWdt), EiNoIx)
End Function

Function LyzNmLy(Nm$, Ly$(), Optional B As EmIxCol = EiBeg1) As String()
Dim L$(), J&, S$
If Si(Ly) = 0 Then
    PushI LyzNmLy, Nm & "(No Lin)"
    Exit Function
End If
L = AddIxPfx(Ly, B)
'Brw L:Stop
S = Space(Len(Nm))
PushI LyzNmLy, Nm & L(0)
For J = 1 To UB(L)
    PushI LyzNmLy, S & L(J)
Next
End Function

Function LyzMsg(Msg$) As String()
LyzMsg = LyzFunMsg("", Msg)
End Function
Function LyzMsgNav(Msg$, Nav()) As String()
LyzMsgNav = AddAy(LyzMsg(Msg), IndentSy(LyzNav(Nav)))
End Function

Function LyzNNAp(NN$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
LyzNNAp = LyzNNAv(NN, Av)
End Function

Function LyzNNAv(NN$, Av()) As String()
LyzNNAv = LyzNyAv(Ny(NN), Av)
End Function

Function LinzFunMsgNav$(Fun$, Msg$, Nav())
LinzFunMsgNav = LinzFunMsg(Fun, Msg) & " " & LinzNav(Nav)
End Function

Sub DmpNNAp(NN$, ParamArray Ap())
Dim Av(): Av = Ap
D LyzNyAv(Ny(NN), Av)
End Sub

Function LyzNyAv(Ny$(), Av(), Optional Sep$ = ": ") As String()
Dim J%, O$(), N$()
ResiMax Ny, Av
N = AlignAy(Ny)
For J = 0 To UB(Ny)
    PushIAy LyzNyAv, LyzNv(N(J), Av(J), Sep)
Next
End Function

Function LinzLyzMsgNav$(Msg$, Nav())
LinzLyzMsgNav = EnsSfxDot(Msg) & " | " & LinzNav(Nav)
End Function

Function LinzNav$(Nav())
Dim Ny$(), Av()
AsgNyAv Nav, Ny, Av
LinzNav = LinzNyAv(Ny, Av)
End Function

Function LinzNyAv$(Ny$(), Av())
Dim J%, U1%, U2%, N$, V$, O$()
U1 = UB(Ny)
U2 = UB(Av)
For J = 0 To Max(U1, U2)
    If J <= U1 Then N = QteSq(Ny(J)) Else N = "[?]"
    If J <= U2 Then V = Cell(Av(J)) Else V = "?"
    PushI O, N & " " & V
Next
LinzNyAv = JnVbarSpc(O)
End Function

Sub AsgNyAv(Nav(), ONy$(), OAv())
If Si(Nav) = 0 Then
    Erase ONy
    Erase OAv
    Exit Sub
End If
Dim TT$: TT = Nav(0)
ONy = TermAy(TT)
OAv = AeFstEle(Nav)
End Sub
Private Sub Z_LyzNav()
Dim Nav(): Nav = Array("aa bb", 1, 2)
D LyzNav(Nav)
End Sub

Function LyzNav(Nav()) As String()
Dim Ny$(), Av()
AsgNyAv Nav, Ny, Av
LyzNav = LyzNyAv(Ny, Av)
End Function

Function SclzNyAv$(Ny$(), Av())
SclzNyAv = JnSemi(LyzNyAv(Ny, Av))
End Function

Private Function LyzFunMsg(Fun$, Msg$) As String()
Dim O$(), MsgL1$, MsgRst$
AsgBrk1Dot Msg, MsgL1, MsgRst
PushI LyzFunMsg, EnsSfxDot(MsgL1) & IIf(Fun = "", "", "  @" & Fun)
PushIAy LyzFunMsg, IndentSy(WrpLy(SplitCrLf(MsgRst)))
End Function



