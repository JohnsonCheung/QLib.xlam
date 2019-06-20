Attribute VB_Name = "QIde_Src_Syc"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_ConstMth."
Private Const Asm$ = "QIde"
Public Const DoczSycv$ = "It is Sy.  It comes from the MthLy or from SycFt"
Public Const DoczSycm$ = "Sy-Const-MthL.  It is MthL in a any mod or cls."
Public Const DoczSyc$ = "Sy-Const.  Each Module/Class may have som fun of type String() have cxt as Erase XX|X ..|nn=XX|Erase XX"
Public Const DoczSycn$ = "Sy-Const-Nm.  It is same as mthn"
Public Const DoczSycFt$ = "Sy-Const-Ft.  It comes from Sycn & Mdn"
Private CnstnInEdt$, MdnInEdt$
Function AA() As String()
Erase XX
X "dklfjsdl fkj sdlfkjsdl f"
X "sdfkljsd fksdljf"
X "skldfj "
X ""
AA = XX
End Function

Function IsSycm(MthLy$()) As Boolean
Dim L$
L = MthLy(0): If MthTy(L) <> "Function" Then Exit Function
If Not HasSfx(L, " As String()") Then Exit Function
If BetBkt(L) <> "" Then Exit Function
L = MthLy(1): If L <> "Erase XX" Then Exit Function
Dim U&: U = UB(MthLy)
L = MthLy(U): If L <> "End Function" Then Exit Function
L = MthLy(U - 1): If L <> "Erase XX" Then Exit Function
End Function
Private Function Sycm(M As CodeModule, Mthn$) As String()
Dim O$(): O = MthLyzM(M, Mthn)
If Si(O) = 0 Then Exit Function
If Not IsSycm(O) Then Thw CSub, "Given mthn is not a Sycm", "Mdn Mthn MthL", Mdn(M), Mthn, O
Sycm = O
End Function

Private Function SycmzLy$(Ly$(), Mthn$, Optional IsPub As Boolean)
Erase XX
Const C$ = "?Function ?() As String()"
Dim Mdy$: If Not IsPub Then Mdy = "Private "
X FmtQQ(C, Mdy, Mthn)
X "Erase XX"
Dim I: For Each I In Itr(Ly)
    Dim L$: L = "X """ & Replace(I, vbDblQte, vb2DblQte) & """"
    X L
Next
X Mthn & " = XX"
X "End Function"
SycmzLy = JnCrLf(XX)
End Function
Function Sycv(M As CodeModule, Mthn) As String()
Sycv = SycvzMthLy(MthLyzM(M, Mthn))
End Function

Function SycvzMthLy(MthLy$()) As String()
'Fm SycMthLy :
If IsSycm(MthLy) Then Thw CSub, "Given MthLy is not Syc.  SycMth must 1. Return String() 2. Ctx is [Erase|X..|Mthn=XX|Erase]", "SycMthLy", MthLy
Dim L$(): L = AyeLasNEle(AyeFstNEle(MthLy, 2), 2)
Stop
Erase XX
Dim I: For Each I In Itr(L)
    Dim A$: A = Replace(RmvLasChr(Mid(I, 4)), vb2DblQte, vbDblQte)
    X A
Next
SycvzMthLy = XX
End Function

Function TakVbStr$(VbStr$)
If FstChr(VbStr) <> """" Then Thw CSub, "FstChr of VbStr must be DblQte", "VbStr", VbStr
Dim P%: P = InStr(2, VbStr, """")
If P = 0 Then Thw CSub, "There is no ending DblQte", "VbStr", VbStr
TakVbStr = Mid(VbStr, 2, P - 2)
End Function
Function TakVbStrzSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI TakVbStrzSy, TakVbStr(CStr(I))
Next
End Function

Private Sub Z_Sycv()
Const TstId& = 3
Const CSub$ = CMod & "Z_CnstBrkMthL"
Dim MthL$, Cas$, IsEdt As Boolean
GoSub T0
GoSub T1
Exit Sub
T0:
    IsEdt = False
    Cas = "Complex"
    MthL = TstTxt(TstId, CSub, Cas, "MthL", IsEdt:=True)
    Ept = TstTxt(TstId, CSub, Cas, "Ept", IsEdt)
    If IsEdt Then Return
    GoTo Tst
T1:
   
    Return
Tst:
'    Act = SycVal(MthL)
    Brw Act
    Stop
    C
    Return
End Sub

Function CnstBrkzMd1$(M As CodeModule, SycNm$)
Dim J%, L$, O$
For J = 1 To M.CountOfDeclarationLines
    L = M.Lines(J, 1)
    O = CnstBrkzLinNm(L, SycNm)
    If O <> "" Then CnstBrkzMd1 = O: Exit Function
Next
End Function

Function CnstBrkzLinNm$(Lin, SycNm)
Dim L$: L = RmvMthMdy(Lin)
If Not ShfPfx(L, "Const ") Then Exit Function
If ShfNm(L) <> SycNm Then Exit Function
If ShfTyChr(L) = "$" Then Thw CSub, "Given constant name is found, but is not a Str", "ConstLin SycNm", Lin, SycNm
Dim O$: O = Bet(L, """", """")
If O = "" Then Thw CSub, "Between DblQte is nothing", "ConstLin SycNm", Lin, SycNm
CnstBrkzLinNm = O
End Function

Function CnstLy(Src$()) As String()
Dim S$(): S = Src
Dim J%
While Si(S) > 0
    J = J + 1: If J > 10000 Then ThwLoopingTooMuch CSub
    PushNB CnstLy, ShfCnstLin(S)
Wend
End Function

Function DCnst(Src$()) As Drs
Dim Dry()
Dim Ly$(): Ly = CnstLy(Src)
Dim L: For Each L In Itr(Ly)
Next
DCnst = DrszFF("Mdy Cnstn TyChr Lin", Dry)
End Function
Function Cnstn$(Lin)
Dim L$: L = Lin
ShfMdy L
If ShfPfx(L, "Const") Then Cnstn = TakNm(L)
End Function

Function CnstLnozMN(M As CodeModule, Cnstn$) As Lnx
Dim J&, L$
For J = 1 To M.CountOfDeclarationLines
    L = M.Lines(J, 1)
    If HasPfx(L, "Const CMod$") Then
        CnstLnozMN = Lnx(L, J - 1)
        Exit Function
    End If
Next
End Function

Function ShfTermCnst(OLin$) As Boolean
ShfTermCnst = ShfTerm(OLin, "Const")
End Function

Private Sub Z_HasCnstn()
Debug.Assert HasCnstn(CMd, "CMod")
End Sub
Function HasCnstn(M As CodeModule, Cnstn$) As Boolean
Dim J%
For J = 1 To M.CountOfDeclarationLines
    If HitCnstn(M.Lines(J, 1), Cnstn) Then HasCnstn = True: Exit Function
Next
End Function

Function CnstnzL$(L)
Dim Lin$: Lin = RmvMdy(L)
If ShfTermCnst(Lin) Then CnstnzL = Nm(LTrim(Lin))
End Function

Function DrzStrCnst(Lin) As Variant()
Dim L$: L = RmvMdy(Lin)
If Not ShfCnst(L) Then Exit Function
Dim N$: N = ShfNm(L): If N = "" Then Exit Function
If Not ShfPfx(L, "$") Then Exit Function
If Not ShfPfx(L, " = """) Then Exit Function
Dim P%: P = InStr(L, """"): If P = 0 Then Stop
DrzStrCnst = Array(N, Left(L, P - 1))
End Function
Function StrValzCnstn$(Lin, Cnstn$)
Dim L$: L = RmvMdy(Lin)
If Not ShfCnst(L) Then Exit Function
If ShfNm(L) <> Cnstn$ Then Exit Function
If Not ShfPfx(L, "$") Then Exit Function
If Not ShfPfx(L, " = """) Then Stop
Dim P%: P = InStr(L, """"): If P = 0 Then Stop
StrValzCnstn = Left(L, P - 1)
End Function
Function DStrCnstP() As Drs
DStrCnstP = DStrCnst(SrczP(CPj))
End Function
Function DStrCnst(Src$()) As Drs
Dim ODry(), L
For Each L In Itr(Src)
    PushISomSi ODry, DrzStrCnst(L)
Next
DStrCnst = DrszFF("Cnstn StrVal", ODry)
End Function

Function StrValzCnstLy$(Ly$(), Cnstn$)
Dim L
For Each L In Itr(Ly)
    Dim O$: O = StrValzCnstn(L, Cnstn)
    If O <> "" Then StrValzCnstLy = O: Exit Function
Next
End Function

Function StrValzCnstLin(Lin)
Stop
'StrValzCnstLin = StrValzCnstBrk(CnstBrk(Lin))
End Function

Function CMCnstLy(CmSrc$()) As String()
Dim L
For Each L In Itr(CmSrc)
PushI CMCnstLy, CMCnstLin(L)
Next
End Function
Function CMCnstLin$(CMSrcLin)
Dim N, T1$, L$, O$
L = CMSrcLin
T1 = ShfT1(L)
O = "Private Const C_" & T1 & "$ = """ & L
For Each N In NyzMacro(CMSrcLin)
    O = Replace(O, QteBigBkt(N), "?")
Next
CMCnstLin = O & """"
End Function

Function CMFunLinesAy(CmSrc$()) As String()
Dim L
For Each L In Itr(CmSrc)
PushI CMFunLinesAy, CMFunLines(L)
Next
End Function
Function CMFunLines$(CMSrcLin)
If InStr(CMSrcLin, "{") = 0 Then Exit Function
Dim O$(), Nm$, Pm$, PmOnlyNm$, Ny$(), NyOnlyNm$()
Nm = T1(CMSrcLin)
Ny = AywDist(NyzMacro(CMSrcLin))
Pm = JnComma(Ny)
'NyOnlyNm = TakNm zAy(Ny)
PmOnlyNm = JnComma(NyOnlyNm)
PushI O, FmtMacro("Private Function M_{Nm}$({Pm})", Nm, Pm)
PushI O, FmtMacro("M_{Nm} = FmtQQ(C_{Nm}, {PmNmOnly})", Nm, PmOnlyNm)
PushI O, "End Function"
End Function


Sub EdtCnst(Cnstn$)
CnstnInEdt = Cnstn
MdnInEdt = CMdn
ExpCnstValInEdt
BrwFt EnsFt(CnstFtInEdt)
End Sub

Private Sub ExpCnstValInEdt()
ExpCnstVal MdnInEdt, CnstnInEdt
End Sub

Private Sub ExpCnstVal(Mdn$, Cnstn$)
Dim M As CodeModule: Set M = Md(Mdn)
If Not HasMthzM(M, Cnstn) Then
    Debug.Print "Cnst mth not fnd"
    Exit Sub
End If
Dim MLy$(): MLy = MthLyzM(M, Cnstn)
If IsSycm(MLy) Then
    Debug.Print "Not a cnst mth:"
    D MLy
    Exit Sub
End If
Dim Ft$: Ft = CnstFt(Mdn, Cnstn)
WrtAy SycvzMthLy(MLy), Ft
End Sub

Private Function CnstFt$(Mdn$, Mthn$)
If Mdn = "" Then Debug.Print "Mdn is blank": Exit Function
If Mthn = "" Then Debug.Print "Cnstn is blank": Exit Function
Dim H$: H = TmpFdr("Cnst")
Dim A$: A = AddFdrEns(H, Mdn)
CnstFt = A & Mthn & ".txt"
End Function

Private Function CnstFtInEdt$()
CnstFtInEdt = CnstFt(MdnInEdt, CnstnInEdt)
End Function
Sub ImpCnst(Optional IsPub As Boolean)
ImpCnstzN MdnInEdt, CnstnInEdt
End Sub
Sub ImpCnstzN(Mdn$, Mthn$, Optional IsPub As Boolean)
Dim Ft$: Ft = CnstFt(Mdn, Mthn)
If Not HasFfn(Ft) Then
    Debug.Print "CnstFt  not found"
    Debug.Print "Mdn     : "; Mdn
    Debug.Print "Cnstn   : "; Mthn
    Debug.Print "CnstFt  : "; Ft
    Exit Sub
End If
Dim M As CodeModule: Set M = Md(Mdn)
Dim FmFt$(): FmFt = LyzFt(Ft)
Dim NewL$: ' NewL = Sycm(Mthn, FmFt, IsPub)
RplMth M, Mthn, NewL
End Sub

Private Property Get Z_CrtSchm1() As String()
Erase XX
X "Tbl A *Id | *Nm     | *Dte AATy Loc Expr Rmk"
X "Tbl B *Id | AId *Nm | *Dte"
X "Fld Txt AATy"
X "Fld Loc Loc"
X "Fld Expr Expr"
X "Fld Mem Rmk"
X "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
X "Ele Expr Txt [Expr=Loc & 'abc']"
X "Des Tbl     A     AA BB "
X "Des Tbl     A     CC DD "
X "Des Fld     ANm   AA BB "
X "Des Tbl.Fld A.ANm TF_Des-AA-BB"
Z_CrtSchm1 = XX
Erase XX
End Property
Function FtzCnstQNm$(CnstQNm$)
Dim Mdn, Nm$
FtzCnstQNm = ConstPrpPth(Mdn) & Nm & ".txt"
End Function
Private Function ConstPrpPth$(Mdn)
ConstPrpPth = AddFdrEns(TmpHom, "ConstPrp", Mdn)
End Function

Function IsLinCnstStr(Lin) As Boolean
If Not IsLinMth(Lin) Then Exit Function
If BetBkt(Lin) <> "" Then Exit Function
If TakTyChr(Lin) = "$" Then Exit Function
IsLinCnstStr = True
End Function

