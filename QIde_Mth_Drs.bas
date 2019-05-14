Attribute VB_Name = "QIde_Mth_Drs"
Option Explicit
Private Const CMod$ = "MIde_Mth_Drs."
Private Const Asm$ = "QIde"

Function MthDrszFb(Fb, Optional WhStr$) As Drs
MthDrszFb = MthDrszV(VbezPjf(Fb), WhStr)
ClsPjf Fb
End Function

Function MthDrsInVbe(Optional WhStr$) As Drs
MthDrsInVbe = MthDrszV(CVbe, WhStr)
End Function

Function MthDrszFxa(Fxa$, Optional WhStr$, Optional Xls As Excel.Application) As Drs
Dim A As Excel.Application: Set A = DftXls(Xls)
MthDrszFxa = MthDrszP(PjzFxa(Fxa), WhStr)
If IsNothing(Xls) Then QuitXls Xls
End Function

Function MthDrszMd(A As CodeModule, Optional WhStr$) As Drs
MthDrszMd = Drs(MthFny, MthDryzMd(A, WhStr))
End Function

Function MthDrszP(P As VBProject, Optional WhStr$) As Drs
Dim O As Drs
O = Drs(MthFny, MthDryzP(P, WhStr))
O = AddColzValIdzCntzDrs(O, "Lines", "Pj")
O = AddColzValIdzCntzDrs(O, "Nm", "PjMth")
MthDrszP = O
End Function

Function MthDrszPjf(Pjf$, Optional WhStr$) As Drs
Dim V As Vbe, App, P As VBProject, PjDte As Date
OpnPjf Pjf ' Either Excel.Application or Access.Application
Set V = VbezPjf(Pjf)
Set P = PjzPjf(V, Pjf)
Select Case True
Case IsFb(Pjf):  PjDte = PjDtezAcs(CvAcs(App))
Case IsFxa(Pjf): PjDte = DtezFfn(Pjf)
Case Else: Stop
End Select
MthDrszPjf = DrsAddCol(MthDrszP(P), "PjDte", PjDte)
If IsFb(Pjf) Then
    CvAcs(App).CloseCurrentDatabase
End If
End Function

Function MthDrszPjfSy1(Optional WhStr$) As Drs
MthDrszPjfSy1 = MthDrszPjfSy(PjfSy, WhStr)
End Function

Function MthDrszPjfSy(PjfSy$(), Optional WhStr$) As Drs
Dim I
For Each I In PjfSy
    ApdDrs MthDrszPjfSy, MthDrszPjf(CStr(I), WhStr)
Next
End Function

Function MthDrszV(A As Vbe, Optional WhStr$) As Drs
MthDrszV = Drs(MthFny, MthDryzV(A, WhStr))
End Function

Function MthDryzMd(A As CodeModule, Optional WhStr$) As Variant()
Dim P$, T$, M$
P = PjnzM(A)
T = ShtCmpTyzMd(A)
M = Mdn(A)
MthDryzMd = DryInsColzV3(MthDryzSrc(Src(A), WhStr), P, T, M)
End Function

Function MthDryzP(P As VBProject, Optional WhStr$) As Variant()
Dim M
For Each M In MdItr(P, WhStr)
    PushIAy MthDryzP, MthDryzMd(CvMd(M), WhStr)
Next
End Function

Function MthDryzSrc(Src$(), Optional WhStr$) As Variant()
Dim MthLin
For Each MthLin In Itr(MthLinyzSrc(Src, WhStr))
    PushI MthDryzSrc, DrOfMthLin(CStr(MthLin))
Next
End Function

Function MthDryzV(A As Vbe, Optional WhStr$) As Variant()
Dim P
For Each P In PjItr(A, WhStr)
    PushIAy MthDryzV, MthDryzP(CvPj(P), WhStr)
Next
End Function

Function MthInfAy_Md(A As CodeModule) As MthInf()
'MthInfAy_Md = MthInfAy_Src(PjnMth(A), Mdn(A), CurSrc(A))
End Function

Function MthInfAy_MdzPjSrc(Pjn$, Mdn, Src$()) As MthInf()
Dim Ix&():  Ix = MthIxy(Src)
Dim I
For Each I In Itr(Ix)
    PushObj MthInfAy_MdzPjSrc, MthInfMdzPjSrcFm(Pjn, Mdn, Src, CLng(I))
Next
End Function

Function MthInfAy_Pj(P As VBProject) As MthInf()
Dim C As VBComponent
For Each C In P.VBComponents
    PushObjAy MthInfAy_Pj, MthInfAy_Md(C.CodeModule)
Next
End Function

Function MthInfAyzV(A As Vbe) As MthInf()
Dim P As VBProject
For Each P In A.VBProjects
    PushObjAy MthInfAyzV, MthInfAy_Pj(P)
Next
End Function

Function MthInfMdzPjSrcFm(Pjn$, Mdn, Src$(), FmIx&) As MthInf
Dim O As New MthInf, L$
O.MthLin = ContLin(Src, FmIx)
L = Src(FmIx)
'O.ShtMdy = ShtMdy(ShfMdy(L))
O.ShtKd = ShtMthKd(ShfMthTy(L))
Set MthInfMdzPjSrcFm = O
End Function

Function DrOfMthzSI(Src$(), MthIx) As Variant()
Dim L$, Lines$, Rmk$(), Lno, Cnt%
    L = ContLin(Src, MthIx)
    Lno = MthIx + 1
    Lines = MthLineszSIW(Src, MthIx)
    Cnt = LinCnt(Lines)
    Rmk = TopRmkLy(Src, MthIx)
Dim Dr():  Dr = DrOfMthLin(L): If Si(Dr) = 0 Then Stop
DrOfMthzSI = AddAy(Dr, Array(Lno, Cnt, Lines, Rmk))
End Function

Function DrOfMthLin(MthLin, Optional B As WhMth) As Variant()
If Not HitMthLin(MthLin, B) Then Exit Function
Dim X As MthLinRec
X = MthLinRec(MthLin)
With X
DrOfMthLin = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function
Function MthDr(Src$(), MthLin, MthIx, Optional B As WhMth) As Variant()
If Not HitMthLin(MthLin, B) Then Exit Function
Dim X As MthLinRec
X = MthLinRec(MthLin)
With X
MthDr = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function

Function DrOfMthLinsP(Optional WhStr$) As Drs
DrOfMthLinsP = DrOfMthLinszP(CPj, WhStr)
End Function

Function DrOfMthLinszP(P As VBProject, Optional WhStr$) As Drs
DrOfMthLinszP = Drs(MthLinFny, DryOfMthLinzP(CPj, WhStr))
End Function

Function DryOfMthLinzM(A As CodeModule, Optional WhStr$) As Variant()
Dim P$, T$, M$
P = PjnzM(A)
T = ShtCmpTyzMd(A)
M = Mdn(A)
DryOfMthLinzM = DryInsColzV3(DryOfMthLinzS(Src(A)), P, T, M)
End Function

Function DryOfMthLinzP(P As VBProject, Optional WhStr$) As Variant()
Dim M
For Each M In MdItr(P, WhStr)
    PushAy DryOfMthLinzP, DryOfMthLinzM(CvMd(M), WhStr)
Next
End Function

Function DryOfMthLinzS(Src$(), Optional WhStr$) As Variant()
Dim MthLin, W As WhMth
Set W = WhMthzStr(WhStr)
For Each MthLin In Itr(MthLinyzSrc(Src))
    PushISomSi DryOfMthLinzS, DrOfMthLin(CStr(MthLin), W)
Next
End Function

Function MthWb(Optional WhStr$) As Workbook
Set MthWb = ShwWb(MthWbPjfSy(PjfSy, WhStr))
End Function

Function MthWbFmt(A As Workbook) As Workbook
Dim Ws As Worksheet, Lo As ListObject
Set Ws = WszCdNm(A, "MthLoc"): If IsNothing(Ws) Then Stop
Set Lo = LozWs(Ws, "T_MthLoc"): If IsNothing(Lo) Then Stop
Dim Ws1 As Worksheet:  GoSub X_Ws1
Dim Pt1 As PivotTable: GoSub X_Pt1
Dim Lo1 As ListObject: GoSub X_Lo1
Dim Pt2 As PivotTable: GoSub X_Pt2
Dim Lo2 As ListObject: GoSub X_Lo2
Ws1.Outline.ShowLevels , 1
Set MthWbFmt = WbzWs(Ws)
Exit Function
X_Ws1:
    Set Ws1 = AddWs(WbzWs(Ws))
    Ws1.Outline.SummaryColumn = xlSummaryOnLeft
    Ws1.Outline.SummaryRow = xlSummaryBelow
    Return
X_Pt1:
'    Set Pt1 = PtzLo(Lo, A1zWs(Ws1), "MdTy Nm VbeLinesId Lines", "Pj")
    SetPtOutLin Pt1, "Lines"
    SetPtWdt Pt1, "VbeLinesId", 12
    SetPtWdt Pt1, "Nm", 30
    SetPtRepeatLbl Pt1, "MdTy Nm"
    Return
X_Lo1:
    Set Lo1 = PtCpyToLo(Pt1, Ws1.Range("G1"))
    Erase XX
    X "Nm T_MthLines"
    X "Wdt 30 Nm"
    X "Wdt 100 Lines"
    X "Lvl 2 Lines"
    FmtLo Lo1, XX
    Erase XX
    Return
X_Pt2:
    Set Pt2 = PtzLo(Lo1, Ws1.Range("M1"), "MdTy Nm", "Lines")
    SetPtRepeatLbl Pt2, "MdTy"
    Return
X_Lo2:
    Set Lo2 = PtCpyToLo(Pt2, Ws1.Range("Q1"))
    SetLoNm Lo2, "T_UsrEdtMthLoc"
    Return
Set MthWbFmt = A
End Function

Function MthWbPjfSy(PjfSy$(), Optional WhStr$) As Workbook
Set MthWbPjfSy = MthWbFmt(WbzWs(MthWszPjfSy(PjfSy, WhStr)))
End Function

Function MthWsInVbe(Optional WhStr$) As Worksheet
Set MthWsInVbe = MthWszV(CVbe, WhStr)
End Function

Function MthWszP(P As VBProject, Optional WhStr$, Optional Vis As Boolean) As Worksheet
Set MthWszP = SetVisOfWs(WszDrs(MthDrszP(P, WhStr)), Vis)
End Function

Function MthWszPjfSy(PjfSy$(), Optional WhStr$) As Worksheet
Dim O As Drs
O = MthDrszPjfSy(PjfSy, WhStr)
O = AddColzValIdzCntzDrs(O, "Nm", "Vbe_Mth")
O = AddColzValIdzCntzDrs(O, "Lines", "Vbe")
Dim Ws As Worksheet
Set Ws = WszDrs(O)
SetWsCdNmAndLoNm Ws, "MthFull"
Set MthWszPjfSy = Ws
End Function

Function MthWszV(A As Vbe, Optional WhStr$) As Worksheet
Set MthWszV = WszDrs(MthDrszV(A, WhStr))
End Function

Function PjnMth$(A As CodeModule)
PjnMth = A.Parent.Collection.Parent.Name
End Function

Function MthDrszVAy(A() As Vbe) As Drs
Dim I, R%, M As Drs
For Each I In Itr(A)
    M = DrsInsCV(MthDrszV(CvVbe(I)), "Vbe", R)
    If R = 0 Then
        MthDrszVAy = M
    Else
        ApdDrs MthDrszVAy, M
    End If
    R = R + 1
    Debug.Print R; "<=== MthDrszVAy"
Next
End Function

Function MthWszVAy(A() As Vbe) As Worksheet
'Set MthWszVAy = WszDrs(MthDrszVAy(A))
End Function

Property Get MthFny() As String()
MthFny = SyzSS("Pj MdTy Md Mdy Ty Nm Ret Pm Rmk Lno Cnt Lines TopRmk")
End Property

Property Get MthLinFny() As String()
MthLinFny = SyzSS("Pj MdTy Md Mdy Ty Nm Ret Pm Rmk")
End Property

Property Get MthWs() As Worksheet
Set MthWs = MthWszPjfSy(PjfSy)
End Property

Property Get PjfSy() As String()
End Property

Private Property Get ZzVAy() As Vbe()
PushObj ZzVAy, CVbe
Const Fb$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
'PushObj ZzVAy, AcsOpnFb(Fb).Vbe
End Property

Private Sub ZZ()
Dim A As WhPjMth
Dim B As Variant
Dim C As WhMdMth
Dim D As CodeModule
Dim E As WhMth
Dim F$
Dim G As Workbook
Dim H$()
Dim I As VBProject
Dim J&
Dim K() As Vbe
Dim L As Vbe
End Sub

Private Sub Z_MthDrszMd()
BrwDrs MthDrszMd(CMd)
End Sub

Private Sub Z_MthDrszPjf()
Dim Pjf$
Pjf = PjfSy()(0)
ShwWs WszDrs(MthDrszPjf(Pjf))
End Sub

Private Sub Z_DryOfMthLinzP()
Dim A(): A = DryOfMthLinzP(CPj)
Stop
End Sub

Private Sub Z_DrOfMthLinyzV()
BrwDry DrOfMthLinyzV(CVbe)
End Sub

Private Sub Z_MthWb()
ShwWb MthWb
End Sub

Private Sub Z_MthWbFmt()
Dim Wb As Workbook
Const Fx$ = "C:\Users\user\Desktop\Vba-Lib-1\Mth.xlsx"
MthWbFmt ShwWb(WbzFx(Fx))
Stop
End Sub

Private Sub Z_MthWszVAy()
ShwWs MthWszVAy(ZzVAy)
End Sub
