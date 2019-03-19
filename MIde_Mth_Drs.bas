Attribute VB_Name = "MIde_Mth_Drs"
Option Explicit
Function MthInfAyzbe(A As Vbe) As MthInf()
Dim P As VBProject
For Each P In A.VBProjects
    PushObjzAy MthInfAyzbe, MthInfAy_Pj(P)
Next
End Function
Function MthInfAy_Pj(A As VBProject) As MthInf()
Dim C As VBComponent
For Each C In A.VBComponents
    PushObjzAy MthInfAy_Pj, MthInfAy_Md(C.CodeModule)
Next
End Function
Function MthInfAy_Md(A As CodeModule) As MthInf()
'MthInfAy_Md = MthInfAy_Src(PjNmMth(A), MdNm(A), SrcMd(A))
End Function
Function MthInfAy_MdzPjSrc(PjNm$, MdNm$, Src$()) As MthInf()
Dim Ix&():  Ix = MthIxAy(Src)
Dim I
For Each I In Itr(Ix)
    PushObj MthInfAy_MdzPjSrc, MthInfMdzPjSrcFm(PjNm, MdNm, Src, I)
Next
End Function

Function PjNmMth$(A As CodeModule)
PjNmMth = A.Parent.Collection.Parent.Name
End Function

Function MthInfMdzPjSrcFm(PjNm$, MdNm$, Src$(), FmIx) As MthInf
Dim O As New MthInf, L$
O.MthLin = ContLin(Src, FmIx)
L = Src(FmIx)
'O.ShtMdy = ShtMdy(ShfMthMdy(L))
O.ShtKd = ShtMthKd(ShfMthTy(L))
Set MthInfMdzPjSrcFm = O
End Function

Function MthDrs(Optional WhStr$) As Drs
Set MthDrs = MthDrszPjfAy(PjfAy, WhStr)
End Function


Function MthDrsMd(A As CodeModule, Optional B As WhMth) As Drs
Set MthDrsMd = Drs(MthFny, MthDryzMd(A, B))
End Function

Function MthLinDryzMd(A As CodeModule, Optional WhStr$) As Variant()
Dim P As VBProject, Ffn$, Pj$, Ty$, Md$, MdTy$
Set P = PjNmzMd(A)
Ffn$ = Pjf(P)
Pj = P.Name
MdTy = ShtCmpTyzMd(A)
Md = MdNm(A)
MthLinDryzMd = DryInsColz4V(MthLinDryzSrc(Src(A)), Ffn, Pj, MdTy, Md)
End Function
Function MthLinDryzSrc(Src$(), Optional WhStr$) As Variant()
Dim MthLin, W As WhMth
Set W = WhMthzStr(WhStr)
For Each MthLin In Itr(MthLinAyzSrc(Src))
    PushISomSz MthLinDryzSrc, MthLinDr(MthLin, W)
Next
End Function

Function MthDryzMd(A As CodeModule, Optional B As WhMth) As Variant()
Dim P As VBProject, Ffn$, Pj$, Ty$, Md$, MdTy$
Set P = PjNmzMd(A)
Ffn$ = Pjf(P)
Pj = P.Name
MdTy = ShtCmpTyzMd(A)
Md = MdNm(A)
MthDryzMd = DryInsColz4V(MthDryzSrc(Src(A)), Ffn, Pj, MdTy, Md)
End Function

Property Get PjfAy() As String()
End Property

Function MthWb(Optional WhStr$) As Workbook
Set MthWb = WbVis(MthWbPjfAy(PjfAy, WhStr))
End Function

Property Get MthWs() As Worksheet
Set MthWs = MthWsPjfAy(PjfAy)
End Property

Function MthDrsFb(Fb, Optional WhStr$) As Drs
Set MthDrsFb = MthDrszVbe(VbePjf(Fb), WhStr)
ClsPjf Fb
End Function

Function MthDrszFxa(Fxa, Optional WhStr$, Optional Xls As Excel.Application) As Drs
Dim A As Excel.Application: Set A = DftXls(Xls)
Set MthDrszFxa = MthDrszPj(PjzFxa(Fxa, A), WhStr)
If IsNothing(Xls) Then XlsQuit Xls
End Function

Function MthWbPjfAy(PjfAy$(), Optional WhStr$) As Workbook
Set MthWbPjfAy = MthWbFmt(WbzWs(MthWsPjfAy(PjfAy, WhStr)))
End Function

Function MthWsPjfAy(PjfAy, Optional WhStr$) As Worksheet
Dim O As Drs
Set O = MthDrszPjfAy(PjfAy, WhStr)
Set O = AddColzValIdzCntzDrs(O, "Nm", "Vbe_Mth")
Set O = AddColzValIdzCntzDrs(O, "Lines", "Vbe")
'Set MthWsPjfAy = WsSetCdNmAndLoNm(WszDrs(O), "MthFull")
End Function


Function MthDrszPjf(Pjf, Optional WhStr$) As Drs
Dim V As Vbe, App, P As VBProject, PjDte As Date
OpnPjf Pjf ' Either Excel.Application or Access.Application
Set V = VbePjf(Pjf)
Set P = PjzPjfVbe(V, Pjf)
Select Case True
Case IsFb(Pjf):  PjDte = AcsPjDte(CvAcs(App))
Case IsFxa(Pjf): PjDte = FfnDte(Pjf)
Case Else: Stop
End Select
Set MthDrszPjf = DrsAddCol(MthDrszPj(P), "PjDte", PjDte)
If IsFb(Pjf) Then
    CvAcs(App).CloseCurrentDatabase
End If
End Function


Function MthDrszPj(A As VBProject, Optional WhStr$) As Drs
Dim O As Drs
Set O = Drs(MthFny, MthDryzPj(A, WhStr))
Set O = AddColzValIdzCntzDrs(O, "Lines", "Pj")
Set O = AddColzValIdzCntzDrs(O, "Nm", "PjMth")
Set MthDrszPj = O
End Function

Function MthDryzPj(A As VBProject, Optional WhStr$) As Variant()
Dim M, W As WhMth
Set W = WhMthzStr(WhStr)
For Each M In MdItr(A, WhStr)
    PushIAy MthDryzPj, MthDryzMd(CvMd(M), W)
Next
End Function

Function MthWsPj(A As VBProject, Optional WhStr$) As Worksheet
Set MthWsPj = WszDrs(MthDrszPj(A, WhStr))
End Function

Private Sub Z_MthWb()
ShwWb MthWb
End Sub

Private Sub Z_MthDrsMd()
BrwDrs MthDrsMd(CurMd)
End Sub

Private Sub Z_MthWbFmt()
Dim Wb As Workbook
Const Fx$ = "C:\Users\user\Desktop\Vba-Lib-1\Mth.xlsx"
MthWbFmt WbVis(WbzFx(Fx))
Stop
End Sub

Function MthDrszPjfAy(PjfAy, Optional WhStr$) As Drs
Dim I
For Each I In PjfAy
    PushDrs MthDrszPjfAy, MthDrszPjf(I, WhStr)
Next
End Function

Private Sub Z_MthDRszPjf()
Dim Pjf$
Pjf = PjfAy()(0)
ShwWs WszDrs(MthDrszPjf(Pjf))
End Sub

Function MthDrszVbe(A As Vbe, Optional WhStr$) As Drs
Dim P, Dry()
For Each P In PjItr(A, WhStr)
    PushIAy Dry, MthDryzPj(CvPj(P), WhStr)
Next
Set MthDrszVbe = Drs(MthFny, Dry)
End Function

Function MthWszVbe(A As Vbe, Optional WhStr$) As Worksheet
Set MthWszVbe = WszDrs(MthDrszVbe(A, WhStr))
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
'    Set Pt1 = NewPtLoAtRDCP(Lo, A1zWs(Ws1), "MdTy Nm VbeLinesId Lines", "Pj")
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
    Set Pt2 = NewPtLoAtRDCP(Lo1, Ws1.Range("M1"), "MdTy Nm", "Lines")
    SetPtRepeatLbl Pt2, "MdTy"
    Return
X_Lo2:
    Set Lo2 = PtCpyToLo(Pt2, Ws1.Range("Q1"))
    LoSetNm Lo2, "T_UsrEdtMthLoc"
    Return
Set MthWbFmt = A
End Function

Function MthLinDryzPj(A As VBProject, Optional WhStr$) As Variant()
Dim M
For Each M In MdItr(A, WhStr)
    PushAy MthLinDryzPj, MthLinDryzMd(CvMd(M), WhStr)
Next
End Function

Function MthLinDr(MthLin, Optional B As WhMth) As Variant()
Dim L$, Mdy$, Ty$, Nm$, Prm$, Ret$, TopRmk$, LinRmk$
L = MthLin
Mdy = ShfMthMdy(L)
Ty = ShfShtMthTy(L): If Ty = "" Then Exit Function
Nm = ShfNm(L)
Ret = ShfMthSfx(L)
Prm = ShfBktStr(L)
If ShfX(L, "As") Then
    If Ret <> "" Then Stop
    Ret = ShfTerm(L)
End If
If ShfPfx(L, "'") Then
    LinRmk = L
End If
MthLinDr = Array(Mdy, Ty, Nm, Prm, Ret, LinRmk)
End Function

Function MthDryzSrc(Src$()) As Variant()
PushNonZSz MthDryzSrc, MthLinDr(Src)
Dim Ix, L$
For Each Ix In MthIxItr(Src)
    PushI MthDryzSrc, MthLinDr(ContLin(Src, Ix))
Next
End Function

Function MthInfSrcFm(Src$(), MthFmIx&) As Variant()
Dim L$, Lines$, TopRmk$, Lno&, Cnt%
'    L = ContLin(A, MthFmIx)
    Lno = MthFmIx + 1
'    Lines = MthLineszSrcNm(A, MthFmIx)
    Cnt = LinCnt(Lines)
    TopRmk = MthTopRmkIx(Src, MthFmIx)
Dim Dr(): ' Dr = MthLinDr_Lin(L): If Si(Dr) = 0 Then Stop
MthInfSrcFm = AyAdd(Dr, Array(Lno, Cnt, Lines, TopRmk))
End Function

Property Get MthFny() As String()
MthFny = SySsl("Mdy Ty Nm Prm Ret LinRmk Lno Cnt Lines TopRmk")
End Property

Function VbeAyMthDrs(A() As Vbe) As Drs
Dim I, R%, M As Drs
For Each I In Itr(A)
    Set M = DrsInsCV(MthDrszVbe(CvVbe(I)), "Vbe", R)
    If R = 0 Then
        Set VbeAyMthDrs = M
    Else
        Stop
        PushObj VbeAyMthDrs, M
        Stop
    End If
    R = R + 1
    Debug.Print R; "<=== VbeAyMthDrs"
Next
End Function

Function VbeAyMthWs(A() As Vbe) As Worksheet
'Set VbeAyMthWs = WszDrs(VbeAyMthDrs(A))
End Function

Private Property Get ZZVbeAy() As Vbe()
PushObj ZZVbeAy, CurVbe
Const Fb$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
'PushObj ZZVbeAy, AcsOpnFb(Fb).Vbe
End Property

Private Sub Z_MthLinDryzPj()
Dim A(): A = MthLinDryzPj(CurPj)
Stop
End Sub

Private Sub Z_VbeAyMthWs()
WsVis VbeAyMthWs(ZZVbeAy)
End Sub

Private Sub Z_MthLinDryzVbe()
BrwDry MthLinDryzVbe(CurVbe)
End Sub

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

Private Sub Z()
End Sub

