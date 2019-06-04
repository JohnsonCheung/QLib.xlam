Attribute VB_Name = "QIde_Mth_Drs"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Drs."
Private Const Asm$ = "QIde"
Type MthInf
    Mdn As String
    FmLno As Long
    ToLno As Long
    LinCnt As Integer
    Lines As String
    MthLin As String
    Mthn As String
    ShtMdy As String
    ShtKd As String
    TyChr As String
    RetTy As String
    LinRmk As String
    TopRmk As String
End Type
Type MthInfs: N As Long: Ay() As MthInf: End Type

Function MthDrszFb(Fb) As Drs
MthDrszFb = MthDrszV(VbezPjf(Fb))
ClsPjf Fb
End Function

Function MthDrsInVbe() As Drs
MthDrsInVbe = MthDrszV(CVbe)
End Function

Function MthDrszFxa(Fxa$, Optional Xls As Excel.Application) As Drs
Dim A As Excel.Application: Set A = DftXls(Xls)
MthDrszFxa = MthDrszP(PjzFxa(Fxa))
If IsNothing(Xls) Then QuitXls Xls
End Function

Function MthDrszMd(M As CodeModule) As Drs
MthDrszMd = Drs(MthFny, MthDryzMd(M))
End Function

Function MthDrszP(P As VBProject) As Drs
Dim O As Drs
O = Drs(MthFny, MthDryzP(P))
O = AddColzValIdzCntzDrs(O, "Lines", "Pj")
O = AddColzValIdzCntzDrs(O, "Nm", "PjMth")
MthDrszP = O
End Function

Function MthDrszPjf(Pjf$) As Drs
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

Function MthDrszPjfy1() As Drs
MthDrszPjfy1 = MthDrszPjfy(Pjfy)
End Function

Function MthDrszPjfy(Pjfy$()) As Drs
Dim I
For Each I In Pjfy
    ApdDrs MthDrszPjfy, MthDrszPjf(CStr(I))
Next
End Function

Function MthDrszV(A As Vbe) As Drs
MthDrszV = Drs(MthFny, MthDryzV(A))
End Function

Function MthDryzMd(M As CodeModule) As Variant()
Dim P$, T$, N$
P = PjnzM(M)
T = ShtCmpTyzMd(M)
N = Mdn(M)
MthDryzMd = DryInsColzV3(MthDryzS(Src(M)), P, T, N)
End Function

Function MthDryzP(P As VBProject) As Variant()
Dim M
For Each M In MdItr(P)
    PushIAy MthDryzP, MthDryzMd(CvMd(M))
Next
End Function

Function MthDryzS(Src$()) As Variant()
Dim MthLin
For Each MthLin In Itr(MthLinAyzS(Src))
    PushI MthDryzS, Dr_MthLin(CStr(MthLin))
Next
End Function

Function MthDryzV(A As Vbe) As Variant()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MthDryzV, MthDryzP(P)
Next
End Function

Function MthInfszM(M As CodeModule) As MthInfs
'MthInfAy_Md = MthInfAy_Src(PjnMth(A), Mdn(A), CSrc(A))
End Function
Sub PushMthInf(O As MthInfs, M As MthInf)

End Sub
Sub PushMthInfs(O As MthInfs, M As MthInfs)

End Sub
Function MthInfszPMS(Pjn$, Mdn, Src$()) As MthInfs
Dim I
For Each I In Itr(MthIxy(Src))
    PushMthInf MthInfszPMS, MthInfzPMSI(Pjn, Mdn, Src, I)
Next
End Function

Function MthInfszP(P As VBProject) As MthInfs
Dim C As VBComponent
For Each C In P.VBComponents
    PushMthInfs MthInfszP, MthInfszM(C.CodeModule)
Next
End Function

Function MthInfszV(A As Vbe) As MthInfs
Dim P As VBProject
For Each P In A.VBProjects
    PushMthInfs MthInfszV, MthInfszP(P)
Next
End Function

Function MthInfzPMSI(Pjn$, Mdn, Src$(), FmIx) As MthInf
Dim L$, O As MthInf
O.MthLin = ContLin(Src, FmIx)
'L = Src(FmIx)
'O.ShtMdy = ShtMdy(ShfMdy(L))
'O.ShtKd = ShtMthKd(ShfMthTy(L))
'MthInfMdzPMSI = O
Stop
End Function

Function Dr_MthzSI(Src$(), MthIx) As Variant()
Dim L$, Lines$, Rmk$(), Lno, Cnt%
    L = ContLin(Src, MthIx)
    Lno = MthIx + 1
    Lines = MthLineszSI(Src, MthIx)
    Cnt = LinCnt(Lines)
    Rmk = TopRmkLy(Src, MthIx)
Dim Dr():  Dr = Dr_MthLin(L): If Si(Dr) = 0 Then Stop
Dr_MthzSI = AddAy(Dr, Array(Lno, Cnt, Lines, Rmk))
End Function

Function Dr_MthLin(MthLin) As Variant()
'If Not HitMthLin(MthLin, B) Then Exit Function
Dim X As MthLinRec
X = MthLinRec(MthLin)
With X
Dr_MthLin = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function
Function MthDr(Src$(), MthLin, MthIx) As Variant()
'If Not HitMthLin(MthLin, B) Then Exit Function
Dim X As MthLinRec
X = MthLinRec(MthLin)
With X
MthDr = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function

Function Dr_MthLinsP() As Drs
Dr_MthLinsP = Dr_MthLinszP(CPj)
End Function

Function Dr_MthLinszP(P As VBProject) As Drs
Dr_MthLinszP = Drs(MthLinFny, Dry_MthLinzP(CPj))
End Function

Function Dry_MthLinzM(M As CodeModule) As Variant()
Dim P$, T$, N$
P = PjnzM(M)
T = ShtCmpTyzMd(M)
N = Mdn(M)
Dry_MthLinzM = DryInsColzV3(Dry_MthLinzS(Src(M)), P, T, N)
End Function

Function Dry_MthLinzP(P As VBProject) As Variant()
Dim M
For Each M In MdItr(P)
    PushAy Dry_MthLinzP, Dry_MthLinzM(CvMd(M))
Next
End Function

Function Dry_MthLinzS(Src$()) As Variant()
Dim MthLin, W As WhMth
'Set W = WhMthzStr
For Each MthLin In Itr(MthLinAyzS(Src))
    'PushISomSi Dry_MthLinzS, Dr_MthLin(CStr(MthLin), W)
Next
End Function

Function MthWb() As Workbook
Set MthWb = ShwWb(MthWbPjfy(Pjfy))
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

Function MthWbPjfy(Pjfy$()) As Workbook
Set MthWbPjfy = MthWbFmt(WbzWs(MthWszPjfy(Pjfy)))
End Function

Function MthWsInVbe() As Worksheet
Set MthWsInVbe = MthWszV(CVbe)
End Function

Function MthWszP(P As VBProject) As Worksheet
Set MthWszP = WszDrs(MthDrszP(P))
End Function

Function MthWszPjfy(Pjfy$()) As Worksheet
Dim O As Drs
O = MthDrszPjfy(Pjfy)
O = AddColzValIdzCntzDrs(O, "Nm", "Vbe_Mth")
O = AddColzValIdzCntzDrs(O, "Lines", "Vbe")
Dim Ws As Worksheet
Set Ws = WszDrs(O)
SetWsCdNmAndLoNm Ws, "MthFull"
Set MthWszPjfy = Ws
End Function

Function MthWszV(A As Vbe) As Worksheet
Set MthWszV = WszDrs(MthDrszV(A))
End Function

Function PjnMth$(M As CodeModule)
PjnMth = M.Parent.Collection.Parent.Name
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
Set MthWs = MthWszPjfy(Pjfy)
End Property

Property Get Pjfy() As String()
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
Pjf = Pjfy()(0)
ShwWs WszDrs(MthDrszPjf(Pjf))
End Sub

Private Sub Z_Dry_MthLinzP()
Dim A(): A = Dry_MthLinzP(CPj)
Stop
End Sub

Private Sub Z_Dr_MthLinAyzV()
'BrwDry Dr_MthLinAyzV(CVbe)
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
