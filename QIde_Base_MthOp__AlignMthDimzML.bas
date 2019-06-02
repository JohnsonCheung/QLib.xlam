VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QIde_Base_MthOp__AlignMthDimzML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Friend Function Cre(Xl As Boolean, Mc As Drs) As Drs

End Function

Friend Function Mc(Xp As Boolean, Md As CodeModule, MthLno&) As Drs
If Xp Then Exit Function
Mc = DMthCxt(Md, MthLno)
End Function

Friend Function Mcb(Xl As Boolean, Mctr As Drs) As Drs
'Fm  Xl
'Fm  Mctr L Gpno MthLin IsRmk #Mth-Cxt-TopRmk ! For each gp, the front rmk lines are TopRmk, rmv them
'Ret Mcb  L Gpno MthLin IsRmk #Mth-Cxt-Brk.   ! Brk the MthLin into V Sfx Expr Rmk
'         Dcl V Expr R1 R2 R3
If Xl Then Exit Function
Dim MthLin$, CMthLin$(), CIsRmk() As Boolean, CV$(), WV%, J%, Dr(), Dry()
AsgCol Mctr, "MthLin IsRmk", CMthLin, CIsRmk
CV = DimNy(CMthLin)
WV = WdtzAy(CV)
For J = 0 To UB(CMthLin)
    Dr = Mctr.Dry(J)
    MthLin = CMthLin(J)
    If CIsRmk(J) Then
        PushIAy Dr, McbRmk(RmvFstChr(MthLin))
    Else
        PushIAy Dr, McbLin(MthLin, CV(J), WV)
    End If
    PushI Dry, Dr
Next
Mcb = Drs(AddFF(Mctr.Fny, "Dcl V Expr R1 R2 R3"), Dry)
End Function

Friend Function McbLin(MthLin$, V$, WV%) As Variant()
Dim L$, Dcl$, Sfx$, A$, Expr$, R1$, R2$, R3$
L = MthLin
If ShfT1(L) <> "Dim" Then Stop
If V <> ShfNm(L) Then Stop
AsgBrkBet L, ":", vbSngQuote, Sfx, A, L
If HasPfx(Sfx, "As ") Then Sfx = Space(WV - Len(V) + 1) & Sfx
Dcl = V & Sfx
Expr = Aft(A, "=")
If Expr = "" Then Expr = V
AsgBrkBet L, vbPround, vbExcM, R1, R2, R3
McbLin = Array(Dcl, V, Expr, R1, R2, R3)
End Function

Friend Function McbRmk(RmkLin$)
Dim R1$, R2$, R3$
AsgBrkBet RmkLin, vbPround, vbExcM, R1, R2, R3
McbRmk = Array("", "", "", R1, R2, R3)
End Function

Friend Function Mcf(Xl As Boolean, Mctr As Drs) As Drs
'A     : L Gpno MthLin V Sfx Expr R1 R2 R3
'Ret   : Add 5 Columns: F0 FV FSfx FExpr FR1 FR2 for each gp
If Xl Then Exit Function
Dim MaxGpno%: MaxGpno = MaxzAy(IntAyzDrsC(Mctr, "Gpno"))
Dim O As Drs, IGpno%, A As Drs, B As Drs, C As Drs
For IGpno = 1 To MaxGpno
    A = DrswColEq(Mctr, "Gpno", IGpno)
    B = Mcf0(A)
    C = AddColzFiller(B, "Dcl V Expr R1 R2")
    O = AddDrs(O, C)
Next
'BrwDrs O, ShwZer:=True: Stop
Mcf = O
End Function

Friend Function Mcg(Xl As Boolean, Mcc As Drs) As Drs
'Fm  Xl
'Fm  Mcc L MthLin  #Mth-Cxt-Spc  ! no spc line & las lin
'Ret L Gpno MthLin #Mth-Cxt-Gpno ! with L in seq will be one gp
Dim Dr(), LasL&, Gpno%, L&, Dry(), J%
For J = 0 To UB(Mcc.Dry) - 1
    Dr = Mcc.Dry(J)
    L = Dr(0)
    If LasL + 1 <> L Then
        Gpno = Gpno + 1
    End If
    LasL = L
    PushI Dry, Array(L, Gpno, Dr(1))
Next
Mcg = DrszFF("L Gpno MthLin", Dry)
End Function

Friend Function McbWV%(CMthLin$())

End Function

Friend Function Xc(Xp As Boolean, Mc As Drs) As Boolean
Xc = True
If Xp Then Exit Function
If NoReczDrs(Mc) Then Exit Function
Xc = False
End Function

Friend Function Mctr(Xl As Boolean, Mcr As Drs) As Drs
' L Gpno MthLin IsRmk    #Mth-Cxt-TopRmk ! For each gp, the front rmk lines are TopRmk, rmv them
Dim IGpno%, MaxGpno, A As Drs, B As Drs, O As Drs
MaxGpno = MaxzAy(IntAyzDrsC(Mcr, "Gpno"))
For IGpno = 1 To MaxGpno
    A = DrswColEq(Mcr, "Gpno", IGpno)
    B = MctrI(A)
    O = AddDrs(O, B)
Next
Mctr = O
End Function
Friend Function MctrI(A As Drs) As Drs
' Fm  A     L Gpno MthLin IsRmk    #Mth-Cxt-TopRmk ! All Gpno are eq
' Ret MctrI L Gpno MthLin IsRmk ! Rmk TopRmk
MctrI.Fny = A.Fny
Dim J%
    Dim Dr
    For Each Dr In Itr(A.Dry)
        If Not Dr(3) Then GoTo Fnd
        J = J + 1
    Next
    Exit Function
Fnd:
    For J = J To UB(A.Dry)
        PushI MctrI.Dry, A.Dry(J)
    Next
End Function


Friend Function Mcr(Xl As Boolean, Mcg As Drs) As Drs
'Ret Mcr L Gpno MthLin IsRmk    #Mth-Cxt-isRmk
Dim Dr
For Each Dr In Itr(Mcg.Dry)
    PushI Dr, FstChr(LTrim(Dr(2))) = "'"
    Push Mcr.Dry, Dr
Next
Mcr.Fny = AddFF(Mcg.Fny, "IsRmk")
End Function
Private Function WNcd3(MthLin$) As Dictionary
Dim MthPm$: MthPm = BetBkt(MthLin)
Set WNcd3 = New Dictionary
Dim P
For Each P In Itr(TrimAy(Split(MthPm, ",")))
    WNcd3.Add TakNm(P), P
Next
'BrwDic WNcd3: Stop
End Function
Function AddColzFFDry(A As Drs, FF$, NewDry()) As Drs
AddColzFFDry = Drs(AddFF(A.Fny, FF), NewDry)
End Function
Private Function WNcd1(A As Drs) As Drs
Dim Dr, ODry()
For Each Dr In Itr(A.Dry)
    PushI Dr, BetBkt(Dr(1))
    PushI ODry, Dr
Next
WNcd1 = AddColzFFDry(A, "Pm", ODry)
End Function
Private Function WNcd4$(Pm$, Dic_V_Dcl1 As Dictionary)
Dim P, O$()
For Each P In Itr(TrimAy(Split(Pm, ",")))
    PushI O, Dic_V_Dcl1(P)
Next
WNcd4 = JnCommaSpc(O)
End Function
Private Function WNcd21$(Dcl)
Dim O$: O = Dcl
While HasSubStr(O, "  As ")
    O = Replace(O, "  As ", " As ")
Wend
WNcd21 = O
End Function
Private Function WNcd2(A As Drs) As Drs
'Fm A      V Expr Dcl Pm
'Ret WNcd2 V Expr Dcl Pm Dcl1
Dim ODry(), Dr
For Each Dr In Itr(A.Dry)
    PushI Dr, WNcd21(Dr(2))
    PushI ODry, Dr
Next
WNcd2 = Drs(AddFF(A.Fny, "Dcl1"), ODry)
End Function
Private Sub WNcd5(Dcl1$, OTyChr$, ORetAs$)
Dim A$: A = RmvNm(Dcl1)
If HasPfx(A, " As ") Then
    OTyChr = ""
    ORetAs = A
    Exit Function
End If
If HasSfx(A, "()") Then
    OTyChr = ""
    ORetAs = TyNmzTyChr(TakTyChr(A)) & "()"
End If
OTyChr = A
ORetAs = ""
End Sub
Friend Function Ncd(Xnc As Boolean, Nco$(), Mcb As Drs, Ml$) As Drs
'Ret Ncd Chdn Pm
If Xnc Then Exit Function
Dim A As Drs: A = DrswColEqSel(Mcb, "IsRmk", False, "V Expr Dcl") 'V Expr Dcl
Dim B As Drs: B = WNcd1(A) 'V Expr Dcl Pm
Dim C As Drs: C = WNcd2(B) 'V Expr Dcl Pm Dcl1
Dim D As Drs: D = SelDrs(C, "V Dcl1")
Dim E As Dictionary: Set E = DiczDrsCC(D) 'Dic_V_Dcl1
Dim F As Dictionary: Set F = WNcd3(Ml)    'Dic_MthPm_Dcl
Dim G As Dictionary: Set G = AddDic(E, F)
Dim CV$(), CPm$(), CDcl1$()
    AsgCol C, "V Pm Dcl1", CV, CPm, CDcl1
    Brw CDcl1
    Stop
Dim ODry()
    Dim J%
    For J = 0 To UB(CV)
        Dim Dcl1$: Dcl1 = CDcl1(J)
        Dim TyChr$, RetAs$
            WNcd5 Dcl1, TyChr, RetAs
        PushI ODry, Array(CV(J), TyChr, WNcd4(CPm(J), G), RetAs)
    Next
Ncd = DrszFF("Chdn Pm", ODry)
End Function

Friend Function Xnn(X As Boolean, A$()) As Boolean
Xnn = True
If X Then Exit Function
If Si(A) = 0 Then Exit Function
Xnn = False
End Function
Friend Function Nca(Xnc As Boolean, Md As CodeModule) As String()
If Xnc Then Exit Function
Nca = MthnyzM(Md)
End Function

Friend Function Nco(Xnc As Boolean, Nc As Drs) As String()
'Fm Nc Chdn TyChr MthPm RetAs
If Xnc Then Exit Function
Dim Dr
Const C$ = "Friend Function ??(?)?"
For Each Dr In Itr(Nc.Dry)
    PushI Nco, FmtQQ(C, Dr(0), Dr(1), Dr(2), Dr(3))
Next
Brw Nco: Stop
End Function


Friend Function Omc(Xl As Boolean, Md As CodeModule, Mco As Drs) As Unt
'Fm Xl
'Fm Mco L DimLin
Exit Function
If Xl Then Exit Function
Dim Dr
For Each Dr In Itr(Mco.Dry)
    Md.ReplaceLine Dr(0), Dr(1)
Next
End Function

Friend Function Onc(Xnc As Boolean, Nco$()) As Unt

End Function

Friend Function Mg%(X As Boolean, A As Drs)
'Fm  L Gpno V Sfx Expr Rmk1 Rmk2 F0 FV..FRmk1 ! Adding columns EmpCol FV..FRmk1
'Ret MaxGpno%
If X Then Exit Function
Mg = MaxzAy(IntAyzDrsC(A, "Gpno"))
End Function


Friend Function Cg(A As Drs) As Drs
'Fm  L DimLin      ! all the DimLin and its rmk chr ['] position
'Ret L Gpno DimLin ! all lines with L in seq, it will be same Gpno
Dim Dr, LasL&, Gpno%, L&, Dry()
For Each Dr In Itr(A.Dry)
    L = Dr(0)
    If LasL + 1 <> L Then
        Gpno = Gpno + 1
    End If
    LasL = L
    PushI Dry, Array(L, Gpno, Dr(1))
Next
Cg = DrszFF("L Gpno DimLin", Dry)
End Function
Friend Function N(X As Drs) As Drs
'Fm  L Gpno RmkLin                       ! Keep only IsLas=True
'Ret L Gpno RmkLin RmkChrPos R1 R2 FRmk1 ! find FRmk1 for each Gp
Dim GpnoAy%(): GpnoAy = IntAyzDrsC(X, "Gpno")
Dim MaxGpno%: MaxGpno = MaxzAy(GpnoAy)
Dim IGpno%
For IGpno = 1 To MaxGpno
    Dim A As Drs: A = DrswColEq(A, "Gpno", IGpno)
    Dim B As Drs: B = Ni(A)
    Dim O As Drs: O = AddDrs(O, B)
Next
End Function
Friend Function Ni(A As Drs) As Drs

End Function

Friend Function Mco(Xl As Boolean, Mcd As Drs) As Drs
'Fm : L MthLin DimLin
'Ret: L Dim ! Keep MthLin<>DimLin
Dim Dr, Dry()
For Each Dr In Itr(Mcd.Dry)
    If Trim(Dr(1)) <> Dr(1) Then Stop
    If Trim(Dr(2)) <> Dr(2) Then Stop
    If Dr(1) <> Dr(2) Then
        PushI Dry, Array(Dr(0), Dr(2))
    End If
Next
Mco = DrszFF("L DimLin", Dry)
End Function

Friend Function Mcd(Xl As Boolean, Mcf As Drs) As Drs
'Fm  Xl
'Fm  Mcf L Gpno MthLin IsRmk Dcl V Expr R1 R2 R3 F0 FV FSfx FExpr FR1 FR2 ! Upd F0..FR1
'Ret Mcd L MthLin DimLin                              ! Bld the new DimLin
If Xl Then Exit Function
'BrwDrs Mcf: Stop
Dim Dr, IDcl%, L&, MthLin$, IsRmk As Boolean, DimLin$, ODry()
IDcl = IxzAy(Mcf.Fny, "Dcl")
For Each Dr In Itr(Mcf.Dry)
    L = Dr(0)
    MthLin = Dr(1)
    IsRmk = Dr(3)
    If IsRmk Then
        DimLin = McdRmk(Dr, IDcl%)
    Else
        DimLin = McdLin(Dr, IDcl%)
    End If
    PushI ODry, Array(L, MthLin, DimLin)
Next
Mcd = DrszFF("L MthLin DimLin", ODry)
End Function
Friend Function McdLin$(McfDr, IDcl%)
'Fm  Mcf L Gpno MthLin IsRmk Dcl V Expr R1 R2 R3 F0 FDcl FV FExpr FR1 FR2 ! Upd F0..FR1
'Ret DimLin$
'Brw LyzNNAv("L Gpno MthLin IsRmk Dcl V Expr R1 R2 R3 F0 FDcl FV FExpr FR1 FR2", CvAv(McfDr)): Stop
Dim Dr(): Dr = McfDr
Dim L&:           L = Dr(0)
Dim MthLin$: MthLin = Dr(2)
Dim IsRmk As Boolean: IsRmk = Dr(3)
Dim Dcl$:      Dcl = Dr(IDcl)
Dim V$:         V = Dr(IDcl + 1)
Dim Expr$:   Expr = Dr(IDcl + 2)
Dim R1$:       R1 = Dr(IDcl + 3)
Dim R2$:       R2 = Dr(IDcl + 4)
Dim R3$:       R3 = Dr(IDcl + 5)
Dim F0$:       F0 = Space(Dr(IDcl + 6))
Dim FDcl$:       FDcl = Space(Dr(IDcl + 7))
Dim FV$:   FV = Space(Dr(IDcl + 8))
Dim FExpr$: FExpr = Space(Dr(IDcl + 9))
Dim FR1$:     FR1 = Space(Dr(IDcl + 10))
Dim FR2$:     FR2 = Space(Dr(IDcl + 11))
'
Dim OD$:  OD = Dcl
Dim OL$:  OL = FDcl & FV & V
Dim OE$:  OE = Expr & FExpr
Dim OM$:  OM = XRmk(FR1, FR2, R1, R2, R3)
McdLin = RTrim(FmtQQ("?Dim ?: ? = ? ' ?", F0, OD, OL, OE, OM))
End Function
Friend Function McdRmk$(McfDr, IDcl%)
'Fm  Mcf L Gpno MthLin IsRmk Dcl V Expr R1 R2 R3 F0 FDcl FV FExpr FR1 FR2
'Ret DimLin$
Dim Dr(): Dr = McfDr
Dim L&:           L = Dr(0)
Dim MthLin$: MthLin = Dr(2)
Dim IsRmk As Boolean: IsRmk = Dr(3)
Dim Dcl$:     Dcl = Dr(IDcl)
Dim V$:         V = Dr(IDcl + 1)
Dim Expr$:   Expr = Dr(IDcl + 2)
Dim R1$:       R1 = Dr(IDcl + 3)
Dim R2$:       R2 = Dr(IDcl + 4)
Dim R3$:       R3 = Dr(IDcl + 5)
Dim F0%:       F0 = Dr(IDcl + 6)
Dim FDcl%:   FDcl = Dr(IDcl + 7)
Dim FV%:       FV = Dr(IDcl + 8)
Dim FExpr%: FExpr = Dr(IDcl + 9)
Dim FR1$:     FR1 = Space(Dr(IDcl + 10))
Dim FR2$:     FR2 = Space(Dr(IDcl + 11))
'
Dim OS$:  OS = Space(11 + Len(V) + Len(Dcl) + Len(Expr) + FV + FDcl + FExpr)
Dim OM$:  OM = XRmk(FR1, FR2, R1, R2, R3)
McdRmk = Trim(FmtQQ("'??", OS, OM))
End Function
Friend Function Xnf(Xp, Md As CodeModule, MthLno&)
Xnf = True
If Xp Then Exit Function
If Not HasPfx(SrcLinzNxt(Md, MthLno), "Static F As New") Then Exit Function
Xnf = False
End Function

Friend Function XRmk$(FR1$, FR2$, R1$, R2$, R3$)
If R1 = "" And R2 = "" And R3 = "" Then Exit Function
Dim A$, B$, C$
A = R1 & FR1
If R2 = "" Then
    B = "  " & FR2
Else
    B = " #" & R2 & FR2
End If
If R3 <> "" Then
    C = " ! " & R3
End If
XRmk = RTrim(A & B & C)
End Function

Friend Function Ncc(Md As CodeModule) As String()

End Function
Friend Function Occ(Nc As Drs) As String()
'Fm  Nc  Chdn MthPfx Pm TyChr RetTy #New-Chd-mth.
'Ret Occ MthLin                     #Oup-Create-Chd-mthfun.  MthLin is always a function
Dim Dr, Chdn$, MthPfx$, Pm$, TyChr$, RetTy$
For Each Dr In Itr(Nc.Dry)
    Chdn = Dr(0)
    MthPfx = Dr(1)
    Pm = Dr(2)
    TyChr = Dr(3)
    RetTy = Dr(4)
    Dim MthLin$: MthLin = FmtQQ("Friend Function ???(?)?", MthPfx, Chdn, TyChr, Pm, RetTy)
    PushI Occ, MthLin
Next
End Function

Friend Function Xp(Md As CodeModule, MthLno&) As Boolean
Xp = True
If IsNothing(Md) Then Debug.Print "Md is nothing": Exit Function
If MthLno <= 0 Then Debug.Print "MthLno <= 0": Exit Function
Xp = False
End Function

Friend Function Mlc$(Xp As Boolean, Mlf As Boolean, Ml$)
'Fm  IsFun
'Fm  MthLin
'Ret FunTyChr
If Xp Then Exit Function
If Not Mlf Then Exit Function
Mlc = MthTyChr(Ml)
End Function

Friend Function Mlr$(Xp As Boolean, Mlf As Boolean, Ml$)
'Fm  IsFun
'Fm  MthLin
'Ret FunTyChr
If Xp Then Exit Function
If Not Mlf Then Exit Function
Mlr = MthRetTy(Ml)
End Function


Friend Function Mln$(Xl As Boolean, Ml$)
If Xl Then Exit Function
Dim O$: O = Mthn(Ml): If O = "" Then Stop
Mln = Mthn(Ml)
End Function
Friend Function Mcc(X As Boolean, A As Drs) As Drs
'Fm  Mth MthLin ! mth Mct
'Ret Mth MthLin ! Rmv spc/Brw/Insp/Stop lin and las lin
If X Then Exit Function
Dim Dr, Dry(), L$
For Each Dr In Itr(A.Dry)
    L = LTrim(Dr(1))
    Select Case True
    Case FstChr(L) = "'", T1(L) = "Dim":  PushI Dry, Dr
    End Select
Next
Mcc = Drs(A.Fny, Dry)
End Function
Friend Function Xa(Xc As Boolean, Mcc As Drs) As Boolean
' Fm  Xp          #eXit-Parameter-er.
' Fm  Mc L MthLin #Mth-Cxt.
' Ret Xa          #eXit-Alignable-Er. ! Not alignable or no lines
Xa = True ' Assume error
If Xc Then Exit Function
If NoReczDrs(Mcc) Then Exit Function
Xa = False ' No Error
End Function


Friend Sub Cty(Md As CodeModule, B As Drs)
'L Gpno MthLin        ! no las lin
'Ret L Gpno MthLin Ty !Ty: 1(GpRmk) 2(Dim) 3(DimRmk)
Dim NewGp As Boolean, IsDim As Boolean, IsRmk As Boolean, LasGpno%, Dr, CurGpno%
NewGp = True
LasGpno = B.Dry(0)(1)
For Each Dr In B.Dry
    CurGpno = Dr(1)
    Select Case True
    Case LasGpno = CurGpno
    Case IsRmk
    End Select
Next
End Sub

Friend Function Mcf0(A As Drs) As Drs
'Fm   : L Gpno MthLin V Sfx Expr R1 R2 R3
'Ret  : Add  column-filler-F0  ! Sfx-0-of-Mcf0 is add-column-filler-0-F0
Dim Dr: Dr = A.Dry(0)
Dim F0%: F0 = XF0zDimLin(Dr(2))
Dim ODry()
For Each Dr In Itr(A.Dry)
    PushI Dr, F0
    PushI ODry, Dr
Next
Mcf0 = Drs(AddFF(A.Fny, "F0"), ODry)
End Function

Friend Function Cra() As Drs

End Function
Friend Function Ocr() As Drs

End Function

Friend Function Nce(Nnc As Boolean, Mcb As Drs) As String()
'Ret #New-Chdmth-Ept. ! It is from V and Expr=F.{V}
If Nnc Then Exit Function
Dim A As Drs: A = DrswColEq(Mcb, "IsRmk", False)
Dim CV$(), CExpr$()
    AsgCol A, "V Expr", CV, CExpr
    If Si(CV) = 0 Then Exit Function

Dim V, J%
For Each V In CV
    If HasPfx(CExpr(J), "F.") Then
        PushI Nce, CV(J)  '<=====
    End If
    J = J + 1
Next
'Brw Nce: Stop
End Function

Friend Function Mlf(X As Boolean, A$) As Boolean
'Fm  MthLin
'Ret IsFun
If X Then Exit Function
Mlf = MthKd(A) = "Function"
End Function


Friend Function XF0zDimLin%(DimLin)
If T1(DimLin) <> "Dim" Then Stop
Dim L$: L = LTrim(DimLin)
XF0zDimLin = Len(DimLin) - Len(L)
End Function
Friend Function Ml$(X As Boolean, Md As CodeModule, MthLno&)
If X Then Exit Function
Ml = ContLinzML(Md, MthLno)
End Function

