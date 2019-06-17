Attribute VB_Name = "QIde_B_MthOp"
Option Compare Text
Option Explicit
Private Const Asm$ = ""
Private Const CMod$ = "MIde_Mth_Op."
Enum EmRpt
    EiRptOnly
    EiUpdAndRpt
    EiUpdOnly
    EiPushOnly
    EiUpdAndPush
End Enum
Public Type R123
    R1 As String
    R2 As String
    R3 As String
End Type
Public Type Vr
    V As String
    R() As R123
End Type
Public Type Vpr
    V As String
    Pm() As Vr
    Ret() As R123
End Type

Sub Z3()
Dim M As CodeModule: Set M = Md("QDao_Lnk_ErzLnk")
Dim L&: L = MthLnozMM(M, "ErzLnk")
AlignMthDimzML M, L
End Sub

Function SiVpr%(A() As Vpr):   On Error Resume Next: SiVpr = UBound(A) + 1: End Function
Function SiVr%(A() As Vr):     On Error Resume Next:  SiVr = UBound(A) + 1: End Function
Function SiR123%(A() As R123): On Error Resume Next:  SiR123 = UBound(A) + 1: End Function

Sub PushVpr(O() As Vpr, M As Vpr)
Dim N%: N = SiVpr(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function R123(R1$, R2$, R3$) As R123
With R123
    .R1 = R1
    .R2 = R2
    .R3 = R3
End With
End Function

Function FmtVrAy(A() As Vr) As String()
Dim J%: For J = 0 To SiVr(A) - 1
    PushIAy FmtVrAy, FmtVr(A(J))
Next
End Function

Function FmtVr(A As Vr) As String()
PushI FmtVr, "V(" & A.V & ")"
PushIAy FmtVr, FmtDry(DryzR123Ay(A.R))
End Function

Function DryzR123Ay(R() As R123) As Variant()
Dim J%: For J = 0 To SiR123(R) - 1
    With R(J)
    PushI DryzR123Ay, Array(.R1, .R2, .R3)
    End With
Next
End Function

Sub PushR123(O() As R123, R1$, R2$, R3$)
Dim N%: N = SiR123(O)
ReDim Preserve O(N)
O(N) = R123(R1, R2, R3)
End Sub

Sub PushVr(O() As Vr, V$, R() As R123)
Dim N%: N = SiVr(O)
ReDim Preserve O(N)
O(N) = Vr(V, R)
End Sub

Function Vr(V$, R() As R123) As Vr
With Vr
    .V = V
    .R = R
End With
End Function

Function IsLinzAsg(L) As Boolean
Dim A$: A = L
If ShfDotNm(A) = "" Then Exit Function
IsLinzAsg = T1(A) = "="
End Function

Private Sub Z_IsSngDimCol()
Dim L$
GoSub T0
Exit Sub
T0:
    L = "Dim IsDesAy() As Boolean: IsDesAy = XIsDesAy(Ay)"
    GoTo Tst
Tst:
    Act = IsLinzSngDimColon(L)
    C
    Stop
    Return
End Sub

Function IsLinzSngDimColon(L) As Boolean
'Ret true if L is Single-Dim-Colon: one V aft Dim and Colon aft DclSfx
Dim Lin$: Lin = L
If Not ShfDim(Lin) Then Exit Function
If ShfNm(Lin) = "" Then Exit Function
ShfBkt Lin
ShfDclSfx Lin
IsLinzSngDimColon = FstChr(Lin) = ":"
End Function

Sub Z_IsLinzSngDimColon()
Dim L
'GoSub T0
GoSub ZZ
Exit Sub
T0:
    L = "Dim A As Access.Application: Set A = DftAcs(Acs)"
    Ept = True
    GoTo Tst
Tst:
    Act = IsLinzSngDimColon(L)
    If Act <> Ept Then Stop
    Return
ZZ:
    Dim A As New Aset
    For Each L In SrczP(CPj)
        L = Trim(L)
        If T1(L) = "Dim" Then
            Dim S$: S = IIf(IsLinzSngDimColon(L), "1", "0")
            A.PushItm S & " " & L
        End If
    Next
    A.Srt.Vc
    Return
End Sub

Sub AlignMthDimzML(M As CodeModule, MthLno&, Optional SkpChkSelf As Boolean, Optional Rpt As EmRpt)
Static F As New QIde_B_MthOp__AlignMthDimzML
Dim D1 As Drs, D2 As Drs
'== Exit if parameter error ============================================================================================
                  If F.IsPmEr(M, MthLno) Then Exit Sub       ' X-Parameter-er. M-isnothg | MthLno<=0
Dim Ml$:     Ml = ContLinzML(M, MthLno)
Dim MlNm$: MlNm = Mthn(Ml)              '  # Ml-Name.
                  If F.XSelf(SkpChkSelf, M, MlNm) Then Exit Sub ' #X-Self-aligning-er. ! Mdn<>'QIde...' & MlNm<>'AlignMthDimzML
Dim Mc As Drs: Mc = DMthCxtzML(M, MthLno) ' L MthLin # Mc.
                    If NoReczDrs(Mc) Then Exit Sub

Dim IsUpd As Boolean: IsUpd = IsUpdzRpt(Rpt)

'== Align DblEqRmk (De) ================================================================================================
'   When a rmk lin begins with '== or '--, expand it to 120 = or -
Dim De      As Drs:      De = F.De(Mc)                        ' L MthLin    # Dbl-Eq | Dbl-Dash
Dim DeLNewO As Drs: DeLNewO = F.DeLNewO(De)                   ' L NewL OldL
Dim OUpdDe:                   If IsUpd Then RplLin M, DeLNewO

'== Align Mth Cxt ======================================================================================================
Dim McCln As Drs: McCln = F.McCln(Mc) ' L MthLin # Mc-Cln. ! must SngDimColon | Rmk(but not If Stop Insp == Brw). Cln to Align
If NoReczDrs(McCln) Then Exit Sub
Dim McGp    As Drs:    McGp = F.McGp(McCln)     ' L Gpno MthLin               ! with L in seq will be one gp
Dim McRmk   As Drs:   McRmk = F.McRmk(McGp)     ' L Gpno MthLin IsRmk         ! a column IsRmk is added
Dim McTRmk  As Drs:  McTRmk = F.McTRmk(McRmk)   ' L Gpno MthLin IsRmk         ! For each gp, the front rmk lines are TopRmk, rmv them
Dim McInsp  As Drs:  McInsp = F.McInsp(McTRmk)  ' L Gpno MthLin IsRmk         ! If las lin is rmk and is 'Insp, exl it.
Dim McVSfx  As Drs:  McVSfx = F.McVSfx(McInsp)  ' L Gpno MthLin IsRmk
                                                ' V Sfx Rst
Dim McDcl   As Drs:   McDcl = F.McDcl(McVSfx)   ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl Rst               ! Add Dcl from V & Sfx
Dim McLR    As Drs:    McLR = F.McLR(McDcl)     ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl LHS RHS Rst       ! Add LHS Expr from Rst
Dim McR123  As Drs:  McR123 = F.McR123(McLR)    ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl LHS Expr R1 R2 R3 ! Add R1 R2 R3 from Rst
Dim McFill  As Drs:  McFill = F.McFill(McR123)  ' L Gpno MthLin IsRmk
                                                ' V Sfx Dcl LHS Expr
                                                ' F0 FSfx FExpr FR1 FR2       ! Adding F*
Dim McAlign As Drs: McAlign = F.McAlign(McFill) ' L Align                     ! Bld the new Align

                         D1 = DrseCeqC(McAlign, "MthLin Align")
                         D2 = SelDrs(D1, "L Align MthLin")
Dim McLNewO As Drs: McLNewO = LNewO(D2.Dry)
Dim OAlignCm:                 If IsUpd Then RplLin M, McLNewO

'== Gen Bs (Brw-Stmt) ==================================================================================================
Dim Bs      As Drs:              Bs = F.Bs(McCln)                               ' L BsLin ! FstTwoChr = '@
Dim Bs1     As Drs:             Bs1 = ColEqSel(McR123, "IsRmk", False, "V Sfx")
Dim Bs2     As Drs:             Bs2 = ColNe(Bs1, "V", "")
Dim VSfx    As Dictionary: Set VSfx = DiczDrsCC(Bs2)
Dim Mdn$:                       Mdn = MdnzM(M)
Dim BsLNewO As Drs:         BsLNewO = F.BsLNewO(Bs, VSfx, Mdn, MlNm)
Dim OUpdBs:                           If IsUpd Then RplLin M, BsLNewO

                       D1 = ColPfx(Mc, "MthLin", "Static F")
Dim NoSf As Boolean: NoSf = NoReczDrs(D1)
If Not NoSf Then
    '== Ens Static-F-Dcl (Sf) ==========================================================================================
    Dim Dr1:                Dr1 = D1.Dry(0)
    Dim SfLno&:           SfLno = Dr1(0)
    Dim SfLin$:           SfLin = Dr1(1)
    Dim CclsNm$:         CclsNm = IIf(NoSf, "", Mdn & "__" & MlNm)
    Dim SfLNewO As Drs: SfLNewO = F.SfLNewO(NoSf, SfLno, SfLin, CclsNm) ' Only one or no line
    Dim OEnsSf:                   If IsUpd Then RplLin M, SfLNewO
    
    '== Ens Chd-Cls (Ccls)==============================================================================================
    Dim HasCcls As Boolean:     HasCcls = HasCmpzP(PjzM(M), CclsNm)
    Dim OEnsCcls:                         If IsUpd Then F.OEnsCcls NoSf, HasCcls, M, CclsNm
    Dim Ccls    As CodeModule: Set Ccls = PjzM(M).VBComponents(CclsNm).CodeModule
End If

'== Crt Chd-Mth (Cm)====================================================================================================
Dim CmMd  As CodeModule: Set CmMd = IIf(NoSf, M, Ccls)
Dim CmMdy$:                 CmMdy = IIf(NoSf, "Private", "Friend")
Dim CmV   As Drs:             CmV = ColEqSel(McR123, "IsRmk", False, "V Sfx LHS RHS") ' V Sfx LHR RHS
Dim WiSf  As Boolean:        WiSf = Not NoSf
Dim MlNmDD$:               MlNmDD = BefOrAll(MlNm, "__") & "__"
Dim CmNm  As Drs:            CmNm = F.CmNm(CmV, WiSf, MlNmDD)                         ' V Sfx LHS RHS CmNm ! som CmNm may be blank
Dim CmEpt As Drs:           CmEpt = ColNe(CmNm, "CmNm", "")                           ' V Sfx LHS RHS CmNm ! All CmNm has val
Dim CmEptNm$():           CmEptNm = StrCol(CmEpt, "CmNm")                             '                    ! It is ept mth ny.  They will be used create new chd mth
Dim CmActNm$():           CmActNm = MthNyzM(CmMd)                                     '                    ! It is from chd cls of given md
Dim CmNewNm$():           CmNewNm = MinusAy(CmEptNm, CmActNm)                         '                    ! The new ChdMthNy to be created.
Dim CmNew As Drs:           CmNew = ColIn(CmEpt, "CmNm", CmNewNm)
Dim CdNewCm$:             CdNewCm = F.CdNewCm(CmNew, CmMdy)                           '                    ! Cd to be append to CmMd
Dim OCrtCm:                         If IsUpd Then ApdLines CmMd, CdNewCm

'== Upd Chd-Mth-Lin (Cml) ==============================================================================================
'   If the calling pm has been changed, the chd-mth-lin will be updated.
Dim MlVSfx    As Drs:    MlVSfx = F.MlVSfx(Ml)               ' Ret V Sfx                           ! the MthLin's pm V Sfx
                             D1 = SelDrs(CmV, "V Sfx")
Dim CmlVSfx   As Drs:   CmlVSfx = AddDrs(MlVSfx, D1)
Dim CmlPm     As Drs:     CmlPm = F.CmlPm(CmEpt)             ' V Sfx RHS CmNm Pm
Dim CmlDclPm  As Drs:  CmlDclPm = F.CmlDclPm(CmlPm, CmlVSfx) ' V Sfx RHS CmNm Pm DclPm             ! use [CmlVSfx] & [Pm] to bld [DclPm]
Dim CmlMthRet As Drs: CmlMthRet = F.CmlMthRet(CmlDclPm)      ' V Sfx RHS CmNm Pm DclPm TyChr RetAs
Dim CmlEpt    As Drs:    CmlEpt = F.CmlEpt(CmlMthRet, CmMdy) ' V CmNm EptL
                             D1 = DMth(CmMd)                 ' L Mdy Ty Mthn MthLin
                                If NoSf Then D1 = ColEq(D1, "Mdy", "Prv")
                                If WiSf Then D1 = ColEq(D1, "Mdy", "Frd")
Dim CmlAct   As Drs:   CmlAct = SelDrszAs(D1, "L Mthn:CmNm MthLin:ActL") ' L CmNm ActL
Dim CmlJn    As Drs:    CmlJn = JnDrs(CmlEpt, CmlAct, "CmNm", "L ActL")  ' V CmNm EptL L ActL ! som EptL & ActL may eq
                           D2 = DrseCeqC(CmlJn, "EptL ActL")             ' V CmNm EptL L ActL ! All EptL & ActL are diff
Dim CmlLNewO As Drs: CmlLNewO = SelDrszAs(D2, "L EptL:NewL ActL:OldL")   ' L NewL OldL
Dim OUpdCml:                    If IsUpd Then RplLin CmMd, CmlLNewO

'== Rpl Mth-Brw (Mb)====================================================================================================
'   Des: Mth-Brw is a remarked Insp-stmt in each las lin of cm.  It insp all the inp oup
'   Lgc: Fnd-and-do MbLNewO
'        Fnd-and-do NewMb
'BrwDrs CmlEpt: Stop
Dim CmLis   As Drs:   CmLis = SelDrszAs(CmlEpt, "CmNm:Mthn EptL:MthLin") ' Mthn MthLin
Dim MbEpt   As Drs:   MbEpt = F.MbEpt(CmLis, Mdn)                        ' Mthn MthLin MbStmt
Dim Cm$():               Cm = StrCol(CmLis, "Mthn")
Dim MbAct   As Drs:   MbAct = F.MbAct(Cm, CmMd)                          ' L Mthn OldL               ! OldL is MbStmt
Dim MbJn    As Drs:    MbJn = JnDrs(MbEpt, MbAct, "Mthn", "OldL L")      ' Mthn MthLin MbStmt OldL L
Dim MbSel   As Drs:   MbSel = SelDrszAs(MbJn, "L MbStmt:NewL OldL")      ' L NewL OldL
Dim MbLNewO As Drs: MbLNewO = DrseCeqC(MbSel, "NewL OldL")
Dim OUpdMb:                   If IsUpd Then RplLin CmMd, MbLNewO

'== Crt Mth-Brw (Mb)====================================================================================================
                     D1 = LJnDrs(MbEpt, MbAct, "Mthn", "L", "HasAct") ' Mthn MthLin MbStmt L HasAct
                     D2 = ColEq(D1, "HasAct", False)                  ' Mthn MthLin MbStmt L HasAct
Dim MbNew As Drs: MbNew = SelDrszAs(D2, "Mthn MbStmt:NewL")
Dim OCrtMb:               If IsUpd Then F.OCrtMb CmMd, MbNew

'== Upd Chd-Rmk (Cr) ===================================================================================================
'   If any of the calling pm's rmk is changed, the chd-mth-rmk will be updated
Const XRmv$ = "Delete * from [#Sel] where Len(Trim(R1))=0 and Len(Trim(R2))=0 and Len(TRim(R3))=3"

Dim Cr1Sel   As Drs:       Cr1Sel = SelDrs(McR123, "V R1 R2 R3")         ' V R1 R2 R3
Dim Db As Database: Set Db = TmpDb
Dim CrTSel:                  CrtTblzDrs Db, "#Sel", Cr1Sel
Dim CrRmv:                   Db.Execute XRmv                             ' V R1 R2 R3 ! rmv those R1 2 3 are blank
Dim CrFillV:                 FillLasVzDb Db, "#Sel", "V"                ' V R1 R2 R3   ! Fill Blank-V with LasV
Dim CrAli:                   AlignColzDb Db, "#Sel", "#Ali", "V", "R1 R2 R3"    ' V R1 R2 R3   ! R1 2 3 are aligned
Stop
Dim Cr1Vpm   As Drs:       Cr1Vpm = SelDrs(CmlDclPm, "V Pm")             ' V Pm         ! each V is calling what Pm.  It is less than CrVr
Dim CrVp     As Drs:         CrVp = SplitSSCol(Cr1Vpm, "Pm")             ' V Pm
Dim Cr1Vpr As Drs:    '     Cr1Vpr = JnDrs(CrVp, CrRetA, "V", "R1 R2 R3") ' V Pm R1 R2 R3
Dim Cr1Fst As Drs:    '     Cr1Fst = AddColzFst(Cr1Vpr, "V")              ' V Pm R1 R2 R3 Fst
Dim Cr1Rmk As Drs:    '     Cr1Rmk = F.Cr1Rmk(Cr1Fst)              ' V Pm R1 R2 R3 Fst Rmk
Dim Cr1RmkPm As Drs:  '   Cr1RmkPm = SelDrs(Cr1Rmk, "V Rmk")

Dim CrRetF As Drs:    '     CrRetF = AddColzFst(CrRetA, "V")                    ' V R1 R2 R3 Fst
Dim CrRetR As Drs:    '     CrRetR = F.CrRetR(CrRetF)                           ' V R1 R2 R3 Fst Rmk
Dim CrRmkRet As Drs:  '   CrRmkRet = SelDrs(CrRetR, "V Rmk")

'Dim Cr1Rmk As Drs: Cr1Rmk = AddDrs(CrRmkRet, CrRmkPm)
                                                                                                   
'Dim Cr1VRmk  As S1S2s:           Cr1VRmk = F.CrVRmk(CrVpr)              ' V RmkLines
'Dim Cr1CmNm  As Drs:             Cr1CmNm = SelDrs(CmlEpt, "V CmNm")     ' V CmNm
'Dim Cr1VCmNm As Dictionary: Set Cr1VCmNm = DiczDrsCC(CrCmNm)
'Dim Cr1Ept   As S1S2s:            Cr1Ept = MapS1(CrVRmk, CrVCmNm)       ' CmNm RmkLines
'BrwS1S2s CrEpt: Stop
Stop
'== Upd Chd-Rmk (Cr) ===================================================================================================
'   If any of the calling pm's rmk is changed, the chd-mth-rmk will be updated
Dim CrSel   As Drs:              CrSel = SelDrs(McR123, "V R1 R2 R3") ' V R1 R2 R3
Dim CrWiRmk As Drs:            CrWiRmk = F.CrWiRmk(CrSel)             ' V R1 R2 R3      ! rmv those R1 2 3 are blank
Dim CrVr()  As Vr:                CrVr = F.CrVr(CrWiRmk)              ' V [R1 R2 R3]    ! each V has what rmk
Dim CrVpm   As Drs:              CrVpm = SelDrs(CmlDclPm, "V Pm")     ' V Pm            ! each V is calling what Pm.  It is less than CrVr
Dim CrVpr() As Vpr:              CrVpr = F.CrVpr(CrVpm, CrVr)         ' V [R1 R2 R3] [V [R1 R2 R3]] ! = [V Ret Pm] Sam Cnt as CrVpm.
'                                                                                                   ! Putting Vpr.Pm to Vr accroding to Vpm
Dim CrVRmk  As S1S2s:           CrVRmk = F.CrVRmk(CrVpr)              ' V RmkLines
Dim CrCmNm  As Drs:             CrCmNm = SelDrs(CmlEpt, "V CmNm")     ' V CmNm
Dim CrVCmNm As Dictionary: Set CrVCmNm = DiczDrsCC(CrCmNm)
Dim CrEpt   As S1S2s:            CrEpt = MapS1(CrVRmk, CrVCmNm)       ' CmNm RmkLines
BrwS1S2s CrEpt: Stop

'== Stop if dif si
Dim Si1%:                            Si1 = SiVpr(CrVpr)
Dim Si2%:                            Si2 = CrEpt.N
Dim DifSi As Boolean:             DifSi = Si1 <> Si2
Dim Stop1:                                 If DifSi Then Stop

Dim CrAct As S1S2s: CrAct = MthRmkzNy(CmMd, Cm)
Dim CrChg As S1S2s: CrChg = F.CrChg(CrEpt, CrAct)              ' CmNm RmkLines ! Only those need to change
Dim OUpdCr:                 If IsUpd Then F.OUpdCr CrChg, CmMd

'== Rpt <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Rpt:
If IsRptzRpt(Rpt) Then
    Dim CclsMsg$: CclsMsg = CclsMsg
    Select Case True
    Case NoSf: CclsMsg = "No Static F"
    Case Not NoSf: CclsMsg = "Has Static F & Ccls exist: " & Ccls.Parent.Name
    Case Else: CclsMsg = "Has Static F & Ccls is created: " & Ccls.Parent.Name
    End Select
    Insp CSub, "Changes", _
        "EmRpt DblEqRmk Align BrwStmt EnsSf EnsCcls " & _
        "Crt-Chd-Mth Rpl-Mth-Brw Crt-Mth-Brw Upd-Chd-MthLin Rfh-Chd-Mth-Rmk", _
        StrzRpt(Rpt), _
        FmtDrs(DeLNewO), FmtDrs(McLNewO), FmtDrs(BsLNewO), FmtDrs(SfLNewO), CclsMsg, _
        CdNewCm, FmtDrs(MbLNewO), FmtDrs(MbNew), FmtDrs(CmlLNewO), FmtS1S2s(CrChg)
End If
'Insp CSub, "Cr", "CrVpr CrS1S2s", FmtDrs(CrVpr), FmtS1S2s(CrS1S2s): Stop
End Sub

Function RplMth(M As CodeModule, Mthn, NewL$) As Boolean
'Ret True if Rplaced
Dim Lno&: Lno = MthLnozMM(M, Mthn)
If Not HasMthzM(M, Mthn) Then
    RplMth = True
    M.AddFromString NewL '<===
    Exit Function
End If
Dim OldL$: OldL = MthLineszM(M, Mthn)
If OldL = NewL Then Exit Function
RplMth = True
RmvMth M, Mthn '<==
M.InsertLines Lno, NewL '<==
End Function

Sub AlignMthDim(Optional Rpt As EmRpt)
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Sub
AlignMthDimzML M, CMthLno, Rpt:=Rpt
End Sub

Function AddColzCV(A As Drs, C$, V) As Drs
Dim Dr, Dry()
For Each Dr In Itr(A.Dry)
    PushI Dr, V
    PushI Dry, Dr
Next
AddColzCV = AddColzFFDry(A, C, Dry)
End Function

Function AddColzFFDry(A As Drs, FF$, NewDry()) As Drs
AddColzFFDry = Drs(AddFF(A.Fny, FF), NewDry)
End Function
Function InsColzFFDry(A As Drs, FF$, NewDry()) As Drs
InsColzFFDry = Drs(SyzAdd(SyzSS(FF), A.Fny), NewDry)
End Function

Function AddColzFiller(A As Drs, CC$) As Drs
Dim O As Drs: O = A
Dim C
For Each C In SyzSS(CC)
    O = AddColzFillerC(O, C)
Next
AddColzFiller = O
End Function

Private Function AddColzFillerC(A As Drs, C) As Drs
'Fm   A
'Fm   C #ColumnNm.
'Ret  Drs{ <A> {F<C>} } ! Add a new column {F<C>} add end which is Filler-column
'                       ! Filler column of a given-column-A is an integer-columns with value = MaxWdt(col-A) - Len(cur-value-of-col-A)
If NoReczDrs(A) Then Stop
Dim W%: W = WdtzAy(StrColzDrs(A, C))
Dim I%: I = IxzAy(A.Fny, C)
Dim ODry(): ODry = A.Dry
Dim Dr, J&
For Each Dr In Itr(ODry)
    PushI Dr, W - Len(Dr(I))
    ODry(J) = Dr
    J = J + 1
Next
AddColzFillerC = Drs(AddFF(A.Fny, "F" & C), ODry)
End Function

Sub RmvMthzMNn(M As CodeModule, Mthnn)
Dim I
For Each I In TermAy(Mthnn)
    RmvMthzMN M, I
Next
End Sub

Sub RmvMth(M As CodeModule, Mthn)
RmvMthzMN M, Mthn
End Sub

Sub RmvMthzMN(M As CodeModule, Mthn)
With MthSC(M, Mthn)
    If .S2 > 0 Then M.DeleteLines .S2, .C2
    If .S1 > 0 Then M.DeleteLines .S1, .C1
End With
End Sub

Sub CpyMthAs(M As CodeModule, Mthn, AsMthn)
If HasMthzM(M, AsMthn) Then Inf CSub, "AsMth exist.", "Mdn FmMth AsMth", Mdn(M), Mthn, AsMthn: Exit Sub
End Sub
Sub BrwMd(Md As CodeModule)
If Md.CountOfLines = 0 Then BrwStr "No lines in Md[" & Mdn(Md) & "]": Exit Sub
Brw Src(Md), "Md" & Mdn(Md)
End Sub
Private Sub Z_RmvMthzMN()
Dim Md As CodeModule
Const Mthn$ = "ZZRmv1"
Dim Bef$(), Aft$()
Crt:
        Set Md = TmpMod
        ApdLines Md, LineszVbl("|'sdklfsdf||'dsklfj|Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
Tst:
        Bef = Src(Md)
        RmvMthzMN Md, Mthn
        Aft = Src(Md)

Insp:   Insp CSub, "RmvMth Test", "Bef RmvMth Aft", Bef, Mthn, Aft
Rmv:    RmvMd Md
End Sub


Sub MovMthzNM(Mthn, ToMdn)
MovMthzMNM CMd, Mthn, Md(ToMdn)
End Sub

Sub MovMthzMNM(Md As CodeModule, Mthn, ToMd As CodeModule)
CpyMth Mthn, Md, ToMd
RmvMthzMN Md, Mthn
End Sub

Function EmpFunLines$(FunNm)
EmpFunLines = FmtQQ("Function ?()|End Function", FunNm)
End Function

Function EmpSubLines$(Subn)
EmpSubLines = FmtQQ("Sub ?()|End Sub", Subn)
End Function
Sub AddSub(Subn)
ApdLines CMd, EmpSubLines(Subn)
JmpMth Subn
End Sub

Sub AddFun(FunNm)
ApdLines CMd, EmpFunLines(FunNm)
JmpMth FunNm
End Sub

Sub Z1()
Z_AlignMthDimzML
End Sub

Sub Z11()
Const Mdn$ = "QIde_B_MthOp"
Const Mthn$ = "AlignMthDimzML"
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&: L = MthLnozMM(M, Mthn)
AlignMthDimzML M, L
End Sub

Sub Z_AlignMthDimzML()
Const TMthn$ = "AlignMthDimzML"
Const TMdn$ = "QIde_B_MthOp"
Const TCclsNm$ = TMdn & "__" & TMthn
Const TmMdn$ = "ATmp"
Const TmCclsNm$ = TmMdn & "__" & TMthn

Dim FmM As CodeModule, ToM As CodeModule, M As CodeModule, MthLno&, Lno&
Dim OldL$, NewL$, NewLy$()
Dim S1 As Boolean, S2 As Boolean
'Cpy Mth
    EnsMod CPj, TmMdn
    Set FmM = Md(TMdn)
    Set ToM = Md(TmMdn)
    'NewL
        NewL = MthLineszM(FmM, TMthn)
        NewLy = SplitCrLf(NewL)
        If Not HasPfx(NewLy(1), "Static F") Then Stop
        NewLy(1) = "Static F As New ATmp__AlignMthDimzML"
        NewL = JnCrLf(NewLy)
    
    'Rpl
        S1 = RplMth(ToM, TMthn, NewL)
        'Debug.Print "CpyMth: "; S1

'Cpy Md
    EnsCls CPj, TmCclsNm
    Set FmM = Md(TCclsNm)
    Set ToM = Md(TmCclsNm)
    S2 = CpyMd(FmM, ToM)
    'Debug.Print "CpyMd: "; S2
    If S1 Or S2 Then MsgBox "Copied": Exit Sub

'Align
    Set M = Md(TMdn)
    MthLno = MthLnozMM(M, TMthn)
    ATmp.AlignMthDimzML M, MthLno, SkpChkSelf:=True, Rpt:=EiUpdAndRpt
End Sub

Function CpyMd(FmM As CodeModule, ToM As CodeModule) As Boolean
CpyMd = RplMd(ToM, SrcLines(FmM))
End Function

Function CpyMth(Mthn, FmM As CodeModule, ToM As CodeModule) As Boolean
Dim NewL$
'NewL
    NewL = MthLineszM(FmM, Mthn)
'Rpl
    CpyMth = RplMth(ToM, Mthn, NewL)
'
'Const CSub$ = CMod & "CpyMdMthToMd"
'Dim Nav(): ReDim Nav(2)
'GoSub BldNav: ThwIf_ObjNE Md, ToMd, CSub, "Fm & To md cannot be same", Nav
'If Not HasMthzM(Md, Mthn) Then GoSub BldNav: ThwNav CSub, "Fm Mth not Exist", Nav
'If HasMthzM(ToMd, Mthn) Then GoSub BldNav: ThwNav CSub, "To Mth exist", Nav
'ToMd.AddFromString MthLineszM(Md, Mthn)
'RmvMth Md, Mthn
'If Not IsSilent Then Inf CSub, FmtQQ("Mth[?] in Md[?] is copied ToMd[?]", Mthn, Mdn(Md), Mdn(ToMd))
'Exit Sub
'BldNav:
'    Nav(0) = "FmMd Mth ToMd"
'    Nav(1) = Mdn(Md)
'    Nav(2) = Mthn
'    Nav(3) = Mdn(ToMd)
'    Return

End Function
Function CpyMthAsVer(M As CodeModule, Mthn, Ver%) As Boolean
'Ret True if copied
Dim VerMthn$, NewL$, L$, OldL$
If Not HasMthzM(M, Mthn) Then Inf CSub, "No from-mthn", "Md Mthn", Mdn(M), Mthn: Exit Function
VerMthn = Mthn & "_Ver" & Ver
'NewL
    L = MthLineszM(M, Mthn)
    NewL = Replace(L, "Sub " & Mthn, "Sub " & VerMthn, Count:=1)
'Rpl
    CpyMthAsVer = RplMth(M, VerMthn, NewL)

End Function

Private Function AlignMthDimzML__X()

End Function
