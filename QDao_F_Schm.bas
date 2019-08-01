Attribute VB_Name = "QDao_F_Schm"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Schm."
Const C_Tbl$ = "Tbl"
Const C_Fld$ = "Fld"
Const C_Ele$ = "Ele"
Const C_DesFld$ = "Des.Fld"
Const C_DesTbl$ = "Des.Tbl"
Type Ef
    EleLy() As String
    FldLy() As String
End Type
Type DoLLin
    D As Drs ' L Lin
End Type
Private Type FdRslt
    Som As Boolean
    Fd As Dao.Field2
End Type
Function DyoLLinzLy(Ly$()) As Variant()
Dim L, Lno&: For Each L In Itr(Ly)
    Lno = Lno + 1
    PushI DyoLLinzLy, Array(Lno, L)
Next
End Function
Function DoLLin(DyoLLin()) As Drs
DoLLin.Dy = DyoLLin
DoLLin.Fny = SyzSS("L Lin")
End Function

'Const M000$ = "Lno#[?] ?"
'Const MDF_NTermShouldBe3OrMore$ = "Should have 3 or more terms"
'Const MdfInvalidFld$ = "Invalid-Fld[?] Vdt-Fld[?]"
'Const MDT_$ = ""
'Const CDT_Tbl_NotIn_Tny$ = "T[?] is invalid.  Valid T[?]"
'Const MDupE$ = "This E[?] is dup"
'Const CM_LinTyEr$ = "Invalid DaoTy[?].  Valid Ty[?]"
':FunPfx-Mf:  :FunNmRul #Msg-Fld#      ! this error msg comes from Fld
':FunPfx-Mdt: :FunNmRul #Msg-Des-Tbl#  ! des-tbl
':FunPfx-Mdf: :FunNmRul #Msg-Des-Fld#  ! des-fld
':FunPfx-Me:  :FunNmRul #Msg-Ele#      ! ele
':FunPfx-Mt:  :FunNmRul #Msg-Tbl#      ! tbl
':FUnPFx-Ms:  :FunNmRul #Msg-Schm#     ! schm
Function ClnLin$(Lin)
If IsEmp(Lin) Then Exit Function
If FstChr(Lin) = "." Then Exit Function
If IsLinSngTerm(Lin) Then Exit Function
If IsLinDD(Lin) Then Exit Function
ClnLin = BefDD(Lin)
End Function
Function ClnDrs(Ly$()) As Drs

End Function
Function ErzLTDH(A As DoLTDH, T1ss$) As String()

End Function
Function ErzSchm(Schm$()) As String()
Dim X As DoLTDH:    X = DoLTDH(Schm)
Dim XD As Drs:     XD = X.D
Dim E As Drs:       E = DwEqExl(XD, "T1", "Ele")
Dim F As Drs:       F = DwEqExl(XD, "T1", "Fld")
Dim D As Drs:       D = DwEqExl(XD, "T1", "Des")
Dim T As Drs:       T = DwEqExl(XD, "T1", "Tbl")
Dim Tny$():       Tny = StrColzFst(T)
Dim Eny$():       Eny = StrColzFst(E)
Dim AllFny$(): AllFny = FnyzTdLy(StrColzSnd(T))
Dim AllEny$(): AllEny = StrColzFst(E)
'=======================================================================================================================
ErzSchm = AddAyAp( _
    ErzLTDH(X, "Des Ele Fld Tbl"), _
    Et(Tny, T, E), _
    Ef(AllFny, F), _
    Ee(E), _
    ErD(Tny, T, D))
End Function

Private Function EdLinEr(A As Drs) As String()
Dim I
'For Each I In Itr(A)
'    If RmvTT(CvLnx(I).Lin) = "" Then PushI EdLinEr, MdLin3T(CvLnx(I))
'Next
End Function

Private Function EdtInvalidFld(D As Drs, T$, Fny$()) As String()
'Fny$ is the fields in T$.
'D-Lnx.Lin is Des $T $F $D.  For the T$=$T line, the $F not in Fny$(), it is error
Dim I
'For Each I In Itr(D)
'    PushNB EdtInvalidFld, EdtInvalidFld1(CvLnx(I), T, Fny)
'Next
End Function

Private Function EdtInvalidFld1$(D As LLin, T$, Fny$())
Dim Tbl$, Fld$, Des$, X$
Asg3TRst D.Lin, X, Tbl, Fld, Des
If X <> "Des" Then Stop
If Tbl <> T Then Exit Function

If Not HasEle(Fny, Fld) Then
'    EdtInvalidFld1 = MdfInvalidFld(D, Fld, Fny)
End If
End Function

Private Function EdfEr(A As Drs, T As Drs) As String() _
'Given A is D-Lnx having fmt = $Tbl $Fld $D, _
'This Sub checks if $Fld is in Fny
Dim J%, Fld$, Tbl$, Tny$(), FnyAy$()

'For J = 0 To UB(A)
    'AsgTT A(J).Lin, Tbl, Fld
    If Tbl <> "." Then
'        PushIAy EdfEr, EdfEr1(A(J), Tny, FnyAy)
    End If
'Next
End Function

Private Function EdfEr1(A(), Tny$(), FnyAy$()) As String()
End Function

Private Function EdtTblNin1$(D(), Tny$())
End Function

Private Function EdtTblNin(D As Drs, Tny$()) As String()
'D-Lnx.Lin is Des $T $F $D. If $T<>"." and not in Tny, it is error
Dim J%
'For J = 0 To UB(D)
'    PushNB EdtTblNin, EdtTblNin1(D(J), Tny)
'Next
End Function

Private Function EdEleDup(EleLLno As Drs, Eny$()) As String()
Dim Ele
For Each Ele In Itr(AwDup(Eny))
'    Push EdEleDup, MeEleDup(LnoAyzEle(E, Ele), Ele)
Next
End Function

Private Function EeTermEr(EleLLin As Drs) As String()
Dim L$, Ty$
Dim E$  ' EleNm of the Ele-Lin
'L = A.Lin
'    AsgAp ShfVy(L, "*Nm *Ty ?Req ?ZLen TxtSz VTxt Dft VRul Expr"), _
                     .E, Ty, .Req, .ZLen, .TxtSz, .VTxt, .Dft, .VRul, .Expr
    '.Ty = DaoTy(Ty)
    If L <> "" Then
        Dim ExcessEle$
'        Push EeTermEr, MeEleExc(A, ExcessEle)
    End If
'    If .Ty = 0 Then
'        Push OEr, Msg_LinTyEr(A.Ix, Ty)
'    End If
End Function

Private Function EeTermErs(A As Drs) As String()

End Function

Private Function EtFldNEle(T As Drs, E As Drs) As String()

End Function

Private Function EfEleNin(F As Drs, Eny$()) As String()
Dim J%, O$(), Eless$, E$
'For J = 0 To UB(F)
    'With F(J)
'        E = T1(F(J).Lin)
'        If Not HasEle(Eny, E) Then PushI EfEleNin, MfEle_NotIn_Eny(F(J), E, Eless)
    'End With
'Next
EfEleNin = O
End Function

Private Function EfEleHasNoDef(F As Drs, AllEny$()) As String() _

End Function

Private Function Er_1_OneLiner(F()) As String()
Dim LikFF$, A$, V$
'    AsgAp Sy4N3TRst(F.Lin), .E, .LikT, V, A
'    .LikFny = SyzSS(LikFF)
End Function

Private Function Ef1_LinEr(A As Drs) As String()
Dim J%
'For J = 0 To UB(A)
'    PushIAy Ef1_LinEr, Er_1_OneLiner(A(J))
'Next
End Function

Private Function EtDupTbl(T As Drs, Tny$()) As String()
Dim Tbl$, I
For Each I In Itr(AwDup(Tny))
    Tbl = I
    Push EtDupTbl, MtDupT(LnoAyzTbl(T, Tbl), Tbl)
Next
End Function

Private Function EtNoTLin(T As Drs) As String()
'If Si(A) > 0 Then Exit Function
PushI EtNoTLin, MtNoTLin
End Function

Private Function EtLinErzDr(T_DroLLin()) As String()
Dim T
Dim Lno&: Lno = T_DroLLin(0)
Dim L$:     L = T_DroLLin(1)
Dim Tbl$
'    L = T.Lin
    Tbl = ShfT1(L)
    L = Replace(L, "*", Tbl)
'1
Select Case SubStrCnt(L, "|")
Case 0, 1
Case Else: PushI EtLinErzDr, MtVbar_Cnt(T): Exit Function
End Select

'2
If Not IsNm(Tbl) Then
    PushI EtLinErzDr, MtTblNNm(T)
    Exit Function
End If
'
Dim Fny$()
    Fny = SyzSS(Replace(L, "|", " "))
    
If HasSubStr(L, "|") Then
'3
    Dim IdFld$
    IdFld = Trim(Bef(L, "|"))
    If IdFld <> Tbl & "Id" Then
        PushI EtLinErzDr, MtIdMis(Lno)
        Exit Function
    End If
End If
'5
    Dim Dup$()
    Dup = AwDup(Fny)
    If Si(Dup) > 0 Then
'        PushI EtLinErzDr, MtFldDup(L, Tbl, Dup)
        Exit Function
    End If
'6
If Si(Fny) = 0 Then
'    PushI EtLinEr_zLnx, MtFldMis(T)
    Exit Function
End If
'7
Dim F$, I
For Each I In Itr(Fny)
    F = I
    If Not IsNm(F) Then
 '       PushI EtLinEr_zLnx, MtFldIsNotANm(T, F)
    End If
Next
End Function

Private Function EtLinEr(T As DoLLin) As String()
Dim I
'For Each I In Itr(A)
'    PushIAy EtLinEr zDr, EtLinEr_zLnx(CvLnx(I))
'Next
End Function

Private Function LnoAyzEle(E As Drs, Ele) As Long()
Dim J%
'For J = 0 To UBound(E)
'    If T1(E(J).Lin) = Ele Then
'        PushI LnoAyzEle, E(J).Ix + 1
'    End If
'Next
End Function

Private Function LnoAyzTbl(A As Drs, T) As Long()
Dim J%
'For J = 0 To UB(A)
'    If T1(A(J).Lin) = T Then
'        PushI LnoAyzTbl, A(J).Ix + 1
'    End If
'Next
End Function

Private Function MsgMultiLno(LnoAy&(), M$)
'MsgMultiLno = FmtQQ(M000, JnSpc(LnoAy), M)
End Function

Private Function Msg$(Lno&, M$)
Msg = FmtQQ("Lno#? ?", Lno, M)
End Function

Private Function MdfInvalidFld$(ErLin(), Efld$, VdtFny$())
'MdfInvalidFld = Msg(ErLin, FmtQQ(MdfInvalidFld, Efld, TLin(VdtFny)))
End Function

Private Function MdLin3T$(D())
'MdLin3T = Msg(D, MDF_NTermShouldBe3OrMore)
End Function

Private Function MdtTblNin$(D As DoLLin, T, Tblss$)
'MdtTblNin = Msg(A, FmtQQ(CDT_Tbl_NotIn_Tny, T, Tblss))
End Function

Private Function MeEleDup$(LnoAy&(), E)
'MeEleDup = MsgMultiLno(LnoAy, FmtQQ(M000, E))
End Function

Private Function MtDupT$(LnoAy&(), T)
MtDupT = MsgMultiLno(LnoAy, FmtQQ("This Tbl[?] is dup", T))
End Function

Private Function MeEleExc$(A(), ExcessEle$)
'MeEleExc = Msg(A, FmtQQ("Excess Ele Item [?]", ExcessEle))
End Function

Private Function MfExcessTxtSz$(A())
'MfExcessTxtSz = Msg(A, "Non-Txt-Ty should not have TxtSz")
End Function

Private Function MfEle_NotIn_Eny$(A(), E$, Eless$)
'MfEle_NotIn_Eny = Msg(A, FmtQQ("Ele of is not in F-Lin not in Eny", E, Eless))
End Function

Private Function MeEleNin$(A(), E$, Eless$)
'MeEleNin = Msg(A, FmtQQ("E[?] is invalid.  Valid E is [?]", E, Eless))
End Function

Private Function MdfFldNin$(A(), T$, F$, Fssl$)
'MdfFldNin = Msg(A, FmtQQ("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function MtFldDup$(A(), T$, Fny$())
'MtFldDup = Msg(A, FmtQQ("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function

Private Function MtFldIsNotANm$(A(), F)
'MtFldIsNotANm = Msg(A, FmtQQ("FldNm[?] is not a name", F))
End Function

Private Function MtIdMis$(Lno&)
Const M$ = "The field before first | must be *Id field"
'MtIdMis = Msg(A, M)
End Function

Private Function MtFldMis(A())
'MtFldMis = Msg(A, "No field")
End Function

Private Property Get MtNoTLin()
MtNoTLin = "No T-Line"
End Property

Private Function MtTblNNm$(A)
'MtTblNNm = Msg(A, "Tbl is not a name")
End Function

Private Function MtVbar_Cnt$(T)
Const M$ = "The T-Lin should have 0 or 1 Vbar only"
'MtVbar_Cnt = Msg(A, M)
End Function

Private Function MtFldEr$(A(), F$)
'MtFldEr = Msg(A, FmtQQ("Fld[?] cannot be found in any Ele-Lines"))
End Function

Private Function Msg_LinTyEr$(A(), Ty$)
'Msg_LinTyEr = Msg(A, FmtQQ(CM_LinTyEr, Ty, FmtDrs(DShtTy)))
End Function

Private Function ErD(Tny$(), T As Drs, D As Drs) As String()
ErD = AddAyAp( _
    EdLinEr(D), _
    EdFldEr(D, T))
    '    EdtInvalidFld(D, Tny), _

End Function
Private Function EdFldEr(D As Drs, T As Drs) As String()

End Function
Private Function Ee(E As Drs) As String()
Dim Eny$()
Ee = AddAy(EeTermErs(E), EdEleDup(E, Eny))
End Function

Private Function Ef(AllEny$(), F As Drs) As String()
Ef = AddAy(Ef1_LinEr(F), EfEleHasNoDef(F, AllEny))
End Function

Private Function Et(Tny$(), T As Drs, E As Drs) As String()
'Et = AddAyAp( _
EtLinEr(T), _
EtNoTLin(T), _
EtFldNEle(T, E), _
EtDupTbl(T, Tny))
End Function

Private Sub Z_EtLinEr_zLnx()
Dim T(): ReDim T(1): T(0) = 999
GoSub Cas0
Stop
GoSub Cas1
GoSub Cas2
GoSub Cas3
GoSub Cas4
GoSub Cas5
GoSub Cas6
Exit Sub
Dim EptEr$(), ActEr$()
Cas0:
    T(1) = "Tbl 1"
    Ept = Sy("--- #1000[Tbl 1] FldNm[1] is not a name")
    GoTo Tst
Cas1:
    T(1) = "A"
    Push EptEr, "should have a |"
    Ept = Sy("")
    GoTo Tst
Cas2:
    T(1) = "A | B B"
    Ept = Sy("")
    Push EptEr, "dup fields[B]"
    GoTo Tst
Cas3:
    T(1) = "A | B B D C C"
    Ept = Sy("")
    Push EptEr, "dup fields[B C]"
    GoTo Tst
Cas4:
    T(1) = "A | * B D C"
    Ept = Sy("")
    With Ept
        .T = "A"
        .Fny = SyzSS("A B D C")
    End With
    GoTo Tst
Cas5:
    T(1) = "A | * B | D C"
    Ept = Sy("")
    With Ept
        .T = "A"
        .Fny = SyzSS("A B D C")
        .Sk = SyzSS("B")
    End With
    GoTo Tst
Cas6:
    T(1) = "A |"
    Ept = Sy("")
    Push EptEr, "should have fields after |"
    GoTo Tst
Tst:
    Act = EtLinErzDr(T)
    C
    Return
End Sub

Private Sub Z()
Z_EtLinEr_zLnx
Exit Sub
'AAAA
'SchmLyEr
End Sub


Sub CrtSchmzVbl(D As Database, SchmVbl$)
CrtSchm D, SplitVBar(SchmVbl)
End Sub

Sub CrtSchm(D As Database, Schm$())
Const CSub$ = CMod & "CrtSchm"
ThwIf_ErMsg ErzSchm(Schm), CSub, "there is error in the Schm", "Schm Db", AddIxPfx(Schm, 1), D.Name
Dim X As DoLTDH:           X = DoLTDH(Schm)
Dim TdLy$():            TdLy = FmtDoLTDH(X, C_Tbl)
Dim E$():                  E = FmtDoLTDH(X, C_Ele)
Dim F$():                  F = FmtDoLTDH(X, C_Fld)
Dim DF$()
Dim DT$()
Dim T() As Dao.TableDef:   T = TdAy(TdLy, E, F)
Dim P$():                  P = SqyCrtPkzTny(PkTny(TdLy))
Dim S$():                  S = SqyCrtSk(TdLy)
Dim DicT As Dictionary: Set DicT = Dic(AwRmvTT(Schm, C_Des, C_Tbl))
Dim DicF As Dictionary: Set DicF = Dic(AwRmvTT(Schm, C_Des, C_Fld))
                   AppTdAy D, T
                   RunSqy D, P
                   RunSqy D, S
SetTblDesDic D, DicT
SetFldDesDic D, DicF
End Sub

Private Function EFzSchm(Schm$()) As Ef
EFzSchm.EleLy = AwRmvT1(Schm, "Ele")
EFzSchm.FldLy = AwRmvT1(Schm, "Fld")
End Function

Private Function PkTny(TdLy$()) As String()
Dim I, L$
For Each I In TdLy
    L = I
    If HasSubStr(L, " *Id ") Then
        PushI PkTny, T1(L)
    End If
Next
End Function

Private Function SqyCrtSk(TdLy$()) As String()
Dim TdLin, I, Sk$()
For Each I In Itr(TdLy)
    TdLin = I
    Sk = SkFny(TdLin)
    If Si(Sk) > 0 Then
        PushI SqyCrtSk, SqlCrtSk_T_SkFny(T1(TdLin), RplStarzAy(Sk, T1(TdLin)))
    End If
Next
End Function

Private Function SkFny(TdLin) As String()
Dim P%, T$, Rst$
P = InStr(TdLin, "|")
If P = 0 Then Exit Function
AsgTRst Bef(TdLin, "|"), T, Rst
Rst = Replace(Rst, T, "*")
SkFny = SyzSS(Rst)
End Function

Private Function TdAy(TdLy$(), E$(), F$()) As Dao.TableDef()
Dim I: For Each I In TdLy
    PushObj TdAy, TdzLin(I, E, F)
Next
End Function

Private Function TdzLin(TdLin, Ele$(), Fld$()) As Dao.TableDef
Dim T: T = T1(TdLin)
Dim O As Dao.TableDef: Set O = New Dao.TableDef
O.Name = T
Dim F, Fd As Dao.Field2
For Each F In FnyzTdLin(TdLin)
    If F = T & "Id" Then
        Set Fd = FdzPk(F)
    Else
        Set Fd = FdzEF(F, Ele, Fld)
    End If
    O.Fields.Append Fd
Next
Set TdzLin = O
End Function

Private Function FdzEF(F, Ele$(), Fld$()) As Dao.Field2
If Left(F, 2) = "Id" Then Stop
Dim E$: E = LookupT1(F, Fld)
If E <> "" Then Set FdzEF = FdzEle(E, Ele, F): Exit Function
Set FdzEF = FdzStdFldNm(F):                    If Not IsNothing(FdzEF) Then Exit Function
Set FdzEF = FdzEle(CStr(F), Ele, F):  If Not IsNothing(FdzEF) Then Exit Function
Thw CSub, FmtQQ("Fld(?) not in EF and not StdFld", F)
End Function

Private Function FdzEle(Ele$, EleLy$(), F) As Dao.Field2
Dim EStr$: EStr = EleStr(EleLy, Ele)
If EStr <> "" Then Set FdzEle = FdzFdStr(F & " " & EStr): Exit Function
Set FdzEle = FdzShtTys(Ele, F): If Not IsNothing(FdzEle) Then Exit Function
EStr = EleStr(EleLy, F)
If EStr <> "" Then Set FdzEle = FdzFdStr(F & " " & EStr): Exit Function
Set FdzEle = FdzShtTys(F, F)
Dim EleNy$(): EleNy = T1Ay(EleLy)
Thw CSub, FmtQQ("Fld(?) of Ele(?) not found in EleLy-of-EleAy(?) and not StdEle", F, Ele, TLin(EleNy))
End Function

Private Function EleStr$(EleLy$(), Ele)
EleStr = RmvT1(FstElezT1(EleLy, Ele))
End Function

Private Function EleStrzStd$(Ele)
End Function

Private Property Get Schm1() As String()
Erase XX
X "Tbl A *Id *Nm     | *Dte AATy Loc Expr Rmk"
X "Tbl B *Id  AId *Nm | *Dte"
X "Fld Txt AATy"
X "Fld Mem Rmk"
X "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
X "Ele Expr Txt [Expr=Loc & 'abc']"
X "Des Tbl  A     AA BB "
X "Des Tbl  A     CC DD "
X "Des Fld  N1   AA BB "
X "Des Fld  A.N1 TF_Des-AA-BB"
Schm1 = XX
Erase XX
End Property

Private Sub Z_CrtSchm()
Dim D As Database, Schm$()
GoSub T1
Exit Sub

T1:
    Set D = TmpDb
    Schm = Schm1
    GoTo Tst
Tst:
    CrtSchm D, Schm
    Dmp TdLyzDb(D)
    Return
End Sub

Sub AppTdAy(D As Database, TdAy() As Dao.TableDef)
Dim T
For Each T In Itr(TdAy)
    D.TableDefs.Append T
Next
End Sub

Function FnyzTdLin(TdLin) As String()
Dim T$, Rst$
AsgTRst TdLin, T, Rst
If HasSfx(T, "*") Then
    T = RmvSfx(T, "*")
    Rst = T & "Id " & Rst
End If
Rst = Replace(Rst, "*", T)
Rst = Replace(Rst, "|", " ")
FnyzTdLin = SyzSS(Rst)
End Function

Sub EnsSchm(D As Database, Schm$())
Stop
ThwIf_ErMsg ErzSchm(Schm), CSub, "there is error in the Schm"
'AppDbTdAy A, TdAy(Smt, AwRmvT1(Schm, CCF), AwRmvT1(Schm, CCE))
'RunSqy A, SqyCrtPk_Tny(PkTnySmt(Smt))
'RunSqy A, SqyCrtSkSmt(Smt)
'Set TblDesDic(A) = TblDesDicSmdt(AwRmvTT(Schm, CCD, CCT))
'Set TblDesDicDb(A) = TblDesDicDbSmdf(AwRmvTT(Schm, CCD, CCF))
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
X "Des Fld     N1   AA BB "
X "Des Tbl.Fld A.N1 TF_Des-AA-BB"
Z_CrtSchm1 = XX
Erase XX
End Property

