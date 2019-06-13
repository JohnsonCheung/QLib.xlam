Attribute VB_Name = "QDao_Schm_Er"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Schm_Er."
Private Const Asm$ = "QDao"
'MDF 'M.sg D.es of F.ld
'MDT 'M.sg D.es of T.bl
Const M000$ = "Lno#[?] ?"
Const MDF_NTermShouldBe3OrMore$ = "Should have 3 or more terms"
Const MDF_InvalidFld$ = "Invalid-Fld[?] Vdt-Fld[?]"
Const MDT_$ = ""
Const CDT_Tbl_NotIn_Tny$ = "T[?] is invalid.  Valid T[?]"
Const MDupE$ = "This E[?] is dup"
Const CM_LinTyEr$ = "Invalid DaoTy[?].  Valid Ty[?]"
Function ClnLin(Lin)
If IsEmp(Lin) Then Exit Function
If IsDotLin(Lin) Then Exit Function
If IsSngTermLin(Lin) Then Exit Function
If IsDDLin(Lin) Then Exit Function
ClnLin = BefDD(Lin)
End Function

Function ClnLnxs(Ly$()) As Lnxs
Dim L$, J&
For J = 0 To UB(Ly)
    L = ClnLin(Ly(J))
    If L <> "" Then
        PushLnx ClnLnxs, Lnx(L, J)
    End If
Next
End Function

Function ErzSchm(Schm$()) As String()
'========================================
Dim E As Lnxs
Dim F As Lnxs
Dim D As Lnxs
Dim T As Lnxs
    Dim X  As Lnxs
    X = ClnLnxs(Schm)
    T = LnxswRmvgT1(X, "Tbl")
    D = LnxswRmvgT1(X, "Des")
    E = LnxswRmvgT1(X, "Ele")
    F = LnxswRmvgT1(X, "Fld")

Dim Tny$(), Eny$()
    Eny = T1Ay(LyzLnxs(E))
    Tny = T1Ay(LyzLnxs(T))
'========================================
Dim AllFny$()
Dim AllEny$()
    AllFny = FnyzTdLy(LyzLnxs(T))
    AllEny = T1Ay(LyzLnxs(E))
ErzSchm = AddAyAp( _
    ErzLnxsT1ss(X, "Des Ele Fld Tbl"), _
    ErT(Tny, T, E), _
    ErF(AllFny, F), _
    ErE(E), _
    ErD(Tny, T, D))
End Function

Private Function ErD_LinEr(A As Lnxs) As String()
Dim I
'For Each I In Itr(A)
'    If RmvTT(CvLnx(I).Lin) = "" Then PushI ErD_LinEr, MsgD_NTermShouldBe3OrMore(CvLnx(I))
'Next
End Function

Private Function ErDT_InvalidFld(D As Lnxs, T$, Fny$()) As String()
'Fny$ is the fields in T$.
'D-Lnx.Lin is Des $T $F $D.  For the T$=$T line, the $F not in Fny$(), it is error
Dim I
'For Each I In Itr(D)
'    PushNonBlank ErDT_InvalidFld, ErDT_InvalidFld1(CvLnx(I), T, Fny)
'Next
End Function

Private Function ErDT_InvalidFld1$(D As Lnx, T$, Fny$())
Dim Tbl$, Fld$, Des$, X$
AsgN3tRst D.Lin, X, Tbl, Fld, Des
If X <> "Des" Then Stop
If Tbl <> T Then Exit Function

If Not HasEle(Fny, Fld) Then
    ErDT_InvalidFld1 = MsgDF_InvalidFld(D, Fld, Fny)
End If
End Function

Private Function ErDF_Er(A As Lnxs, T As Lnxs) As String() _
'Given A is D-Lnx having fmt = $Tbl $Fld $D, _
'This Sub checks if $Fld is in Fny
Dim J%, Fld$, Tbl$, Tny$(), FnyAy$()

'For J = 0 To UB(A)
    'AsgN2t A(J).Lin, Tbl, Fld
    If Tbl <> "." Then
'        PushIAy ErDF_Er, ErDF_Er1(A(J), Tny, FnyAy)
    End If
'Next
End Function

Private Function ErDF_Er1(A As Lnx, Tny$(), FnyAy$()) As String()
End Function

Private Function ErDT_Tbl_NotIn_Tny1$(D As Lnx, Tny$())
End Function

Private Function ErDT_Tbl_NotIn_Tny(D As Lnxs, Tny$()) As String()
'D-Lnx.Lin is Des $T $F $D. If $T<>"." and not in Tny, it is error
Dim J%
'For J = 0 To UB(D)
'    PushNonBlank ErDT_Tbl_NotIn_Tny, ErDT_Tbl_NotIn_Tny1(D(J), Tny)
'Next
End Function

Private Function ErE_DupE(E As Lnxs, Eny$()) As String()
Dim Ele
For Each Ele In Itr(AywDup(Eny))
    Push ErE_DupE, MsgE_DupE(LnoAyzEle(E, Ele), Ele)
Next
End Function

Private Function ErE_ELnx(A As Lnx) As String()
Dim L$, Ty$
Dim E$  ' EleNm of the Ele-Lin
L = A.Lin
'    AsgAp ShfVy(L, "*Nm *Ty ?Req ?ZLen TxtSz VTxt Dft VRul Expr"), _
                     .E, Ty, .Req, .ZLen, .TxtSz, .VTxt, .Dft, .VRul, .Expr
    '.Ty = DaoTy(Ty)
    If L <> "" Then
        Dim ExcessEle$
        Push ErE_ELnx, MsgE_ExcessEleItm(A, ExcessEle)
    End If
'    If .Ty = 0 Then
'        Push OEr, Msg_LinTyEr(A.Ix, Ty)
'    End If
End Function

Private Function ErE_ELnxs(A As Lnxs) As String()

End Function

Private Function ErT_FldHasNoEle(T As Lnxs, E As Lnxs) As String()

End Function

Private Function ErF_Ele_NotIn_Eny(F As Lnxs, Eny$()) As String()
Dim J%, O$(), Eless$, E$
'For J = 0 To UB(F)
    'With F(J)
'        E = T1(F(J).Lin)
'        If Not HasEle(Eny, E) Then PushI ErF_Ele_NotIn_Eny, MsgF_Ele_NotIn_Eny(F(J), E, Eless)
    'End With
'Next
ErF_Ele_NotIn_Eny = O
End Function

Private Function ErF_EleHasNoDef(F As Lnxs, AllEny$()) As String() _

End Function

Private Function Er_1_OneLiner(F As Lnx) As String()
Dim LikFF$, A$, V$
'    AsgAp Sy4N3TRst(F.Lin), .E, .LikT, V, A
'    .LikFny = SyzSS(LikFF)
End Function

Private Function ErF_1_LinEr(A As Lnxs) As String()
Dim J%
'For J = 0 To UB(A)
'    PushIAy ErF_1_LinEr, Er_1_OneLiner(A(J))
'Next
End Function

Private Function ErT_DupTbl(T As Lnxs, Tny$()) As String()
Dim Tbl$, I
For Each I In Itr(AywDup(Tny))
    Tbl = I
    Push ErT_DupTbl, MsgT_DupT(LnoAyzTbl(T, Tbl), Tbl)
Next
End Function

Private Function ErT_NoTLin(A As Lnxs) As String()
'If Si(A) > 0 Then Exit Function
PushI ErT_NoTLin, MsgT_NoTLin
End Function

Private Function ErT_LinEr_zLnx(T As Lnx) As String()
Dim L$
Dim Tbl$
    L = T.Lin
    Tbl = ShfT1(L)
    L = Replace(L, "*", Tbl)
'1
Select Case SubStrCnt(L, "|")
Case 0, 1
Case Else: PushI ErT_LinEr_zLnx, MsgT_Vbar_Cnt(T): Exit Function
End Select

'2
If Not IsNm(Tbl) Then
    PushI ErT_LinEr_zLnx, MsgT_TblIsNotNm(T)
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
        PushI ErT_LinEr_zLnx, MsgT_IdFld(T)
        Exit Function
    End If
End If
'5
    Dim Dup$()
    Dup = AywDup(Fny)
    If Si(Dup) > 0 Then
        PushI ErT_LinEr_zLnx, MsgT_DupF(T, Tbl, Dup)
        Exit Function
    End If
'6
If Si(Fny) = 0 Then
    PushI ErT_LinEr_zLnx, MsgT_NoFld(T)
    Exit Function
End If
'7
Dim F$, I
For Each I In Itr(Fny)
    F = I
    If Not IsNm(F) Then
        PushI ErT_LinEr_zLnx, MsgT_FldIsNotANmEr(T, F)
    End If
Next
End Function

Private Function ErT_LinEr(A As Lnxs) As String()
Dim I
'For Each I In Itr(A)
'    PushIAy ErT_LinEr, ErT_LinEr_zLnx(CvLnx(I))
'Next
End Function

Private Function LnoAyzEle(E As Lnxs, Ele) As Long()
Dim J%
'For J = 0 To UBound(E)
'    If T1(E(J).Lin) = Ele Then
'        PushI LnoAyzEle, E(J).Ix + 1
'    End If
'Next
End Function

Private Function LnoAyzTbl(A As Lnxs, T) As Long()
Dim J%
'For J = 0 To UB(A)
'    If T1(A(J).Lin) = T Then
'        PushI LnoAyzTbl, A(J).Ix + 1
'    End If
'Next
End Function

Private Function MsgMultiLno(LnoAy&(), M$)
MsgMultiLno = FmtQQ(M000, JnSpc(LnoAy), M)
End Function

Private Function Msg$(A As Lnx, M$)
Msg = FmtQQ("Lno#? ?", A.Ix, M)
End Function

Private Function MsgDF_InvalidFld$(ErLin As Lnx, ErFld$, VdtFny$())
MsgDF_InvalidFld = Msg(ErLin, FmtQQ(MDF_InvalidFld, ErFld, TLin(VdtFny)))
End Function

Private Function MsgD_NTermShouldBe3OrMore$(D As Lnx)
MsgD_NTermShouldBe3OrMore = Msg(D, MDF_NTermShouldBe3OrMore)
End Function

Private Function MsgDT_Tbl_NotIn_Tny$(A As Lnx, T, Tblss$)
MsgDT_Tbl_NotIn_Tny = Msg(A, FmtQQ(CDT_Tbl_NotIn_Tny, T, Tblss))
End Function

Private Function MsgE_DupE$(LnoAy&(), E)
MsgE_DupE = MsgMultiLno(LnoAy, FmtQQ(M000, E))
End Function

Private Function MsgT_DupT$(LnoAy&(), T)
MsgT_DupT = MsgMultiLno(LnoAy, FmtQQ("This Tbl[?] is dup", T))
End Function

Private Function MsgE_ExcessEleItm$(A As Lnx, ExcessEle$)
MsgE_ExcessEleItm = Msg(A, FmtQQ("Excess Ele Item [?]", ExcessEle))
End Function

Private Function MsgF_ExcessTxtSz$(A As Lnx)
MsgF_ExcessTxtSz = Msg(A, "Non-Txt-Ty should not have TxtSz")
End Function

Private Function MsgF_Ele_NotIn_Eny$(A As Lnx, E$, Eless$)
MsgF_Ele_NotIn_Eny = Msg(A, FmtQQ("Ele of is not in F-Lin not in Eny", E, Eless))
End Function

Private Function MsgE_FldEleEr$(A As Lnx, E$, Eless$)
MsgE_FldEleEr = Msg(A, FmtQQ("E[?] is invalid.  Valid E is [?]", E, Eless))
End Function

Private Function MsgFzDLy_NotIn_Fny$(A As Lnx, T$, F$, Fssl$)
MsgFzDLy_NotIn_Fny = Msg(A, FmtQQ("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function MsgT_DupF$(A As Lnx, T$, Fny$())
MsgT_DupF = Msg(A, FmtQQ("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function

Private Function MsgT_FldIsNotANmEr$(A As Lnx, F)
MsgT_FldIsNotANmEr = Msg(A, FmtQQ("FldNm[?] is not a name", F))
End Function

Private Function MsgT_IdFld$(A As Lnx)
Const M$ = "The field before first | must be *Id field"
MsgT_IdFld = Msg(A, M)
End Function

Private Function MsgT_NoFld(A As Lnx)
MsgT_NoFld = Msg(A, "No field")
End Function

Private Property Get MsgT_NoTLin()
MsgT_NoTLin = "No T-Line"
End Property

Private Function MsgT_TblIsNotNm$(A As Lnx)
MsgT_TblIsNotNm = Msg(A, "Tbl is not a name")
End Function

Private Function MsgT_Vbar_Cnt$(A As Lnx)
Const M$ = "The T-Lin should have 0 or 1 Vbar only"
MsgT_Vbar_Cnt = Msg(A, M)
End Function

Private Function MsgT_FldEr$(A As Lnx, F$)
MsgT_FldEr = Msg(A, FmtQQ("Fld[?] cannot be found in any Ele-Lines"))
End Function

Private Function Msg_LinTyEr$(A As Lnx, Ty$)
Msg_LinTyEr = Msg(A, FmtQQ(CM_LinTyEr, Ty, FmtDrs(DShtTy)))
End Function

Private Function ErD(Tny$(), T As Lnxs, D As Lnxs) As String()
ErD = AddAyAp( _
    ErD_LinEr(D))
    'ErD_FldEr(D, T))
    '    ErDT_InvalidFld(D, Tny), _

End Function

Private Function ErE(E As Lnxs) As String()
Dim Eny$()
ErE = AddAy(ErE_ELnxs(E), ErE_DupE(E, Eny))
End Function

Private Function ErF(AllEny$(), F As Lnxs) As String()
ErF = AddAy(ErF_1_LinEr(F), ErF_EleHasNoDef(F, AllEny))
End Function

Private Function ErT(Tny$(), T As Lnxs, E As Lnxs) As String()
ErT = AddAyAp( _
ErT_LinEr(T), _
ErT_NoTLin(T), _
ErT_FldHasNoEle(T, E), _
ErT_DupTbl(T, Tny))
End Function

Private Sub Z_ErT_LinEr_zLnx()
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
Dim TLnx As Lnx
Cas0:
    TLnx = Lnx(999, "Tbl 1")
    Ept = Sy("--- #1000[Tbl 1] FldNm[1] is not a name")
    GoTo Tst
Cas1:
    TLnx = Lnx(999, "A")
    Push EptEr, "should have a |"
    Ept = Sy("")
    GoTo Tst
Cas2:
    TLnx.Lin = "A | B B"
    Ept = Sy("")
    Push EptEr, "dup fields[B]"
    GoTo Tst
Cas3:
    TLnx.Lin = "A | B B D C C"
    Ept = Sy("")
    Push EptEr, "dup fields[B C]"
    GoTo Tst
Cas4:
    TLnx.Lin = "A | * B D C"
    Ept = Sy("")
    With Ept
        .T = "A"
        .Fny = SyzSS("A B D C")
    End With
    GoTo Tst
Cas5:
    TLnx.Lin = "A | * B | D C"
    Ept = Sy("")
    With Ept
        .T = "A"
        .Fny = SyzSS("A B D C")
        .Sk = SyzSS("B")
    End With
    GoTo Tst
Cas6:
    TLnx.Lin = "A |"
    Ept = Sy("")
    Push EptEr, "should have fields after |"
    GoTo Tst
Tst:
    Act = ErT_LinEr_zLnx(TLnx)
    C
    Return
End Sub

Private Sub ZZ()
Z_ErT_LinEr_zLnx
Exit Sub
'AAAA
'SchmLyEr
End Sub

