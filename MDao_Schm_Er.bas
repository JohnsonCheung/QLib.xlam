Attribute VB_Name = "MDao_Schm_Er"
Option Explicit
'MDF 'M.sg D.es of F.ld
'MDT 'M.sg D.es of T.bl
Const M000$ = "Lno#[?] ?"
Const M001$ = "Lno#? ?"
Const MDF_NTermShouldBe3OrMore$ = "Should have 3 or more terms"
Const MDF_InvalidFld$ = "Invalid-Fld[?] Vdt-Fld[?]"
Const MDT_$ = ""
Const CDT_Tbl_NotIn_Tny$ = "T[?] is invalid.  Valid T[?]"
Const MDupE$ = "This E[?] is dup"
Const CM_LinTyEr$ = "Invalid DaoTy[?].  Valid Ty[?]"

Function ErzSchm(Schm$()) As String()
'========================================
Dim E() As Lnx
Dim F() As Lnx
Dim D() As Lnx
Dim T() As Lnx
    Dim X()  As Lnx
    X = LnxAyzCln(Schm)
    T = LnxAywRmvT1(X, "Tbl")
    D = LnxAywRmvT1(X, "Des")
    E = LnxAywRmvT1(X, "Ele")
    F = LnxAywRmvT1(X, "Fld")

Dim Tny$(), Eny$()
    Eny = AyTakT1(LyzLnxAy(E))
    Tny = AyTakT1(LyzLnxAy(T))
'========================================
Dim AllFny$()
Dim AllEny$()
    AllFny = FnyzTdLy(LyzLnxAy(T))
    AllEny = AyTakT1(LyzLnxAy(E))
ErzSchm = AyAddAp( _
    ErzLnxAyT1ss(X, "Des Ele Fld Tbl"), _
    ErT(Tny, T, E), _
    ErF(AllFny, F), _
    ErE(E), _
    ErD(Tny, T, D))
End Function

Private Function ErD_LinEr(A() As Lnx) As String()
Dim I
For Each I In Itr(A)
    If RmvTT(CvLnx(I).Lin) = "" Then PushI ErD_LinEr, MsgD_NTermShouldBe3OrMore(CvLnx(I))
Next
End Function

Private Function ErDT_InvalidFld(D() As Lnx, T$, Fny$()) As String()
'Fny$ is the fields in T$.
'D-Lnx.Lin is Des $T $F $D.  For the T$=$T line, the $F not in Fny$(), it is error
Dim I
For Each I In Itr(D)
    PushNonBlankStr ErDT_InvalidFld, ErDT_InvalidFld1(CvLnx(I), T, Fny)
Next
End Function

Private Function ErDT_InvalidFld1$(D As Lnx, T$, Fny$())
Dim Tbl$, Fld$, Des$, X$
Asg3TRst D.Lin, X, Tbl, Fld, Des
If X <> "Des" Then Stop
If Tbl <> T Then Exit Function

If Not HasEle(Fny, Fld) Then
    ErDT_InvalidFld1 = MsgDF_InvalidFld(D, Fld, Fny)
End If
End Function

Private Function ErDF_Er(A() As Lnx, T() As Lnx) As String() _
'Given A is D-Lnx having fmt = $Tbl $Fld $D, _
'This Sub checks if $Fld is in Fny
Dim J%, Fld$, Tbl$, Tny$(), FnyAy$()

For J = 0 To UB(A)
    AsgTT A(J).Lin, Tbl, Fld
    If Tbl <> "." Then
        PushIAy ErDF_Er, ErDF_Er1(A(J), Tny, FnyAy)
    End If
Next
End Function

Private Function ErDF_Er1(A As Lnx, Tny$(), FnyAy$()) As String()
End Function

Private Function ErDT_Tbl_NotIn_Tny1$(D As Lnx, Tny$())
End Function

Private Function ErDT_Tbl_NotIn_Tny(D() As Lnx, Tny$()) As String()
'D-Lnx.Lin is Des $T $F $D. If $T<>"." and not in Tny, it is error
Dim J%
For J = 0 To UB(D)
    PushNonBlankStr ErDT_Tbl_NotIn_Tny, ErDT_Tbl_NotIn_Tny1(D(J), Tny)
Next
End Function

Private Function ErE_DupE(E() As Lnx, Eny$()) As String()
Dim Ele
For Each Ele In Itr(AywDup(Eny))
    Push ErE_DupE, MsgE_DupE(LnoAyzEle(E, Ele), Ele)
Next
End Function

Private Function ErE_ELnx(A As Lnx) As String()
Dim L$, Ty$
Dim E$  ' EleNm of the Ele-Lin
L = A.Lin
'    AsgAp VyzLinLbl(L, "*Nm *Ty ?Req ?ZLen TxtSz VTxt Dft VRul Expr"), _
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

Private Function ErE_ELnxAy(A() As Lnx) As String()

End Function

Private Function ErT_FldHasNoEle(T() As Lnx, E() As Lnx) As String()

End Function

Private Function ErF_Ele_NotIn_Eny(F() As Lnx, Eny$()) As String()
Dim J%, O$(), Eless$, E$
For J = 0 To UB(F)
    With F(J)
        E = T1(F(J).Lin)
        If Not HasEle(Eny, E) Then PushI ErF_Ele_NotIn_Eny, MsgF_Ele_NotIn_Eny(F(J), E, Eless)
    End With
Next
ErF_Ele_NotIn_Eny = O
End Function

Private Function ErF_EleHasNoDef(F() As Lnx, AllEny$()) As String() _

End Function

Private Function Er_1_OneLiner(F As Lnx) As String()
Dim LikFF$, A$, V$
'    AsgAp Sy43TRst(F.Lin), .E, .LikT, V, A
'    .LikFny = SySsl(LikFF)
End Function

Private Function ErF_1_LinEr(A() As Lnx) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy ErF_1_LinEr, Er_1_OneLiner(A(J))
Next
End Function

Private Function ErT_DupTbl(T() As Lnx, Tny$()) As String()
Dim Tbl
For Each Tbl In Itr(AywDup(Tny))
    Push ErT_DupTbl, MsgT_DupT(LnoAyzTbl(T, Tbl), Tbl)
Next
End Function

Private Function ErT_NoTLin(A() As Lnx) As String()
If Si(A) > 0 Then Exit Function
PushI ErT_NoTLin, MsgT_NoTLin
End Function

Private Function ErT_1_OneLinEr(T As Lnx) As String()
Dim L$
Dim Tbl$
    L = T.Lin
    Tbl = ShfTerm(L)
    L = Replace(L, "*", Tbl)
'1
Select Case SubStrCnt(L, "|")
Case 0, 1
Case Else: PushI ErT_1_OneLinEr, MsgT_Vbar_Cnt(T): Exit Function
End Select

'2
If Not IsNm(Tbl) Then
    PushI ErT_1_OneLinEr, MsgT_TblIsNotNm(T)
    Exit Function
End If
'
Dim Fny$()
    Fny = SySsl(Replace(L, "|", " "))
    
If HasSubStr(L, "|") Then
'3
    Dim IdFld$
    IdFld = Trim(StrBef(L, "|"))
    If IdFld <> Tbl & "Id" Then
        PushI ErT_1_OneLinEr, MsgT_IdFld(T)
        Exit Function
    End If
'4
    If Trim(StrBet(L, "|", "|")) = "" Then
        PushI ErT_1_OneLinEr, MsgT_NoFLdBetVV(T)
        Exit Function
    End If
End If
'5
    Dim Dup$()
    Dup = AywDup(Fny)
    If Si(Dup) > 0 Then
        PushI ErT_1_OneLinEr, MsgT_DupF(T, Tbl, Dup)
        Exit Function
    End If
'6
If Si(Fny) = 0 Then
    PushI ErT_1_OneLinEr, MsgT_NoFld(T)
    Exit Function
End If
'7
Dim F
For Each F In Itr(Fny)
    If Not IsNm(F) Then
        PushI ErT_1_OneLinEr, MsgT_FldIsNotANmEr(T, F)
    End If
Next
End Function

Private Function ErT_1_LinEr(A() As Lnx) As String()
Dim I
For Each I In Itr(A)
    PushIAy ErT_1_LinEr, ErT_1_OneLinEr(CvLnx(I))
Next
End Function

Private Function LnoAyzEle(E() As Lnx, Ele) As Long()
Dim J%
For J = 0 To UBound(E)
    If T1(E(J).Lin) = Ele Then
        PushI LnoAyzEle, E(J).Ix + 1
    End If
Next
End Function

Private Function LnoAyzTbl(A() As Lnx, T) As Long()
Dim J%
For J = 0 To UB(A)
    If T1(A(J).Lin) = T Then
        PushI LnoAyzTbl, A(J).Ix + 1
    End If
Next
End Function

Private Function WMsgMultiLno(LnoAy&(), M$)
WMsgMultiLno = FmtQQ(M000, JnSpc(LnoAy), M)
End Function

Private Function WMsg$(Ix, M$)
WMsg = FmtQQ(M001, Ix, M)
End Function

Private Function MsgDF_InvalidFld$(ErLin As Lnx, ErFld$, VdtFny$())
MsgDF_InvalidFld = WMsg(ErLin, FmtQQ(MDF_InvalidFld, ErFld, TLin(VdtFny)))
End Function

Private Function MsgD_NTermShouldBe3OrMore$(D As Lnx)
MsgD_NTermShouldBe3OrMore = WMsg(D, MDF_NTermShouldBe3OrMore)
End Function

Private Function MsgDT_Tbl_NotIn_Tny$(A As Lnx, T, Tblss$)
MsgDT_Tbl_NotIn_Tny = WMsg(A, FmtQQ(CDT_Tbl_NotIn_Tny, T, Tblss))
End Function

Private Function MsgE_DupE$(LnoAy&(), E)
MsgE_DupE = WMsgMultiLno(LnoAy, FmtQQ(M000, E))
End Function

Private Function MsgT_DupT$(LnoAy&(), T)
MsgT_DupT = WMsgMultiLno(LnoAy, FmtQQ("This Tbl[?] is dup", T))
End Function

Private Function MsgE_ExcessEleItm$(A As Lnx, ExcessEle$)
MsgE_ExcessEleItm = WMsg(A, FmtQQ("Excess Ele Item [?]", ExcessEle))
End Function

Private Function MsgF_ExcessTxtSz$(A As Lnx)
MsgF_ExcessTxtSz = WMsg(A, "Non-Txt-Ty should not have TxtSz")
End Function

Private Function MsgF_Ele_NotIn_Eny$(A As Lnx, E$, Eless$)
MsgF_Ele_NotIn_Eny = WMsg(A, FmtQQ("Ele of is not in F-Lin not in Eny", E, Eless))
End Function

Private Function MsgE_FldEleEr$(A As Lnx, E$, Eless$)
MsgE_FldEleEr = WMsg(A, FmtQQ("E[?] is invalid.  Valid E is [?]", E, Eless))
End Function

Private Function MsgFzDLy_NotIn_Fny$(A As Lnx, T$, F$, Fssl$)
MsgFzDLy_NotIn_Fny = WMsg(A, FmtQQ("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function MsgT_DupF$(A As Lnx, T$, Fny$())
MsgT_DupF = WMsg(A, FmtQQ("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function

Private Function MsgT_FldIsNotANmEr$(A As Lnx, F)
MsgT_FldIsNotANmEr = WMsg(A, FmtQQ("FldNm[?] is not a name", F))
End Function

Private Function MsgT_IdFld$(A As Lnx)
Const M$ = "The field before first | must be *Id field"
MsgT_IdFld = WMsg(A, M)
End Function

Private Function MsgT_NoFld(A As Lnx)
MsgT_NoFld = WMsg(A, "No field")
End Function

Private Function MsgT_NoFLdBetVV$(A As Lnx)
MsgT_NoFLdBetVV = WMsg(A, "No field between | |")
End Function

Private Property Get MsgT_NoTLin$()
MsgT_NoTLin = "No T-Line"
End Property

Private Function MsgT_TblIsNotNm$(A As Lnx)
MsgT_TblIsNotNm = WMsg(A, "Tbl is not a name")
End Function

Private Function MsgT_Vbar_Cnt$(A As Lnx)
Const M$ = "The T-Lin should have 0 or 1 Vbar only"
MsgT_Vbar_Cnt = WMsg(A, M)
End Function

Private Function MsgT_FldEr$(A As Lnx, F$)
MsgT_FldEr = WMsg(A, FmtQQ("Fld[?] cannot be found in any Ele-Lines"))
End Function

Private Function Msg_LinTyEr$(A As Lnx, Ty$)
Msg_LinTyEr = WMsg(A, FmtQQ(CM_LinTyEr, Ty, FmtDrs(ShtTyDrs)))
End Function

Private Function ErD(Tny$(), T() As Lnx, D() As Lnx) As String()
ErD = AyAddAp( _
    ErD_LinEr(D))
    'ErD_FldEr(D, T))
    '    ErDT_InvalidFld(D, Tny), _

End Function

Private Function ErE(E() As Lnx) As String()
Dim Eny$()
ErE = AyAdd(ErE_ELnxAy(E), ErE_DupE(E, Eny))
End Function

Private Function ErF(AllEny$(), F() As Lnx) As String()
ErF = AyAdd(ErF_1_LinEr(F), ErF_EleHasNoDef(F, AllEny))
End Function

Private Function ErT(Tny$(), T() As Lnx, E() As Lnx) As String()
ErT = AyAddAp( _
ErT_1_LinEr(T), _
ErT_NoTLin(T), _
ErT_FldHasNoEle(T, E), _
ErT_DupTbl(T, Tny))
End Function

Private Sub Z_ErT_1_OneLinEr()
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
Dim TLnx As New Lnx
Cas0:
    Set TLnx = Lnx(999, "Tbl 1")
    Ept = Sy("--- #1000[Tbl 1] FldNm[1] is not a name")
    GoTo Tst
Cas1:
    Set TLnx = Lnx(999, "A")
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
        .Fny = SySsl("A B D C")
    End With
    GoTo Tst
Cas5:
    TLnx.Lin = "A | * B | D C"
    Ept = Sy("")
    With Ept
        .T = "A"
        .Fny = SySsl("A B D C")
        .Sk = SySsl("B")
    End With
    GoTo Tst
Cas6:
    TLnx.Lin = "A |"
    Ept = Sy("")
    Push EptEr, "should have fields after |"
    GoTo Tst
Tst:
    Act = ErT_1_OneLinEr(TLnx)
    C
    Return
End Sub

Private Sub Z()
Z_ErT_1_OneLinEr
Exit Sub
'AAAA
'SchmLyEr
End Sub

