Attribute VB_Name = "MDao_Schm_Er"
Option Explicit
Const MD_NTermShouldBe3OrMore$ = "Should have 3 or more terms"

Private Function ErD_DLnxAy(A() As Lnx) As String()
Dim I
For Each I In Itr(A)
    If RmvTT(CvLnx(I).Lin) = "" Then PushI ErD_DLnxAy, MsgD_NTermShouldBe3OrMore(CvLnx(I))
Next
End Function

Private Function ErD_F_FnyNotIn_Fny(D() As Lnx, T$, Fny$()) As String()
'Fny$ is the fields in T$.
'D-Lnx.Lin is Des $T $F $D.  For the T$=$T line, the $F not in Fny$(), it is error
Dim I
For Each I In Itr(D)
    PushNonBlankStr ErD_F_FnyNotIn_Fny, ErD_F_FnyNotIn_Fny1(CvLnx(I), T, Fny)
Next
End Function

Private Function ErD_F_FnyNotIn_Fny1$(D As Lnx, T$, Fny$())
Dim Tbl$, Fld$, Des$, X$
Asg3TRst D.Lin, X, Tbl, Fld, Des
If X <> "Des" Then Stop
If Tbl <> T Then Exit Function
If Not HasEle(Fny, Fld) Then
    ErD_F_FnyNotIn_Fny1 = MsgD_FldNotIn_Fny()
End If
End Function

Private Function ErD_FldEr(A() As Lnx, T() As Lnx) As String() _
'Given A is D-Lnx having fmt = $Tbl $Fld $D, _
'This Sub checks if $Fld is in Fny
Dim J%, Fld$, Tbl$, Tny$(), FnyAy$()

For J = 0 To UB(A)
    AsgTT A(J).Lin, Tbl, Fld
    If Tbl <> "." Then
        PushIAy ErD_FldEr, ErD_FldEr1(A(J), Tny, FnyAy)
    End If
Next
End Function

Private Function ErD_FldEr1(A As Lnx, Tny$(), FnyAy$()) As String()

End Function

Private Function ErD_T_Tbl_NotIn_Tny1$(D As Lnx, Tny$())

End Function

Private Function ErD_TblEr(D() As Lnx, Tny$()) As String()
'D-Lnx.Lin is Des $T $F $D. If $T<>"." and not in Tny, it is error
Dim J%
For J = 0 To UB(D)
    PushNonBlankStr ErD_TblEr, ErD_T_Tbl_NotIn_Tny1(D(J), Tny)
Next
End Function

Private Function ErE_DupE(E() As Lnx, Eny$()) As String()
Dim Ele
For Each Ele In Itr(AywDup(Eny))
    Push ErE_DupE, MsgE_DupE(FndELnoAy(E, Ele), Ele)
Next
End Function

Private Function ErE_ELnx(A As Lnx) As String()
Dim L$, Ty$
Dim E$  ' EleNm of the Ele-Lin
L = A.Lin
'    AsgApAy ShfVal(L, "*Nm *Ty ?Req ?ZLen TxtSz VTxt Dft VRul Expr"), _
                     .E, Ty, .Req, .ZLen, .TxtSz, .VTxt, .Dft, .VRul, .Expr
    '.Ty = DaoTy(Ty)
    If L <> "" Then
        Dim ExcessEle$
        Push ErE_ELnx, MsgE_ExcessEleItm(A, ExcessEle)
    End If
'    If .Ty = 0 Then
'        Push OEr, MsgTyEr(A.Ix, Ty)
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

Private Function ErF_FLnx(F As Lnx) As String()
Dim LikFF$, A$, V$
'    AsgApAy Sy43TRst(F.Lin), .E, .LikT, V, A
'    .LikFny = SySsl(LikFF)
End Function

Private Function ErF_FLnxAy(A() As Lnx) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy ErF_FLnxAy, ErF_FLnx(A(J))
Next
End Function

Private Function ErT_DupT(T() As Lnx, Tny$()) As String()
Dim Tbl
For Each Tbl In Itr(AywDup(Tny))
    Push ErT_DupT, MsgT_DupT(FndTLnoAy(T, Tbl), Tbl)
Next
End Function

Private Function ErT_NoTLin(A() As Lnx) As String()
If Sz(A) > 0 Then Exit Function
PushI ErT_NoTLin, MsgT_NoTLin
End Function

Private Function ErT_TLnx(T As Lnx) As String()
Dim L$
Dim Tbl$
    L = T.Lin
    Tbl = ShfTerm(L)
    L = Replace(L, "*", Tbl)
'1
Select Case SubStrCnt(L, "|")
Case 0, 2
Case Else: PushI ErT_TLnx, MsgT_VBar_Cnt(T): Exit Function
End Select

'2
If Not IsNm(Tbl) Then
    PushI ErT_TLnx, MsgT_TblIsNotNm(T)
    Exit Function
End If
'
Dim Fny$()
    Fny = SySsl(Replace(L, "|", " "))
    
If HasSubStr(L, "|") Then
'3
    Dim IdFld$
    IdFld = Trim(TakBef(L, "|"))
    If IdFld <> Tbl & "Id" Then
        PushI ErT_TLnx, MsgT_IdFld(T)
        Exit Function
    End If
'4
    If Trim(TakBet(L, "|", "|")) = "" Then
        PushI ErT_TLnx, MsgT_NoFLdBetVV(T)
        Exit Function
    End If
End If
'5
    Dim Dup$()
    Dup = AywDup(Fny)
    If Sz(Dup) > 0 Then
        PushI ErT_TLnx, MsgT_DupF(T, Tbl, Dup)
        Exit Function
    End If
'6
If Sz(Fny) = 0 Then
    PushI ErT_TLnx, MsgT_NoFld(T)
    Exit Function
End If
'7
Dim F
For Each F In Itr(Fny)
    If Not IsNm(F) Then
        PushI ErT_TLnx, MsgT_FldIsNotANmEr(T, F)
    End If
Next
End Function

Private Function ErT_TLnxAy(A() As Lnx) As String()
Dim I
For Each I In Itr(A)
    PushIAy ErT_TLnxAy, ErT_TLnx(CvLnx(I))
Next
End Function

Private Function FndELnoAy(E() As Lnx, Ele) As Long()
Dim J%
For J = 0 To UBound(E)
    If T1(E(J).Lin) = Ele Then
        PushI FndELnoAy, E(J).Ix + 1
    End If
Next
End Function

Private Function FndTLnoAy(A() As Lnx, T) As Long()
Dim J%
For J = 0 To UB(A)
    If T1(A(J).Lin) = T Then
        PushI FndTLnoAy, A(J).Ix + 1
    End If
Next
End Function

Private Function Msg_LnoAyMsg(LnoAy&(), M$)
Msg_LnoAyMsg = FmtQQ("--- #[?] ?", JnSpc(LnoAy), M)
End Function

Private Function MsgLnxMsg$(A As Lnx, M$)
MsgLnxMsg = FmtQQ("--- #?[?] ?", A.Ix + 1, A.Lin, M)
End Function

Private Property Get MsgD_FldNotIn_Fny$()

End Property

Private Function MsgD_NTermShouldBe3OrMore$(D As Lnx)
MsgD_NTermShouldBe3OrMore = MsgLnxMsg(D, MD_NTermShouldBe3OrMore)
End Function

Private Function MsgD_T_Tbl_NotIn_Tny$(A As Lnx, T, Tblss$)
MsgD_T_Tbl_NotIn_Tny = MsgLnxMsg(A, FmtQQ("T[?] is invalid.  Valid T[?]", T, Tblss))
End Function

Private Function MsgE_DupE$(LnoAy&(), E)
MsgE_DupE = Msg_LnoAyMsg(LnoAy, FmtQQ("This E[?] is dup", E))
End Function

Private Function MsgT_DupT$(LnoAy&(), T)
MsgT_DupT = Msg_LnoAyMsg(LnoAy, FmtQQ("This Tbl[?] is dup", T))
End Function

Private Function MsgE_ExcessEleItm$(A As Lnx, ExcessEle$)
MsgE_ExcessEleItm = MsgLnxMsg(A, FmtQQ("Excess Ele Item [?]", ExcessEle))
End Function

Private Function MsgF_ExcessTxtSz$(A As Lnx)
MsgF_ExcessTxtSz = MsgLnxMsg(A, "Non-Txt-Ty should not have TxtSz")
End Function

Private Function MsgF_Ele_NotIn_Eny$(A As Lnx, E$, Eless$)
MsgF_Ele_NotIn_Eny = MsgLnxMsg(A, FmtQQ("Ele of is not in F-Lin not in Eny", E, Eless))
End Function

Private Function MsgE_FldEleEr$(A As Lnx, E$, Eless$)
MsgE_FldEleEr = MsgLnxMsg(A, FmtQQ("E[?] is invalid.  Valid E is [?]", E, Eless))
End Function

Private Function MsgFzDLy_NotIn_Fny$(A As Lnx, T$, F$, Fssl$)
MsgFzDLy_NotIn_Fny = MsgLnxMsg(A, FmtQQ("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function MsgT_DupF$(A As Lnx, T$, Fny$())
MsgT_DupF = MsgLnxMsg(A, FmtQQ("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function

Private Function MsgT_FldIsNotANmEr$(A As Lnx, F)
MsgT_FldIsNotANmEr = MsgLnxMsg(A, FmtQQ("FldNm[?] is not a name", F))
End Function

Private Function MsgT_IdFld$(A As Lnx)
Const M$ = "The field before first | must be *Id field"
MsgT_IdFld = MsgLnxMsg(A, M)
End Function

Private Function MsgT_NoFld(A As Lnx)
MsgT_NoFld = MsgLnxMsg(A, "No field")
End Function

Private Function MsgT_NoFLdBetVV$(A As Lnx)
MsgT_NoFLdBetVV = MsgLnxMsg(A, "No field between | |")
End Function

Private Property Get MsgT_NoTLin$()
MsgT_NoTLin = "No T-Line"
End Property

Private Function MsgT_TblIsNotNm$(A As Lnx)
MsgT_TblIsNotNm = MsgLnxMsg(A, "Tbl is not a name")
End Function

Private Function MsgT_VBar_Cnt$(A As Lnx)
Const M$ = "The T-Lin should have 0 or 2 VBar only"
MsgT_VBar_Cnt = MsgLnxMsg(A, M)
End Function

Private Function MsgT_FldEr$(A As Lnx, F$)
MsgT_FldEr = MsgLnxMsg(A, FmtQQ("Fld[?] cannot be found in any Ele-Lines"))
End Function

Private Function MsgTyEr$(A As Lnx, Ty$)
MsgTyEr = MsgLnxMsg(A, FmtQQ("Invalid DaoTy[?].  Valid Ty[?]", Ty, ShtTyzDaosl))
End Function

Property Get SampSchm() As String()
SampSchm = SplitCrLf(SampSchmLines)
End Property

Property Get SampSchmLines$()
Const A_1$ = "Tbl $Lo_FmtWs  Wsn |" & _
vbCrLf & "Tbl $Lo_FmtWdt     Wsn Seq Wdt ColNmss Er" & _
vbCrLf & "Tbl $Lo_FmtLvl | Wsn" & _
vbCrLf & "Tbl $Lnk" & _
vbCrLf & "Fld " & _
vbCrLf & "Fld $"

SampSchmLines = A_1
End Property

Function ErSchm(Schm$()) As String()
'========================================
Dim E() As Lnx
Dim F() As Lnx
Dim D() As Lnx
Dim T() As Lnx
    Dim X()  As Lnx
    X = ClnLnxAy(Schm)
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
    AllFny = FnyzTdStrAy(LyzLnxAy(T))
    AllEny = AyTakT1(LyzLnxAy(E))
Dim Er$(), mT$(), mF$(), mEle$(), mDes$()
    Er = LnxAyT1Chk(X, "Des Ele Fld Tbl")
    mT = ErT(Tny, T, E)
    mF = ErF(AllFny, F)
    'mEle = ErE
    'mDes = ErD
ErSchm = AyAddAp(Er, mT, mF, mEle, mDes)
End Function

Function ErD_(Tny$(), T() As Lnx, D() As Lnx) As String()
ErD_ = AyAddAp( _
    ErD_DLnxAy(D), _
    ErD_TblEr(D, Tny), _
    ErD_FldEr(D, T))
End Function

Function ErE(Eny$(), E() As Lnx) As String()
ErE = AyAdd(ErE_ELnxAy(E), ErE_DupE(E, Eny))
End Function

Function ErF(AllEny$(), F() As Lnx) As String()
ErF = AyAdd(ErF_FLnxAy(F), ErF_EleHasNoDef(F, AllEny))
End Function

Function ErT(Tny$(), T() As Lnx, E() As Lnx) As String()
ErT = AyAddAp( _
ErT_TLnxAy(T), _
ErT_NoTLin(T), _
ErT_FldHasNoEle(T, E), _
ErT_DupT(T, Tny))
End Function

Private Sub Z_ErT_TLnx()
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
    Act = ErT_TLnx(TLnx)
    C
    Return
End Sub

Private Sub Z()
Z_ErT_TLnx
Exit Sub
'AAAA
'SchmLyEr
End Sub
