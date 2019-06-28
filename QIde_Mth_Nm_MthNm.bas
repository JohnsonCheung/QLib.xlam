Attribute VB_Name = "QIde_Mth_Nm_MthNm"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Nm_Get."
Private Const Asm$ = "QIde"
':Dta_MthQVNm$ = "It is a String dervied from Nm.  Q for quoted.  V for verb.  It has 3 Patn: NoVerb-[#xxx], MidVerb-[xxx(vvv)xxx], FstVerb-[(vvv)xxx]."
':Mthn: ! Rule1-FstVerbBeingDo: the mthn will not return any value
'       ! Rule2-FstVerbBeingDo: tThe Cmls aft Do is a verb
':Mthn: ! is :Nm less that 64 chr.  The rule for a Mthn is:
'       ! If there is a Subj in pm, put the Subj as fst CmlTerm and return that Subj;
'       ! give a Noun to the subj noun is MulCml.
'       ! Each Mthn must belong to one of these rule:
'       !   Noun | Noun.Verb.Extra | Verb.Variant | Noun.z.Variant
'       ! Pm-Rule
'       !   Subj    : Choose a subj in pm if there is more than one arg"
'       !   MuliNoun: It is Ok to group mul-arg as one subj
'       !   MulNounUseOneCml: Mul-noun as one subj use one Cml
':Noun: ! it is 1 or more Cml to form a Noun."
':Cml:     ! Tag:Type. P1.NumIsLCase:.  P2.LowDashIsLCase:.  P3.FstChrCanAnyNmChr:.
':Sfxss: Tag:NmRul. NmRul means variable or function name.
':VdtVerss:  P1.Opt: Each module may one DoczVdtVerbss.  P2.OneOccurance: "
':NounVerbExtra: "Tag: FunNmRule.  Prp1.TakAndRetNoun: Fst Cml is Noun and Return Noun.  Prp2.OneCmlNoun: Noun should be 1 Cml.  " & _
'                ! Prp3.VdtVerb: Snd Cml should be approved/valid noun.  Prp4.OptExtra: Extra is optional."
Sub AsgDNm(DNm$, O1$, O2$, O3$)
Dim Ay$(): Ay = Split(DNm, ".")
Select Case Si(Ay)
Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
Case Else: Stop
End Select
End Sub


Function QMthn$(M As CodeModule, Lin)
Dim D$: D = MthDnzLin(Lin): If D = "" Then Exit Function
QMthn = MdDn(M) & "." & D
End Function

Function PMthNy(Src$()) As String()
Dim Ix, N$, B As Mthn3
For Each Ix In MthIxItr(Src)
    B = Mthn3zL(Src(Ix))
    If B.Nm <> "" Then
        If B.ShtMdy = "" Or B.ShtMdy = "Pub" Then
            PushI PMthNy, B.Nm
        End If
    End If
Next
End Function

Function MthNyzMthLinAy(MthLinAy$()) As String()
Const CSub$ = CMod & "MthNyzMthLinAy"
Dim I, Nm$, J%, MthLin
For Each I In Itr(MthLinAy)
    Nm = MthnzLin(I)
    If Nm = "" Then Thw CSub, "Given MthLinAy does not have Mthn", "[MthLin with error] Ix MthLinAy", I, J, AddIxPfx(MthLinAy)
    PushI MthNyzMthLinAy, Nm
    J = J + 1
Next
End Function
Function Ens2Dot(S) As StrOpt
Select Case DotCnt(S)
Case 0: Ens2Dot = SomStr(".." & S)
Case 1: Ens2Dot = SomStr("." & S)
Case 2: Ens2Dot = SomStr(S)
End Select
End Function

Function RmvMthMdy$(Lin)
RmvMthMdy = RmvTerm(Lin, MthMdyAy)
End Function

Function Dimn$(Lin)
Dim L$: L = Lin
If ShfTerm(L, "Dim") Then Dimn = Nm(LTrim(L))
End Function
Function DimNy(Ly$()) As String()
Dim L
For Each L In Itr(Ly)
    PushI DimNy, Dimn(L)
Next
End Function
Function Mthn$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfMthTy(L) = "" Then Exit Function
Mthn = Nm(L)
End Function

Private Sub Z_MthDnzLin()
Debug.Print MthDnzLin("Function MthnzMthDn$(MthDn$)")
Dim Lin$
End Sub

Function MthDnzMthn3$(A As Mthn3)
MthDnzMthn3 = JnDotAp(A.Nm, A.ShtMdy, A.ShtTy)
End Function

Function MthDnzLin$(Lin)
MthDnzLin = MthDnzMthn3(Mthn3zL(Lin))
End Function

Function MthTyc$(ShtMthTy$)
Select Case ShtMthTy
Case "Fun": MthTyc = "F"
Case "Sub": MthTyc = "S"
Case "Get": MthTyc = "G"
Case "Let": MthTyc = "L"
Case "Set": MthTyc = "T"
Case Else: Thw CSub, "Invalid ShtMthTy.", "ShtMthTy VdtShtMthTy", ShtMthTy, ShtMthTyAy
End Select
End Function
Function MthMdyc$(ShtMthMdy$)
Select Case ShtMthMdy
Case "Pub": MthMdyc = "P"
Case "Prv": MthMdyc = "V"
Case "Frd": MthMdyc = "F"
Case Else: Thw CSub, "Invalid ShtMthMdy.", "ShtMthMdy VdtShtMthMdy", ShtMthMdy, ShtMthMdyAy
End Select
End Function

Function MthDn$(L)
MthDn = MthDnzMthn3(Mthn3zL(L))
End Function

Function MthnzLin(Lin)
MthnzLin = Mthn(Lin)
End Function

Function PrpNm$(Lin)
Dim L$
L = RmvMdy(Lin)
If ShfKd(L) <> "Property" Then Exit Function
PrpNm = Nm(L)
End Function

Function MthnzDNm$(Mthn)
Dim Ay$(): Ay = Split(Mthn, ".")
Dim Nm$
Select Case Si(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
MthnzDNm = Nm
End Function
Private Sub Z_Mthn()
GoTo Z
Dim A$
A = "Function Mthn(A)": Ept = "Mthn.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = Mthn(A)
    C
    Return
Z:
    Dim O$(), L
    For Each L In SrczV(CVbe)
        PushNB O, Mthn(CStr(L))
    Next
    Brw O
End Sub

Function MthMdy$(Lin)
MthMdy = IfIn(T1(Lin), MthMdyAy)
End Function

Function MthKd$(Lin)
MthKd = TakMthKd(RmvMdy(Lin))
End Function

Function Rpl$(S, SubStr$, By$, Optional Ith% = 1)
Dim P&: P = InStrWiIthSubStr(S, SubStr, Ith)
If P = 0 Then Rpl = S: Exit Function
Rpl = Replace(S, SubStr, By, P, 1)
End Function

Property Get Rel0Mthn2Mdn() As Rel
Dim O As New Rel
End Property

Function ModNyzPMth(PMthn) As String()
ModNyzPMth = ModNyzPjPMth(CPj, PMthn)
End Function
Function PMthNyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsLinPubMth(L) Then PushI PMthNyzS, Mthn(L)
Next
End Function

Private Sub Z_ModNyzPjPMth()
Dim P As VBProject, PMthn
GoSub Z
Exit Sub
Z:
    D ModNyzPjPMth(CPj, "AA")
    Stop
    Return
End Sub
Function HasPMth(Src$(), PMthn) As Boolean
Dim L
For Each L In Itr(Src)
    With Mthn3zL(L)
        If .ShtMdy = "Pub" Then
            If .Nm = PMthn Then
                HasPMth = True
                Exit Function
            End If
        End If
    End With
Next
End Function
Function ModNyzPjPMth(P As VBProject, PMthn) As String()
#If True Then
Dim Src As Drs: Src = DoMthP
Dim Sel As Drs: Sel = Dw2Eq(Src, "Mdy MdTy", "Pub", "Std")
#Else
Dim I, Md As CodeModule
For Each I In ModItrzP(P)
    Set Md = I
    If HasPMth(Src(Md), PMthn) Then PushI ModNyzPjPMth, Mdn(Md)
Next
#End If
End Function

Function MthTy$(Lin)
MthTy = PfxzAyS(RmvMthMdy(Lin), MthTyAy)
End Function

Private Sub Z_MthTy()
Dim O$(), L
For Each L In SrczMdn("Fct")
    Push O, MthTy(CStr(L)) & "." & L
Next
BrwAy O
End Sub

Private Sub Z_MthKd()
Dim A$
Ept = "Property": A = "Private Property Get": GoSub Tst
Ept = "Property": A = "Property Get":         GoSub Tst
Ept = "Property": A = " Property Get":        GoSub Tst
Ept = "Property": A = "Friend Property Get":  GoSub Tst
Ept = "Property": A = "Friend  Property Get": GoSub Tst
Ept = "":         A = "FriendProperty Get":   GoSub Tst
Exit Sub
Tst:
    Act = MthKd(A)
    C
    Return
End Sub


Private Sub Z_MthnsetVWiVerb()
MthnsetVWiVerb.Srt.Vc
End Sub
Private Sub Z_DyoMthnaVerbV()
BrwDy DyoMthnaVerbV
End Sub
Sub PushNDupDy(ODy(), Dr)
If HasDr(ODy, Dr) Then Exit Sub
PushI ODy, Dr
End Sub
Function DoMthnaVerbV() As Drs
DoMthnaVerbV = DrszFF("Mthn Verb", DyoMthnaVerbV)
End Function

Function DyoMthnaVerbV() As Variant()
Dim Mthn, O(): For Each Mthn In Itr(MthNyV)
    PushI O, Sy(Mthn, Verb(Mthn))
Next
DyoMthnaVerbV = O
End Function
Private Sub Z_MthnsetVWoVerb()
MthnsetVWoVerb.Srt.Vc
End Sub

Property Get MthNyVWiVerb() As String()
Dim Mthn, I, J&
For Each I In Itr(MthNyV)
    Mthn = I
'    If HasSubStr(Mthn, "Z_ExprDic") Then Stop
    If J Mod 100 = 0 Then Debug.Print J
    If HasVerb(Mthn) Then PushI MthNyVWiVerb, Mthn
    J = J + 1
Next
End Property
Property Get MthNyVWoVerb() As String()
Dim Mthn, I
For Each I In Itr(MthNyV)
    Mthn = I
    If Not HasVerb(Mthn) Then PushI MthNyVWiVerb, Mthn
Next
End Property

Function HasVerb(Nm) As Boolean
HasVerb = Verb(Nm) <> ""
End Function

Property Get MthnsetVWiVerb() As Aset
Set MthnsetVWiVerb = AsetzAy(MthNyVWiVerb)
End Property

Property Get MthnsetVWoVerb() As Aset
Set MthnsetVWoVerb = AsetzAy(MthNyVWoVerb)
End Property

Function MthnsetV() As Aset
Set MthnsetV = AsetzAy(MthNyV)
End Function

Function MthNyzSI(Src$(), MthIxy&()) As String()
Dim Ix
For Each Ix In Itr(MthIxy)
    PushI MthNyzSI, Mthn3zL(Src(Ix)).Nm
Next
End Function

Function MthNyV() As String()
MthNyV = MthNyzV(CVbe)
End Function

Function MthnsetP() As Aset
Set MthnsetP = AsetzAy(MthNyP)
End Function

Function MthNyP() As String()
MthNyP = MthNyzP(CPj)
End Function

Function MthNyzP(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy MthNyzP, MthNyzM(C.CodeModule)
Next
End Function

Function MthNyzFb(Fb) As String()
MthNyzFb = MthNyzV(VbezPjf(Fb))
ClsPjf Fb
End Function


Private Sub Z_MthNyzFb()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
    For Each Fb In AppFbAy
        PushAy O, MthNyzFb(CStr(Fb))
    Next
    Brw O
    Return
X_BrwOne:
    Dim A$(): A = AppFbAy
    Brw MthNyzFb(A(0))
    Return
End Sub

Function MthNyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushNB MthNyzS, Mthn(L)
Next
End Function

Function PMthNyzM(M As CodeModule) As String()
PMthNyzM = PMthNyzS(Src(M))
End Function

Private Sub Z()
Z_MthNyzFb
MIde_Mth_Nm:
End Sub

Function MthNyzM(M As CodeModule) As String()
MthNyzM = MthNyzS(Src(M))
End Function

Private Sub Z_MthnzS()
GoSub Z
Exit Sub
Z:
   B MthNyzS(SrczP(CPj))
   Return
End Sub

Function MthNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MthNyzV, MthNyzP(P)
Next
End Function

Function MthAsetVbe() As Aset
Set MthAsetVbe = AsetzAy(MthNyV)
End Function

Property Get MthNyzCMd() As String()
MthNyzCMd = MthNyzM(CMd)
End Property

Function HasMth(Src$(), Mthn) As Boolean
HasMth = FstMthIxzN(Src, Mthn) >= 0
End Function

Function HasMthzM(M As CodeModule, Mthn) As Boolean
HasMthzM = HasMth(Src(M), Mthn)
End Function

Function MthnCmlSetVbe() As Aset
Set MthnCmlSetVbe = CmlSetzNy(MthNyV)
End Function
Function DoMthnzV(V As Vbe) As Drs

End Function
Function DoMthnzM(M As CodeModule) As Drs
DoMthnzM = DoMthn(M)
End Function
Function DoMthnzP(P As VBProject) As Drs
Dim C As VBComponent
Dim Pn$: Pn = P.Name
For Each C In P.VBComponents
    Dim Mn$: Mn = C.Name
    Dim A As Drs: A = DoMthn(C.CodeModule)
    Dim B As Drs: B = InsColzDrsCC(A, "Pj Md", Pn, Mn)
    Dim O As Drs: O = AddDrs(O, A)
Next
DoMthnzP = O
End Function

Function DoMthnV() As Drs
DoMthnV = DoMthnzV(CVbe)
End Function

Function DoMthnP() As Drs
DoMthnP = DoMthnzP(CPj)
End Function

Function DoMthnM() As Drs
DoMthnM = DoMthnzM(CMd)
End Function

