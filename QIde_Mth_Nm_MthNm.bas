Attribute VB_Name = "QIde_Mth_Nm_MthNm"
Option Explicit
Private Const CMod$ = "MIde_Mth_Nm_Get."
Private Const Asm$ = "QIde"
Public Const DoczDta_MthQVNm$ = "It is a String dervied from Nm.  Q for quoted.  V for verb.  It has 3 Patn: NoVerb-[#xxx], MidVerb-[xxx(vvv)xxx], FstVerb-[(vvv)xxx]."
Public Const DoczNmRul_FstVerbBeingDo1$ = "The Fun will not return any value"
Public Const DoczNmRul_FstVerbBeingDo2$ = "The Cmls aft Do is a verb"
Sub AsgDNm(DNm$, O1$, O2$, O3$)
Dim Ay$(): Ay = Split(DNm, ".")
Select Case Si(Ay)
Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
Case Else: Stop
End Select
End Sub

Property Get CQMdn$()
CQMdn = QMdnzM(CMd)
End Property
Function CQMthn$()
On Error GoTo X
CQMthn = QMthn(CMd, CMthLin)
Exit Function
X: Debug.Print CSub
End Function

Function QMthn$(A As CodeModule, Lin)
Dim D$: D = MthDnzLin(Lin): If D = "" Then Exit Function
QMthn = QMdnzM(A) & "." & D
End Function

Function PMthny(Src$()) As String()
Dim Ix, N$, B As Mthn3
For Each Ix In MthIxItr(Src)
    B = Mthn3zL(Src(Ix))
    If B.Nm <> "" Then
        If B.ShtMdy = "" Or B.ShtMdy = "Pub" Then
            PushI PMthny, B.Nm
        End If
    End If
Next
End Function

Function MthnyzMthLiny(MthLiny$()) As String()
Const CSub$ = CMod & "MthnyzMthLiny"
Dim I, Nm$, J%, MthLin
For Each I In Itr(MthLiny)
    Nm = MthnzLin(I)
    If Nm = "" Then Thw CSub, "Given MthLiny does not have Mthn", "[MthLin with error] Ix MthLiny", I, J, AddIxPfx(MthLiny)
    PushI MthnyzMthLiny, Nm
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

Function IsPMth(Lin) As Boolean 'Pum = PubMthn
Dim L$: L = Lin
If ShfShtMdy(L) <> "Pub" Then Exit Function
If MthTy(L) = "" Then Exit Function
IsPMth = True
End Function

Function Mthn$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfMthTy(L) = "" Then Exit Function
Mthn = Nm(L)
End Function

Function MthnzMthDn$(MthDn$)
If MthDn = "*Dcl" Then MthnzMthDn = MthDn: Exit Function
Dim A$()
A = SplitDot(MthDn)
If Si(A) <> 3 Then Thw CSub, "MthDn should have 2 dot", "MthDn", MthDn
MthnzMthDn = A(0)
End Function
Private Sub ZZ_MthDnzLin()
Debug.Print MthDnzLin("Function MthnzMthDn$(MthDn$)")
Dim Lin$
End Sub
Function MthDnzLin$(Lin)
Stop
MthDnzLin = MthDnzMthn3(Mthn3zL(Lin))
End Function
Function MthSQNyInVbe() As String()
MthSQNyInVbe = MthSQNyzV(CVbe)
End Function
Function MthSQNyzV(A As Vbe) As String()
Dim QMthn
For Each QMthn In Itr(QMthnyzV(A))
    PushI MthSQNyzV, MthSQNm(CStr(QMthn))
Next
End Function

Function MthSQNm$(QMthn$)
Dim A$(): A = SplitDot(QMthn): If Si(A) <> 5 Then Thw CSub, "QMthn should have 4 dots", "QMthn", QMthn
Dim P$, Md$, M$, T$, N$
AsgAp A, P, Md, M, T, N
MthSQNm = JnDotAp(N, MthMdyc(M) & MthTyc(T), P, Md)
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
Function MthDn$(Lin)
MthDn = MthDnzN3(Mthn3zL(Lin))
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
GoTo ZZ
Dim A$
A = "Function Mthn(A)": Ept = "Mthn.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = Mthn(A)
    C
    Return
ZZ:
    Dim O$(), L
    For Each L In SrczV(CVbe)
        PushNonBlank O, Mthn(CStr(L))
    Next
    Brw O
End Sub

Function MthMdy$(Lin)
MthMdy = FstEleEv(MthMdyAy, T1(Lin))
End Function

Function MthKd$(Lin)
MthKd = TakMthKd(RmvMdy(Lin))
End Function

Function Rpl$(S, SubStr$, By$, Optional Ith% = 1)
Dim P&: P = InStrWiIthSubStr(S, SubStr, Ith)
If P = 0 Then Rpl = S: Exit Function
Rpl = Replace(S, SubStr, By, P, 1)
End Function
Function PoszSubStr(S, SubStr) As Pos
InStr
End Function
Property Get Rel0Mthn2Mdn() As Rel
Dim O As New Rel
End Property

Function ModNyzPum(PMthn) As String()
ModNyzPum = ModNyzPPm(CPj, PMthn)
End Function
Function PMthnyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsPMth(L) Then PushI PMthnyzS, Mthn(L)
Next
End Function

Private Sub ZZ_ModNyzPPm()
Dim P As VBProject, PMthn
GoSub ZZ
Exit Sub
ZZ:
    D ModNyzPPm(CPj, "AA")
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
Function ModNyzPPm(P As VBProject, PMthn) As String()
Dim I, Md As CodeModule
For Each I In ModItrzP(P)
    Set Md = I
    Debug.Print Mdn(Md)
    Stop
    If Mdn(Md) = "QId_Mth_Nm" Then Stop
    If HasPMth(Src(Md), PMthn) Then PushI ModNyzPPm, Mdn(Md)
Next
End Function

Function MthTy$(Lin)
MthTy = PfxzPfxSyPlusSpc(RmvMthMdy(Lin), MthTyAy)
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
Private Sub Z_Dry__Mthn_Verb_InVbe()
BrwDry Dry__Mthn_Verb_InVbe
End Sub
Function Dry__Mthn_Verb_InVbe() As Variant()
Dim Mthn, I, ODry()
For Each I In Itr(MthnyV)
    Mthn = I
    PushI ODry, Sy(Mthn, Verb(Mthn))
Next
Dry__Mthn_Verb_InVbe = DrywDist(ODry)
End Function
Private Sub Z_MthnsetVWoVerb()
MthnsetVWoVerb.Srt.Vc
End Sub

Property Get MthnyVWiVerb() As String()
Dim Mthn, I, J&
For Each I In Itr(MthnyV)
    Mthn = I
'    If HasSubStr(Mthn, "Z_ExprDic") Then Stop
    If J Mod 100 = 0 Then Debug.Print J
    If HasVerb(Mthn) Then PushI MthnyVWiVerb, Mthn
    J = J + 1
Next
End Property
Property Get MthnyVWoVerb() As String()
Dim Mthn, I
For Each I In Itr(MthnyV)
    Mthn = I
    If Not HasVerb(Mthn) Then PushI MthnyVWiVerb, Mthn
Next
End Property

Function HasVerb(Nm) As Boolean
HasVerb = Verb(Nm) <> ""
End Function

Property Get MthnsetVWiVerb() As Aset
Set MthnsetVWiVerb = AsetzAy(MthnyVWiVerb)
End Property

Property Get MthnsetVWoVerb() As Aset
Set MthnsetVWoVerb = AsetzAy(MthnyVWoVerb)
End Property

Function MthnsetV() As Aset
Set MthnsetV = AsetzAy(MthnyV)
End Function

Function MthnyzSI(Src$(), MthIxy&()) As String()
Dim Ix
For Each Ix In Itr(MthIxy)
    PushI MthnyzSI, Mthn3zL(Src(Ix)).Nm
Next
End Function

Function MthnyV() As String()
MthnyV = MthnyzV(CVbe)
End Function

Function MthnsetP() As Aset
Set MthnsetP = AsetzAy(MthnyP)
End Function

Function MthnyP() As String()
MthnyP = MthnyzP(CPj)
End Function
Function PMthnyzV(A As Vbe) As String()

End Function
Function PMthnyV() As String()
PMthnyV = PMthnyzV(CVbe)
End Function

Function MthnyzP(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy MthnyzP, MthnyzM(C.CodeModule)
Next
End Function

Function QMthnyV() As String()
QMthnyV = QMthnyzV(CVbe)
End Function

Function QMthnWsInVbe() As Worksheet
Set QMthnWsInVbe = QMthnWszV(CVbe)
End Function

Function QMthnWszV(A As Vbe) As Worksheet
Dim Dry(): Dry = DryzDotLy(QMthnyzV(A))
Set QMthnWszV = WszDrs(DrszFF("Pj Md Mth Ty Mdy", Dry))
End Function

Function QMthnyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushAy QMthnyzV, QMthnyzP(P)
Next
End Function

Function QMthnyzM(A As CodeModule) As String()
QMthnyzM = AddPfxzAy(MthDNyzS(Src(A)), QMdnzM(A) & ".")
End Function

Function QDry_MthnzP(P As VBProject) As Variant()
Dim QNm
For Each QNm In Itr(QMthnyzP(P))
    PushI QDry_MthnzP, QDr_Mthn(CStr(QNm))
Next
End Function

Function QDr_Mthn(QMthn$) As String()
Dim O$(): O = SplitDot(QMthn)
If Si(O) <> 5 Then Thw CSub, "QMthn should have 4 dot", "QMthn", QMthn
QDr_Mthn = O
End Function

Function QMthnyzP(P As VBProject) As String()
Dim I
For Each I In MdItr(P)
    PushAy QMthnyzP, QMthnyzM(CvMd(I))
Next
End Function

Function PMthDNyzV(A As Vbe) As String()
PMthDNyzV = PMthDNyzV(A)
End Function

Function MthnyzFb(Fb) As String()
MthnyzFb = MthnyzV(VbezPjf(Fb))
ClsPjf Fb
End Function


Private Sub Z_MthnyzFb()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
    For Each Fb In AppFbAy
        PushAy O, MthnyzFb(CStr(Fb))
    Next
    Brw O
    Return
X_BrwOne:
    Dim A$(): A = AppFbAy
    Brw MthnyzFb(A(0))
    Return
End Sub

Function MthnyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlank MthnyzS, Mthn(L)
Next
End Function

Function PMthnyzM(A As CodeModule) As String()
PMthnyzM = PMthnyzS(Src(A))
End Function

Private Sub ZZ()
Z_MthnyzFb
MIde_Mth_Nm:
End Sub

Function MthnyzM(A As CodeModule) As String()
MthnyzM = MthnyzS(Src(A))
End Function

Private Sub ZZ_MthnzS()
GoSub ZZ
Exit Sub
ZZ:
   B MthnyzS(SrczP(CPj))
   Return
End Sub

Function SqzMthDNyzP(P As VBProject) As Variant()
SqzMthDNyzP = SqzMthDNy(MthnyzP(P))
End Function

Function MthDnWszP(P As VBProject) As Worksheet
Set MthDnWszP = ShwWs(WszSq(SqzMthDNyzP(P)))
End Function

Function MthnyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MthnyzV, MthnyzP(P)
Next
End Function

Function MthAsetVbe() As Aset
Set MthAsetVbe = AsetzAy(MthnyV)
End Function

Property Get MthnyzCMd() As String()
MthnyzCMd = MthnyzM(CMd)
End Property

Private Sub Z_MthDNy()
Dim Md1 As CodeModule
Set Md1 = Md("AAAMod")
BrwAy MthnyzM(Md1)
BrwAy MthDNyzM(Md1)
End Sub

Private Sub Z_MthDNyzS()
BrwAy MthDNyzS(CSrc)
End Sub

Function MthDNyV() As String()
MthDNyV = MthDNyzV(CVbe)
End Function

Function MthDNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MthDNyzV, MthDNyzP(P)
Next
End Function

Function MthDNyzM(A As CodeModule) As String()
MthDNyzM = MthDNyzS(Src(A))
End Function
Function MthDNyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlank MthDNyzS, MthDn(L)
Next
End Function

Function HasMth(Src$(), Mthn) As Boolean
HasMth = FstMthIx(Src, Mthn) >= 0
End Function

Function HasMthzM(A As CodeModule, Mthn) As Boolean
HasMthzM = HasMth(Src(A), Mthn)
End Function


Function Md_MthnDic(A As CodeModule) As Dictionary
'Set Md_MthnDic = Src_MthnDic(Src(A))
End Function

Private Sub Z_Src_MthnDic()
'BrwDic Src_MthnDic(CSrc)
End Sub

Function MthnCmlSetVbe() As Aset
Set MthnCmlSetVbe = CmlSetzNy(MthnyV)
End Function
Function Drs_MthnV() As Drs
Drs_MthnV = Drs_MthnzV(CVbe)
End Function

Function Drs_MthnP() As Drs
Drs_MthnP = Drs_MthnzP(CPj)
End Function

Function Drs_MthnM() As Drs
Drs_MthnM = Drs_MthnzM(CMd)
End Function

Private Function Drs_MthnzM(M As CodeModule) As Drs
Drs_MthnzM = Drs(Fny_Mthn, Dry_MthnzM(M))
End Function

Private Function Drs_MthnzV(A As Vbe) As Drs
Drs_MthnzV = Drs(Fny_Mthn, Dry_MthnzV(A))
End Function

Function Drs_MthnzP(P As VBProject) As Drs
Drs_MthnzP = Drs(Fny_Mthn, Dry_MthnzP(P))
End Function

Private Function Dry_MthnzM(M As CodeModule) As Variant()
Dry_MthnzM = DryAddColzC3(Dry_MthnzS(Src(M)), Mdn(M), ShtCmpTy(M.Parent.Type), PjnzM(M))
End Function

Private Function Dry_MthnzV(A As Vbe) As Variant()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy Dry_MthnzV, Dry_MthnzP(P)
Next
End Function

Private Function Dry_MthnzP(P As VBProject) As Variant()
Dim M
For Each M In MdItr(P)
    PushIAy Dry_MthnzP, Dry_MthnzM(CvMd(M))
Next
End Function

Private Function Dry_MthnzS(Src$()) As Variant()
Dim MthLin
For Each MthLin In Itr(MthLinyzS(Src))
    PushISomSi Dry_MthnzS, Dr_Mthn(MthLin)
Next
End Function

Function Dr_Mthn(MthLin) As String()

End Function
