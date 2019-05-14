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

Property Get MdQNm$()
MdQNm = MdQNmzMd(CMd)
End Property
Property Get CurMthQNm$()
On Error GoTo X
CurMthQNm = MthQNm(CMd, CurMthLin)
Exit Property
X: Debug.Print CSub
End Property

Function MthQNm$(A As CodeModule, Lin)
Dim D$: D = MthDnzLin(Lin): If D = "" Then Exit Function
MthQNm = MdQNmzMd(A) & "." & D
End Function

Function MthnyzPub(Src$()) As String()
Dim Ix, N$, B As Mthn3
For Each Ix In MthIxItr(Src)
    Set B = Mthn3(Src(Ix))
    If B.Nm <> "" Then
        If B.ShtMdy = "" Or B.ShtMdy = "Pub" Then
            PushI MthnyzPub, B.Nm
        End If
    End If
Next
End Function

Function MthnyzMthLiny(MthLiny$()) As String()
Const CSub$ = CMod & "MthnyzMthLiny"
Dim I, Nm$, J%, MthLin
For Each I In Itr(MthLiny)
    MthLin = I
    Nm = Mthn(MthLin)
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

Function Mthn(Lin, Optional B As WhMth) As Mthn3
If Not IsMthLin(Lin) Then Exit Function
Mthn = Mthn3(Lin, B).Nm
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
MthDnzLin = MthDnzMthn3(Mthn3zLin(Lin))
End Function
Function MthSQNyInVbe(Optional WhStr$) As String()
MthSQNyInVbe = MthSQNyzV(CVbe, WhStr)
End Function
Function MthSQNyzV(A As Vbe, Optional WhStr$) As String()
Dim MthQNm
For Each MthQNm In Itr(MthQNyzV(A, WhStr))
    PushI MthSQNyzV, MthSQNm(CStr(MthQNm))
Next
End Function

Function MthSQNm$(MthQNm$)
Dim A$(): A = SplitDot(MthQNm): If Si(A) <> 5 Then Thw CSub, "MthQNm should have 4 dots", "MthQNm", MthQNm
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
Function MthDn$(Lin, Optional B As WhMth)
MthDn = Mthn3(Lin, B).DNm
End Function

Function MthnzLin(Lin, Optional B As WhMth)
MthnzLin = Mthn3(Lin, B).Nm
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
    With Mthn3(L)
        If .IsPub Then
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
Private Sub Z_DryOf_Mthn_Verb_InVbe()
BrwDry DryOf_Mthn_Verb_InVbe
End Sub
Function DryOf_Mthn_Verb_InVbe() As Variant()
Dim Mthn, I, ODry()
For Each I In Itr(MthnyV)
    Mthn = I
    PushI ODry, Sy(Mthn, Verb(Mthn))
Next
DryOf_Mthn_Verb_InVbe = DrywDist(ODry)
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

Function MthnsetV(Optional WhStr$) As Aset
Set MthnsetV = AsetzAy(MthnyV(WhStr))
End Function

Function MthnyzSI(Src$(), MthIxy&()) As String()
Dim Ix
For Each Ix In Itr(MthIxy)
    PushI MthnyzSI, Mthn3zLin(Src(Ix)).Nm
Next
End Function

Function MthnyV(Optional WhStr$) As String()
MthnyV = MthNyzV(CVbe, WhStr$)
End Function

Function MthnsetP(Optional WhStr$) As Aset
Set MthnsetP = AsetzAy(MthnyP(WhStr))
End Function

Function MthnyP(Optional WhStr$) As String()
MthnyP = MthnyzP(CPj, WhStr$)
End Function

Function MthnyVzPub(Optional WhStr$) As String()
MthnyVzPub = MthNyzV(CVbe, WhStr & " -Pub")
End Function

Function MthnyzP(P As VBProject, Optional WhStr$) As String()
Dim M, W As WhMth
Set W = WhMthzStr(WhStr)
For Each M In MdItr(P, WhStr)
    PushIAy MthnyzP, MthnyzMd(CvMd(M), W)
Next
End Function

Function MthQNyV(Optional WhStr$) As String()
MthQNyV = MthQNyzV(CVbe, WhStr)
End Function

Function MthQNmWsInVbe(Optional WhStr$) As Worksheet
Set MthQNmWsInVbe = MthQNmWszV(CVbe, WhStr)
End Function

Function MthQNmWszV(A As Vbe, Optional WhStr$) As Worksheet
Dim Dry(): Dry = DryzDotLy(MthQNyzV(A, WhStr))
Set MthQNmWszV = WszDrs(DrszFF("Pj Md Mth Ty Mdy", Dry))
End Function

Function MthQNyzV(A As Vbe, Optional WhStr$) As String()
Dim I
For Each I In PjItr(A, WhStr)
    PushAy MthQNyzV, MthQNyzP(CvPj(I), WhStr)
Next
End Function

Function MthQNyzMd(A As CodeModule, Optional WhStr$) As String()
MthQNyzMd = AddPfxzAy(MthDNyzS(Src(A), WhStr), MdQNmzMd(A) & ".")
End Function

Function MthQNmDryzP(P As VBProject, Optional WhStr$) As Variant()
Dim QNm
For Each QNm In Itr(MthQNyzP(P, WhStr))
    PushI MthQNmDryzP, MthQNmDr(CStr(QNm))
Next
End Function

Function MthQNmDr(MthQNm$) As String()
Dim O$(): O = SplitDot(MthQNm)
If Si(O) <> 5 Then Thw CSub, "MthQNm should have 4 dot", "MthQNm", MthQNm
MthQNmDr = O
End Function

Function MthQNyzP(P As VBProject, Optional WhStr$) As String()
Dim I
For Each I In MdItr(P, WhStr)
    PushAy MthQNyzP, MthQNyzMd(CvMd(I), WhStr)
Next
End Function

Function MthDNyzVzPub(A As Vbe, Optional WhStr$) As String()
MthDNyzVzPub = MthDNyzV(A, WhStr & " -Pub")
End Function

Function MthDNyzMthn(A As Vbe, Mthn) As String()
Dim P As VBProject, M, Md As CodeModule
For Each P In A.VBProjects
    For Each M In P.VBComponents
        PushIAy MthDNyzMthn, MthDNyzMMthn(CvMd(M), Mthn)
    Next
Next
End Function
Function MthDNyzMMthn(Md As CodeModule, Mthn) As String()

End Function
Property Get MMthnyV() As String()
MthnyV
End Property
Function MthnyzFb(Fb) As String()
MthnyzFb = MthNyzV(VbezPjf(Fb))
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


Private Sub Z_MthnyzSrc()
Brw MthnyzSrc(CurSrc)
End Sub

Function MthnyzSrc(Src$(), Optional B As WhMth) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlank MthnyzSrc, Mthn(CStr(L), B)
Next
End Function

Function MthnyPubzMd(A As CodeModule, Optional WhStr$) As String()
MthnyPubzMd = MthnyzSrc(Src(A), WhMthzStr(WhStr))
End Function

Private Sub ZZ()
Z_MthnyzFb
Z_MthnyzSrc
MIde_Mth_Nm:
End Sub


Function MthnyzMd(A As CodeModule, Optional B As WhMth) As String()
MthnyzMd = MthnyzSrc(Src(A), B)
End Function



Private Sub ZZ_MthnyzSrc()
Dim Act$()
   Act = MthnyzSrc(CurSrc)
   BrwAy Act
End Sub

Function SqzMthDNyzP(P As VBProject) As Variant()
SqzMthDNyzP = SqzMthDNy(MthnyzP(P, True))
End Function

Function MthDnWszP(P As VBProject) As Worksheet
Set MthDnWszP = ShwWs(WszSq(SqzMthDNyzP(P)))
End Function

Function MthNyzV(A As Vbe, Optional WhStr$) As String()
Dim I
For Each I In PjItr(A, WhStr)
    PushIAy MthNyzV, MthnyzP(CvPj(I), WhStr)
Next
End Function

Function MthAsetVbe(Optional WhStr$) As Aset
Set MthAsetVbe = AsetzAy(MthnyV(WhStr))
End Function

Property Get MthnyzCMd() As String()
MthnyzCMd = MthnyzMd(CMd)
End Property

Private Sub Z_MthDNy()
Dim Md1 As CodeModule
Set Md1 = Md("AAAMod")
BrwAy MthnyzMd(Md1)
BrwAy MthDNyzM(Md1)
End Sub


Private Sub Z_MthDNyzS()
BrwAy MthDNyzS(CurSrc)
End Sub

Function MthDNyV(Optional WhStr$) As String()
MthDNyV = MthDNyzV(CVbe, WhStr)
End Function

Function MthDNyzV(A As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In PjItr(A, WhStr)
    PushIAy MthDNyzV, MthDNyzP(P, WhStr)
Next
End Function

Function MthDNyzM(A As CodeModule, Optional WhStr$) As String()
MthDNyzM = MthDNyzS(Src(A), WhMthzStr(WhStr))
End Function
Function MthDNyzS(Src$(), Optional WhStr$) As String()
Dim L, B As WhMth
Set B = WhMthzStr(WhStr)
For Each L In Itr(Src)
    PushNonBlank MthDNyzS, MthDn(CStr(L), B)
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
'BrwDic Src_MthnDic(CurSrc)
End Sub

Function MthnCmlSetVbe(Optional WhStr$) As Aset
Set MthnCmlSetVbe = CmlSetzNy(MthnyV(WhStr))
End Function
Function DrsOfMthnV(Optional WhStr$) As Drs
DrsOfMthnV = DrsOfMthnzV(CVbe, WhStr)
End Function

Function DrsOfMthnP(Optional WhStr$) As Drs
DrsOfMthnP = DrsOfMthnzP(CPj, WhStr)
End Function

Function DrsOfMthnM(Optional WhStr$) As Drs
DrsOfMthnM = DrsOfMthnzM(CMd, WhStr)
End Function

Private Function DrsOfMthnzM(M As CodeModule, Optional WhStr$) As Drs
DrsOfMthnzM = Drs(FnyOfMthn, MthnDryzMd(M, WhMthzStr(WhStr)))
End Function

Private Function DrsOfMthnzV(A As Vbe, Optional WhStr$) As Drs
DrsOfMthnzV = Drs(FnyOfMthn, MthnDryzV(A, WhStr))
End Function

Function DrsOfMthnzP(P As VBProject, Optional WhStr$) As Drs
DrsOfMthnzP = Drs(FnyOfMthn, MthnDryzP(P, WhStr))
End Function

Private Function MthnDryzMd(M As CodeModule, Optional B As WhMth) As Variant()
MthnDryzMd = DryAddColzC3(MthnDryzSrc(Src(M), B), Mdn(M), ShtCmpTy(M.Parent.Type), PjnzM(M))
End Function

Private Function MthnDryzV(A As Vbe, Optional WhStr$) As Variant()
Dim P
For Each P In PjItr(A, WhStr)
    PushIAy MthnDryzV, MthnDryzP(CvPj(P), WhStr)
Next
End Function

Private Function MthnDryzP(P As VBProject, Optional WhStr$) As Variant()
Dim M, W As WhMth
Set W = WhMthzStr(WhStr)
For Each M In MdItr(P, WhStr)
    PushIAy MthnDryzP, MthnDryzMd(CvMd(M), W)
Next
End Function

Private Function MthnDryzSrc(Src$(), Optional B As WhMth) As Variant()
Dim MthLin
For Each MthLin In Itr(MthLinyzSrc(Src))
    PushISomSi MthnDryzSrc, Mthn3(MthLin, B).MthnDr
Next
End Function


