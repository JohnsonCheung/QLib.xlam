Attribute VB_Name = "QIde_Mth_Nm_MthNm"
Option Compare Text
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


Function QMthn$(M As CodeModule, Lin)
Dim D$: D = MthDnzLin(Lin): If D = "" Then Exit Function
QMthn = MdDNm(M) & "." & D
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

Function IsPMth(Lin) As Boolean 'Pum = PubMthn
Dim L$: L = Lin
If ShfShtMdy(L) <> "Pub" Then Exit Function
If MthTy(L) = "" Then Exit Function
IsPMth = True
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
Function PMthNyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsPMth(L) Then PushI PMthNyzS, Mthn(L)
Next
End Function

Private Sub Z_ModNyzPPm()
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
For Each I In Itr(MthNyV)
    Mthn = I
    PushI ODry, Sy(Mthn, Verb(Mthn))
Next
Dry__Mthn_Verb_InVbe = DrywDist(ODry)
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
Function DPubMthzV(A As Vbe) As String()

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
    PushNonBlank MthNyzS, Mthn(L)
Next
End Function

Function PMthNyzM(M As CodeModule) As String()
PMthNyzM = PMthNyzS(Src(M))
End Function

Private Sub ZZ()
Z_MthNyzFb
MIde_Mth_Nm:
End Sub

Function MthNyzM(M As CodeModule) As String()
MthNyzM = MthNyzS(Src(M))
End Function

Private Sub Z_MthnzS()
GoSub ZZ
Exit Sub
ZZ:
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
Function DMthnzV(V As Vbe) As Drs

End Function
Function DMthnzM(M As CodeModule) As Drs
DMthnzM = DMthn(M)
End Function
Function DMthnzP(P As VBProject) As Drs
Dim C As VBComponent
Dim Pn$: Pn = P.Name
For Each C In P.VBComponents
    Dim Mn$: Mn = C.Name
    Dim A As Drs: A = DMthn(C.CodeModule)
    Dim B As Drs: B = InsColzDrsCC(A, "Pj Md", Pn, Mn)
    Dim O As Drs: O = DrszAdd(O, A)
Next
DMthnzP = O
End Function

Function DMthnV() As Drs
DMthnV = DMthnzV(CVbe)
End Function

Function DMthnP() As Drs
DMthnP = DMthnzP(CPj)
End Function

Function DMthnM() As Drs
DMthnM = DMthnzM(CMd)
End Function

