Attribute VB_Name = "MIde_Mth_Nm"
Option Explicit
Const CMod$ = "MIde_Mth_Nm."
Type MdMth
    Md As CodeModule
    MthNm As String
End Type
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
MdQNm = MdQNmzMd(CurMd)
End Property
Property Get CurMthQNm$()
On Error GoTo X
CurMthQNm = MthQNm(CurMd, CurMthLin)
Exit Property
X: Debug.Print CSub
End Property

Function MthQNm$(A As CodeModule, Lin$)
Dim D$: D = MthDNmzLin(Lin): If D = "" Then Exit Function
MthQNm = MdQNmzMd(A) & "." & D
End Function

Function MthNyzPub(Src$()) As String()
Dim Ix, N$, B As MthNm3
For Each Ix In MthIxItr(Src)
    Set B = MthNm3(Src(Ix))
    If B.Nm <> "" Then
        If B.ShtMdy = "" Or B.ShtMdy = "Pub" Then
            PushI MthNyzPub, B.Nm
        End If
    End If
Next
End Function

Function MthNyzMthLinAy(MthLinAy$()) As String()
Const CSub$ = CMod & "MthNyzMthLinAy"
Dim I, Nm$, J%, MthLin$
For Each I In Itr(MthLinAy)
    MthLin = I
    Nm = MthNm(MthLin)
    If Nm = "" Then Thw CSub, "Given MthLinAy does not have MthNm", "[MthLin with error] Ix MthLinAy", I, J, SyAddIxPfx(MthLinAy)
    PushI MthNyzMthLinAy, Nm
    J = J + 1
Next
End Function
Function Ens2Dot(S$) As StrOpt
Select Case DotCnt(S)
Case 0: Ens2Dot = SomStr(".." & S)
Case 1: Ens2Dot = SomStr("." & S)
Case 2: Ens2Dot = SomStr(S)
End Select
End Function

Function MdMth(MthQNm$) As MdMth
Const CSub$ = CMod & "MdMthOpt"
Dim Ny$()
With Ens2Dot(MthQNm)
    If Not .Som Then Thw CSub, "MthQNm should have 2 or less dot", "MthQNm", MthQNm
    Ny = SplitDot(.Str)
End With
Set MdMth.Md = Md(Ny(0) & "." & Ny(1))
MdMth.MthNm = Ny(2)
End Function

Function RmvMthMdy$(Lin$)
RmvMthMdy = RmvTermAy(Lin, MthMdyAy)
End Function

Function MthDNmzMthNm3$(A As MthNm3)
If A.Nm = "" Then Exit Function
MthDNmzMthNm3 = A.Nm & "." & A.ShtTy & "." & A.ShtMdy
End Function

Function RmvMthNm3$(Lin$)
Dim L$: L = Lin
RmvMthMdy L
If ShfMthTy(L) = "" Then Exit Function
If ShfNm(L) = "" Then Thw CSub, "Not as SrcLin", "Lin", Lin
RmvMthNm3 = L
End Function

Function RelOf_MthSDNm_To_MdNm_InVbe(Optional WhStr$) As Rel
'Set MthSDNmToMdNmRelInVbe = MthSDNmToMdNmRelzVbe(CurVbe, WhStr)
End Function

Function RelOf_MthSDNm_To_MdNm_zVbe(A As Vbe, Optional WhStr$) As Rel
Dim MthQNm
For Each MthQNm In MthQNyzVbe(A, WhStr)
'    PushI MthSQNmToMdNmRelzVbe, MthSDNm
Next
'Set MthSDNmToMdNmRelInVbe = MthSDNmToMdNmRelzVbe(CurVbe, WhStr)
End Function

Function MthNm3(Lin$, Optional B As WhMth) As MthNm3
Dim L$: L = Lin
Dim O As New MthNm3
With O
    .MthMdy = ShfMthMdy(L)
    .MthTy = ShfMthTy(L)
    If .MthTy = "" Then Set MthNm3 = New MthNm3: Exit Function
    .Nm = Nm(L)
End With
If HitMthNm3(O, B) Then
    Set MthNm3 = O
Else
    Set MthNm3 = New MthNm3
End If
End Function

Function MthNm$(Lin$, Optional B As WhMth)
If Not IsMthLin(Lin) Then Exit Function
MthNm = MthNm3(Lin, B).Nm
End Function

Function MthNmzMthDNm$(MthDNm$)
If MthDNm = "*Dcl" Then MthNmzMthDNm = MthDNm: Exit Function
Dim A$()
A = SplitDot(MthDNm)
If Si(A) <> 3 Then Thw CSub, "MthDNm should have 2 dot", "MthDNm", MthDNm
MthNmzMthDNm = A(0)
End Function

Function MthDNmzLin$(MthLin$)
MthDNmzLin = MthDNmzMthNm3(MthNm3(MthLin))
End Function
Function MthSQNyInVbe(Optional WhStr$) As String()
MthSQNyInVbe = MthSQNyzVbe(CurVbe, WhStr)
End Function
Function MthSQNyzVbe(A As Vbe, Optional WhStr$) As String()
Dim MthQNm
For Each MthQNm In Itr(MthQNyzVbe(A, WhStr))
    PushI MthSQNyzVbe, MthSQNm(CStr(MthQNm))
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
Function MthDNm$(Lin$, Optional B As WhMth)
MthDNm = MthNm3(Lin, B).DNm
End Function

Function MthNmzLin$(Lin$, Optional B As WhMth)
MthNmzLin = MthNm3(Lin, B).Nm
End Function

Function PrpNm$(Lin$)
Dim L$
L = RmvMdy(Lin)
If ShfKd(L) <> "Property" Then Exit Function
PrpNm = Nm(L)
End Function

Function MthNmzDNm$(MthNm$)
Dim Ay$(): Ay = Split(MthNm, ".")
Dim Nm$
Select Case Si(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
MthNmzDNm = Nm
End Function
Private Sub Z_MthNm()
GoTo ZZ
Dim A$
A = "Function MthNm$(A)": Ept = "MthNm.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = MthNm(A)
    C
    Return
ZZ:
    Dim O$(), L
    For Each L In SrczVbe(CurVbe)
        PushNonBlankStr O, MthNm(CStr(L))
    Next
    Brw O
End Sub

Function MthMdy$(Lin$)
MthMdy = FstEleEv(MthMdyAy, T1(Lin))
End Function

Function MthKd$(Lin$)
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
Property Get Rel0MthNm2MdNm() As Rel
Dim O As New Rel

End Property
Function ModNyzPubMthNm(PubMthNm) As String()
Dim I, A$
A = PubMthNm
For Each I In ModItr
    If HasEle(MthNyzPub(Src(CvMd(I))), A) Then PushI ModNyzPubMthNm, MdNm(CvMd(I))
Next
End Function

Function MthTy$(Lin$)
MthTy = PfxzPfxAyPlusSpc(RmvMthMdy(Lin), MthTyAy)
End Function

Private Sub Z_MthTy()
Dim O$(), L
For Each L In SrczMdNm("Fct")
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


Private Sub Z()
Z_MthKd
Z_MthTy
MIde_Mth_Lin_XX:
End Sub

