Attribute VB_Name = "QIde_B_MthLin"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Is."
Private Const Asm$ = "QIde"
Type MthLinRec
    ShtMdy As String
    ShtTy As String
    Nm As String
    TyChr As String
    RetTy As String
    Pm As String
    Rmk As String
    IsRetVal As Boolean
    ShtRetTy As String
End Type
Private Function ShfRetTyzAftPm$(OAftPm$)
Dim A$: A = ShfTermAftAs(OAftPm)
If LasChr(A) = ":" Then
    ShfRetTyzAftPm = RmvLasChr(A)
    OAftPm = ":" & OAftPm
Else
    ShfRetTyzAftPm = A
End If
End Function
Private Function RmkzAftRetTy$(AftRetTy$)
Select Case True
Case AftRetTy = "", FstChr(AftRetTy) = ":": Exit Function
End Select
Dim L$: L = LTrim(AftRetTy)
If FstChr(L) = "'" Then RmkzAftRetTy = LTrim(RmvFstChr(L)): Exit Function
Thw CSub, "Something wrong in AftRetTy", "AftRetTy", AftRetTy
End Function
Function ArgNyzPm(Pm$) As String()
Dim Ay$(): Ay = Split(Pm, ", ")
Dim I
For Each I In Itr(Ay)
    PushI ArgNyzPm, TakNm(I)
Next
End Function
Function MthLinRec(MthLin) As MthLinRec
Dim L$: L = MthLin
With MthLinRec
    .ShtMdy = ShfShtMdy(L)
    .ShtTy = ShfShtMthTy(L): If .ShtTy = "" Then Exit Function
    .Nm = ShfNm(L)
    .TyChr = ShfTyChr(L)
    .Pm = ShfBktStr(L)
    .RetTy = ShfRetTyzAftPm(L)
    .Rmk = RmkzAftRetTy(L)
    .IsRetVal = HasEle(SyzSS("Get Fun"), .ShtTy)
    .ShtRetTy = ShtRetTy(.TyChr, .RetTy, .IsRetVal)
End With
End Function


Function MthLinzML$(M As CodeModule, Lno&)
MthLinzML = ContLinzML(M, MthLno(M, Lno))
End Function



Function MthLinAyM() As String()
MthLinAyM = MthLinAyzM(CMd)
End Function
Function MthLinAyzM(M As CodeModule) As String()
MthLinAyzM = MthLinAyzS(Src(M))
End Function
Function MthLinAyzS(Src$()) As String()
Dim Ix: For Each Ix In Itr(MthIxy(Src))
    PushI MthLinAyzS, ContLin(Src, Ix)
Next
End Function

Function MthLinAyzSN(Src$(), Mthn) As String()
Dim Ix
For Each Ix In Itr(MthIxyzSN(Src, Mthn))
    PushI MthLinAyzSN, ContLin(Src, Ix)
Next
End Function

Private Sub Z_MthLinAyzS()
Dim MthNy$(), Src$()
Src = CSrc
MthNy = Sy("Src_MthDclDy", "Mth_MthDclLin")
Ept = Sy("Function Mth_MthDclLin(A As Mth)", "Function Src_MthDclDy(A$()) As Variant()")
GoSub Tst
Exit Sub
Tst:
    Act = MthLinAyzS(Src)
    C
    Return
End Sub

Function MthLinAyP() As String()
MthLinAyP = StrCol(DoMthP, "MthLin")
End Function

Function MthLinAyzPub(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLinPubMth(L) Then PushI MthLinAyzPub, L
Next
End Function

Function MthLinAyzP(P As VBProject) As String()
MthLinAyzP = StrCol(DoMthzP(P), "MthLin")
End Function

Function MthLinAyV() As String()
MthLinAyV = MthLinAyzV(CVbe)
End Function

Function MthLinAyzV(V As Vbe) As String()
Dim P As VBProject
For Each P In V.VBProjects
    PushIAy MthLinAyzV, MthLinAyzP(P)
Next
End Function


'aaa
Private Property Get XX1()

End Property

'BB
Private Property Let XX1(V)

End Property
Function PMthNyzP(P As VBProject) As String()

End Function
Function MthLzPum(PMthn)

End Function

Function MthLzPP$(P As VBProject, PMthn)
Dim B$(): B = ModNyzPMth(PMthn)
If Si(B) <> 1 Then
    Thw CSub, "Should be 1 module found", "PMthn [#Mod having PMthn] ModNy-Found", PMthn, Si(B), B
End If
MthLzPP = MthLzSP(SrczMdn(B(0)), PMthn)
End Function
'
Function MthLzSP$(Src$(), PMthn)

End Function
'
Property Get CMthL$() 'Cur
CMthL = MthLzM(CMd, CMthn)
End Property

Sub VcMthLAyP()
Vc FmtLinesAy(MthLAyP)
End Sub
Function MthLAyP() As String()
MthLAyP = MthLAyzP(CPj)
End Function

Function MthLAyzP(P As VBProject) As String()
Dim I
For Each I In MdItr(P)
    PushIAy MthLAyP, MthLAyzM(CvMd(I))
Next
End Function

Function MthLAyzM(M As CodeModule) As String()
MthLAyzM = MthLAyzS(Src(M))
End Function

Function MthLAyzS(Src$()) As String()
Dim Ix
For Each Ix In Itr(MthIxy(Src))
    PushI MthLAyzS, MthLzSI(Src, Ix)
Next
End Function
Function MdzMthn(P As VBProject, Mthn) As CodeModule
Dim C As VBComponent, O As CodeModule
For Each C In P.VBComponents
    If HasEle(PMthNyzM(C.CodeModule), Mthn) Then
        If Not IsNothing(O) Then Thw CSub, FmtQQ("Mthn fnd in 2 or more md: [?] & [?]", Mdn(O), C.Name)
        Set O = C.CodeModule
    End If
Next
If IsNothing(O) Then Thw CSub, "Mthn not fnd in any codemodule of given pj", "Pj Mthn", "P.Name,Mthn"
End Function

Function MthLzPN$(P As VBProject, Mthn)
MthLzPN = MthLzM(MdzMthn(P, Mthn), Mthn)
End Function

Function MthLzN$(Mthn)
MthLzN = MthLzPN(CPj, Mthn)
End Function

Function MthLzM$(M As CodeModule, Mthn)
MthLzM = MthLzSN(Src(M), Mthn)
End Function

Function MthLyzM(M As CodeModule, Mthn) As String()
MthLyzM = SplitCrLf(MthLzM(M, Mthn))
End Function

Function MthLzMTN$(Md As CodeModule, ShtMthTy$, Mthn)
Dim S$(): S = Src(Md)
Dim Ix&: Ix = MthIxzSTN(S, ShtMthTy, Mthn)
MthLzMTN = MthLzSI(S, Ix)
End Function

Function MthLzSI$(Src$(), MthIx)
Dim EIx&:       EIx = EndLix(Src, MthIx)
Dim MthLy$(): MthLy = AwFT(Src, MthIx, EIx)
MthLzSI = JnCrLf(MthLy)
End Function

Function MthLinzSTN$(Src$(), ShtMthTy$, Mthn)
MthLinzSTN = Src(MthIxzSTN(Src, ShtMthTy, Mthn))
End Function

Function MthLzSN$(Src$(), Mthn)
Dim Ix, O$()
For Each Ix In Itr(MthIxyzSN(Src, Mthn))
    PushI O, MthLzSI(Src, Ix)
Next
MthLzSN = JnDblCrLf(O)
End Function

Function MthLzSTN$(Src$(), ShtMthTy$, Mthn)
Dim Ix&: Ix = MthIxzSTN(Src, ShtMthTy, Mthn)
MthLzSTN = MthLzSI(Src, Ix)
End Function


Function PMthLinAy(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsLinPubMth(L) Then PushI PMthLinAy, L
Next
End Function

Function PMthLinItr(Src$())
Asg Itr(PMthLinAy(Src)), PMthLinItr
End Function


Function NMthLin%(M As CodeModule, MthLno&)
Dim K$, J&, N&, E$
K = MthKd(M.Lines(MthLno, 1))
If K = "" Then Thw CSub, "Given MthLno is not a MthLin", "Md MthLno MthLin", Mdn(M), MthLno, M.Lines(MthLno, 1)
E = "End " & K
For J = MthLno To M.CountOfLines
    N = N + 1
    If M.Lines(J, 1) = E Then NMthLin = N: Exit Function
Next
ThwImpossible CSub
End Function

