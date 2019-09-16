Attribute VB_Name = "MxParseMthLin"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxParseMthLin."
Type MthLinRec
    ShtMdy As String
    ShtTy As String
    NM As String
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

Private Function MthLinRmkzAftRetTy$(AftRetTy$)
Select Case True
Case AftRetTy = "", FstChr(AftRetTy) = ":": Exit Function
End Select
Dim L$: L = LTrim(AftRetTy)
If FstChr(L) = "'" Then MthLinRmkzAftRetTy = LTrim(RmvFstChr(L)): Exit Function
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
    .NM = ShfNm(L)
    .TyChr = ShfTyChr(L)
    .Pm = ShfBktStr(L)
    .RetTy = ShfRetTyzAftPm(L)
    .Rmk = MthLinRmkzAftRetTy(L)
    .IsRetVal = HasEle(SyzSS("Get Fun"), .ShtTy)
    .ShtRetTy = ShtRetTy(.TyChr, .RetTy, .IsRetVal)
End With
End Function


Function MthLinzML$(M As CodeModule, Lno&)
MthLinzML = ContLinzM(M, MthLno(M, Lno))
End Function

Function MthLinAyM() As String()
MthLinAyM = MthLinAyzM(CMd)
End Function
Function MthLinAyzM(M As CodeModule) As String()
MthLinAyzM = MthLinAy(Src(M))
End Function
Function MthLinAy(Src$()) As String()
Dim Ix: For Each Ix In Itr(MthIxy(Src))
    PushI MthLinAy, ContLin(Src, Ix)
Next
End Function

Function MthLinAyN(Src$(), Mthn) As String()
Dim Ix
For Each Ix In Itr(MthIxyzN(Src, Mthn))
    PushI MthLinAyN, ContLin(Src, Ix)
Next
End Function

Private Sub Z_MthLinAy()
Dim MthNy$(), Src$()
Src = CSrc
MthNy = Sy("Src_MthDclDy", "Mth_MthDclLin")
Ept = Sy("Function Mth_MthDclLin(A As Mth)", "Function Src_MthDclDy(A$()) As Variant()")
GoSub Tst
Exit Sub
Tst:
    Act = MthLinAy(Src)
    C
    Return
End Sub

Property Get CMthLin$()
CMthLin = MthLinzML(CMd, CLno)
End Property

Function MthLinAyP() As String()
MthLinAyP = StrCol(DoPubFun, "MthLin")
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
Private Sub SetXX1(V)

End Sub
Function PubMthNyzP(P As VBProject) As String()

End Function
Function MthLzPum(PubMthn)

End Function

Function MthLzPP$(P As VBProject, PubMthn)
Dim B$(): B = ModNyzPubMth(PubMthn)
If Si(B) <> 1 Then
    Thw CSub, "Should be 1 module found", "PubMthn [#Mod having PubMthn] ModNy-Found", PubMthn, Si(B), B
End If
MthLzPP = MthLzSP(SrczMdn(B(0)), PubMthn)
End Function
'
Function MthLzSP$(Src$(), PubMthn)

End Function
'
Property Get CMthL$() 'Cur
CMthL = MthLzM(CMd, CMthn)
End Property

Sub VcMthlAyP()
Vc FmtLinesAy(MthlAyP)
End Sub

Function MthlAyP() As String()
MthlAyP = MthlAyzP(CPj)
Stop
End Function

Function MthlAyzP(P As VBProject) As String()
Dim I
For Each I In MdItr(P)
    PushIAy MthlAyP, MthlAyzM(CvMd(I))
Next
End Function

Function MthlAyzM(M As CodeModule) As String()
MthlAyzM = MthlAyzS(Src(M))
End Function

Function MthlAyzS(Src$()) As String()
Dim Ix
For Each Ix In Itr(MthIxy(Src))
    PushI MthlAyzS, MthLzIx(Src, Ix)
Next
End Function
Function MdzMthn(P As VBProject, Mthn) As CodeModule
Dim C As VBComponent, O As CodeModule
For Each C In P.VBComponents
    If HasEle(PubMthNyzM(C.CodeModule), Mthn) Then
        If Not IsNothing(O) Then Thw CSub, "Mthn fnd in 2 or more md", "Mthn Mdn", Mdn(O), C.Name
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
MthLzM = MthLzNm(Src(M), Mthn)
End Function

Function MthLyzM(M As CodeModule, Mthn) As String()
MthLyzM = SplitCrLf(MthLzM(M, Mthn))
End Function

Function MthLzNmTy$(M As CodeModule, Mthn, ShtMthTy$)
Dim S$(): S = Src(M)
Dim Ix&: Ix = MthIxzNmTy(S, Mthn, ShtMthTy)
MthLzNmTy = MthLzIx(S, Ix)
End Function

Function MthLzIx$(Src$(), MthIx)
Dim EIx&:       EIx = EndLix(Src, MthIx)
Dim MthLy$(): MthLy = AwFT(Src, MthIx, EIx)
MthLzIx = JnCrLf(MthLy)
End Function

Function MthLinzNmTy$(Src$(), Mthn, ShtMthTy$)
MthLinzNmTy = Src(MthIxzNmTy(Src, Mthn, ShtMthTy))
End Function

Function MthLzNm$(Src$(), Mthn)
Dim Ix, O$()
For Each Ix In Itr(MthIxyzN(Src, Mthn))
    PushI O, MthLzIx(Src, Ix)
Next
MthLzNm = JnDblCrLf(O)
End Function

Function MthLzSTN$(Src$(), ShtMthTy$, Mthn)
Dim Ix&: Ix = MthIxzNmTy(Src, Mthn, ShtMthTy)
MthLzSTN = MthLzIx(Src, Ix)
End Function


Function PubMthLinAy(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsLinPubMth(L) Then PushI PubMthLinAy, L
Next
End Function

Function PubMthLinItr(Src$())
Asg Itr(PubMthLinAy(Src)), PubMthLinItr
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