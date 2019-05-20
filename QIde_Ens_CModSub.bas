Attribute VB_Name = "QIde_Ens_CModSub"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_CSub."
Type MdMdyg
    Md As CodeModule
    NewLines As String
End Type

Private Function EnsgCModSub(M As CodeModule) As MdMdyg
Dim S$():          S = Src(M)
Dim MC As Mdygs:  MC = EnsgCSubs(S, MthRgs(S))                      'MC = Mdyg-CSub
Dim MM As Mdygs:  MM = EnsgCMod(DclLy(S), Mdn(M), IsUsingCMod(MC))  'MM = Mdyg-CMod
Dim Ms As Mdygs:  Ms = AddMdygs(MM, MC)
Dim NL$:          NL = JnCrLf(MdySrc(S, Ms))
         EnsgCModSub = MdMdyg(M, NL)
End Function
Function MdMdyg(M As CodeModule, NewLines$) As MdMdyg
Set MdMdyg.Md = M
MdMdyg.NewLines = NewLines
End Function
Private Sub Z_EnsgCModSubzP()
Dim Pj As VBProject, Act As MdMdyg, Ept As MdMdyg
GoSub ZZ
Exit Sub
ZZ:
    BrwRplgMds EnsgCModSubzP(CPj)
    Return
Tst:
    Act = EnsgCModSub(Pj)
'    Brw FmtMdMdyg(Act): Stop
    Return
End Sub

Private Function EnsgCModSubzP(P As VBProject) As RplgMds
If P.Protection = vbext_pp_locked Then Thw CSub, "Pj is locked", "Pj", P.Name
Dim C As VBComponent
For Each C In P.VBComponents
    'PushRplgMd RplgMdszEnsCMSub, EnsgCModSub(C.CodeModule) '<===
Next
End Function


Private Function CModLin$(IsUsing As Boolean, Mdn)
If IsUsing Then CModLin = FmtQQ("Const CMod$ = ""?.""", Mdn)
End Function

Private Function EnsgCSubs(Src$(), Mths As MthRgs) As Mdygs
Dim J%
For J = 0 To Mths.N - 1
    PushMdygs EnsgCSubs, EnsgCSub(Src, Mths.Ay(J))
Next
End Function
Function CnstLinIx&(Src$(), Cnstn)
Dim L, O&
For Each L In Itr(Src)
    If IsLinzCnstn(L, Cnstn) Then CnstLinIx = O: Exit Function
    O = O - 1
Next
CnstLinIx = -1
End Function
Function IsLinzCnstn(L, Cnstn) As Boolean
End Function
Sub AA()
Dim A
A = "ASD"
Debug.Print StrPtr(A)
End Sub
Private Function EnsgCMod(Dcl$(), Mdn$, UseMod As Boolean) As Mdygs
Dim OL As SomLnx, NL As SomLnx, NewLno&, OldCModLno&
OldCModLno = CnstLinIx(Dcl, "CMod") + 1
OL = OldCModLin(Dcl, OldCModLno)
NL = NewCModLin(UseMod, Mdn, NewLno)
EnsgCMod = MdygszON(OL, NL)
End Function

Private Function OldCModLin(Dcl$(), OldCModLno&) As SomLnx

End Function

Private Function NewCModLin(UseMod As Boolean, Mdn$, Lno&) As SomLnx
If Not UseMod Then Exit Function
Dim L$: L = FmtQQ("Private CMod$ = ""?.""", Mdn)
NewCModLin = SomLnx(Lnx(L, Lno - 1))
End Function

Private Function EnsgCSub(Src$(), Mth As MthRg) As Mdygs
Dim MthLy$(): MthLy = AywFE(Src, Mth.FmIx, Mth.EIx)
Dim O As SomLnx: O = SomOldCSub(MthLy, Mth.FmIx, Mth.EIx)
Dim N As SomLnx: N = SomNewCSub(MthLy, Mth.FmIx, Mth.Mthn)
EnsgCSub = MdygszON(O, N)
End Function

Private Function SomNewCSub(MthLy$(), MthIx&, Mthn$) As SomLnx
If Not IsUsingCSub(MthLy) Then Exit Function
Dim Lin$: Lin = FmtQQ("Const CSub$ = CMod & ""?""", Mthn)
Dim Ix&: Ix = NxtSrcIx(MthLy) + MthIx
SomNewCSub = SomLnx(Lnx(Lin, Ix))
End Function

Private Function SomOldCSub(Src$(), FmIx&, EIx&) As SomLnx
Dim Ix&
For Ix = FmIx + 1 To EIx - 2
    If HasPfx(Src(Ix), "Const CSub") Then
        SomOldCSub = SomLnx(Lnx(Src(Ix), Ix))
        Exit Function
    End If
Next
End Function

Private Function IsUsingCMod(EnsgCSubs As Mdygs) As Boolean
Dim J%
For J = 0 To EnsgCSubs.N - 1
    Select Case EnsgCSubs.Ay(J).Act
'    Case EmMdyg.EiIns, EmMdyg.EiRpl: IsUsingCMod = True: Exit Function
    End Select
Next
End Function
Private Function IsUsingCSub(MthLy$()) As Boolean
Dim L
IsUsingCSub = True
For Each L In Itr(MthLy)
    If HasSubStr(L, "CSub, ") Then Exit Function
    If HasSubStr(L, "(CSub") Then Exit Function
Next
IsUsingCSub = False
End Function

Sub EnsCModSubP()
EnsCModSubzP CPj
End Sub
Sub EnsCModSubM()
EnsCModSubzM CMd
End Sub
Sub EnsCModSubzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsCModSubzM C.CodeModule
Next
End Sub
Sub EnsCModSubzM(A As CodeModule)
'MdyMd EnsgCModSub(A)
End Sub

Private Function CModCnstLin$(A As CodeModule)
CModCnstLin = FmtQQ("Private Const CMod$ = ""?.""", Mdn(A))
End Function

Private Sub ZZ_EnsgCModSub()
Dim Md As CodeModule, Act As RplgMd, Ept As RplgMd
GoSub ZZ
'GoSub T0
Exit Sub
ZZ:
    'BrwRplgMd EnsgCModSub(CMd)
    Return
T0:
    Set Md = CMd
    'Ept = SomInsg(2, "Private Const CMod$ = ""BEnsCMod.""")
    GoTo Tst
Tst:
'    Act = EnsgCModSub(Md)
'    If Not IsEqRplgMd(Act, Ept) Then Stop
    Return
End Sub
Sub Z2()
ZZ_EnsgCModSub
End Sub

Private Sub ZZZ()
QIde_Ens_CModSub:
End Sub
Private Sub ZZ_FmtEnsCSubzMd()
Dim Md As CodeModule
'GoSub ZZ1
GoSub ZZ2
Exit Sub
ZZ1:
    Set Md = CMd
    GoTo Tst
ZZ2:
    Dim M
    For Each M In MdItr(CPj)
        Dim O$()
        'O = FmtEnsCSubzMd(CvMd(M))
        If Si(O) > 0 Then Brw O, Mdn(CvMd(M))
    Next
    Return
Tst:
    'Act = FmtEnsCSubzMd(Md)
    Brw Act
    Return
End Sub

