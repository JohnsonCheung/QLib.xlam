Attribute VB_Name = "QIde_Ens_CModSub"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_CSub."

Private Function EnsgCModSub(A As CodeModule) As MdygMd
Dim S$():                     S = Src(A)
Dim MC As Mdygs:          MC = EnsgCSubs(S, MthRgs(S))               'MC = Mdyg-CSub
Dim MM As Mdyg:           MM = EnsgCMod(DclLy(S), IsUsingCMod(MC))   'MM = Mdyg-CMod
Dim M As Mdygs:            M = AddMdygs(SngMdyg(MM), MC)
                    EnsgCModSub = MdygMd(A, M)
End Function

Private Sub Z_EnsgCModSubzP()
Dim Pj As VBProject, Act As MdygMds, Ept As MdygMds
GoSub ZZ
Exit Sub
ZZ:
    BrwMdygMds EnsgCModSubzP(CPj)
    Return
Tst:
    Act = EnsgCModSub(Pj)
    Brw LyzMdygMds(Act): Stop
    Return
End Sub

Private Function EnsgCModSubzP(P As VBProject) As MdygMds
If P.Protection = vbext_pp_locked Then Thw CSub, "Pj is locked", "Pj", P.Name
Dim C As VBComponent
For Each C In P.VBComponents
    PushMdygMd MdygMdszEnsCMSub, EnsgCModSub(C.CodeModule) '<===
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
Private Function EnsgCMod(Dcl$(), UseMod As Boolean) As Mdyg
Dim N$: N = Mdn(A)
.Lnx = CModLnx(A)
.InsLno = LnoOfAftOptAndImpl(A)
.IsUsingCMod = IsUsingCMod
If UseCMod Then
Dim NLno&:  NLno = A.InsLno
Dim NLin$: NLin = CModLin(IsUsing, A.Mdn)
End If
Dim L As Lnx: L = CnstLnxzSN(Dcl, "CMod$")
Dim NLno&: If UseMod Then NLno = LnoOfAftOptAndImpl
Dim NLin$: If UseMod Then A
Dim OLno&: OLno = L.Ix + 1
Dim OLno$: OLin = L.Lin
EnsgCMod = MdygzOONN(OLno, OLin, NLno, NLin)
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

Function NxtSrcIx&(Src$(), Optional FmIx&)
Dim J&
For J = FmIx + 1 To UB(Src)
    If LasChr(Src(J - 1)) <> "_" Then
        NxtSrcIx = J
        Exit Function
    End If
Next
'No need to throw error, just exit it returns -1
'Thw CSub, "Cannot find Lno where to insert CSub of a given method", "Mthn MthLy", A.Mthn, AywFT(Src, A.FmIx, A.EIx)
NxtSrcIx = -1
End Function

Private Function IsUsingCMod(EnsgCSubs As Mdygs) As Boolean
Dim J%
For J = 0 To EnsgCSubs.N - 1
    Select Case EnsgCSubs.Ay(J).Act
    Case EmMdyg.EiIns, EmMdyg.EiRpl: IsUsingCMod = True: Exit Function
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
MdyMd EnsgCModSub(A)
End Sub

Private Function CModCnstLin$(A As CodeModule)
CModCnstLin = FmtQQ("Private Const CMod$ = ""?.""", Mdn(A))
End Function

Private Sub ZZ_EnsgCModSub()
Dim Md As CodeModule, Act As MdygMd, Ept As MdygMd
GoSub ZZ
'GoSub T0
Exit Sub
ZZ:
    BrwMdygMd EnsgCModSub(CMd)
    Return
T0:
    Set Md = CMd
    'Ept = SomInsgLin(2, "Private Const CMod$ = ""BEnsCMod.""")
    GoTo Tst
Tst:
    Act = EnsgCModSub(Md)
'    If Not IsEqMdygMd(Act, Ept) Then Stop
    Return
End Sub
Sub Z2()
ZZ_EnsgCModSub
End Sub
Private Sub ZZZ()
QIde_Ens_CModSub:
End Sub