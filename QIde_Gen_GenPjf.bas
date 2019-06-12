Attribute VB_Name = "QIde_Gen_GenPjf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf_Fba."
Private Const Asm$ = "QIde"

Public Const DoctzSrcRoot$ = "It a Pth with Fdr eq '.src'"
Public Const DoctzSrcp$ = "Src-Pth:Src.p. It a Pth with Fdr is PjFn and ParFdr is SrcRoot"
Public Const DoctzDistp$ = "It an InstPth with Fdr is InstPth and ParFdr is DistRoot"
Public Const DoctzInstPth$ = "It a Pth with Fdr is InstNm"

Private Sub Z_CompressFxa()
CompressFxa Pjf(CPj)
End Sub

Sub CompressFxa(Fxa$)
'PjExp PjzPjf(Xls.Vbe, Fxa)
Dim Srcp$: Srcp = SrcpzPjf(Fxa)
'CrtDistFxa Srcp
RplFfn Fxa, Srcp
End Sub

Function SrcRoot$(Srcp$)
SrcRoot = ParPth(Srcp)
End Function
Function DistpP$() 'Distribution Path
DistpP = Distp(SrcpP)
End Function

Function Distp$(Srcp) 'Distribution Path
Distp = AddFdrEns(UpPth(Srcp, 2), ".Dist")
End Function

Function DistFba$(Srcp)
DistFba = DistPjf(Srcp, ".accdb")
End Function

Private Function DistPjf(Srcp, Ext) '
Dim P$:  P = Distp(Srcp)
Dim F1$: F1 = RplExt(Fdr(ParPth(P)), Ext)
Dim F2$: F2 = NxtFfnzNotIn(F1, PjfnAyV)
Dim F$:  F = NxtFfnzAva(P & F2)
DistPjf = F
End Function
Private Sub Z_DistFxa()
Dim Srcp$
GoSub T0
Exit Sub
T0:
    Srcp = SrcpP
    Ept = "C:\Users\user\Documents\Projects\Vba\QLib\.Dist\QLib(002).xlam"
    GoTo Tst
Tst:
    Act = DistFxa(Srcp)
    C
    Return
End Sub
Function DistFxa$(Srcp)
DistFxa = DistPjf(Srcp, ".xlam")
End Function

Private Sub Z_DistPjf()
Dim Pth, I
For Each I In SrcpSyOfExpgInst
    Pth = I
    Debug.Print Pth
    Debug.Print DistFba(Pth)
    Debug.Print DistFxa(Pth)
    Debug.Print
Next
End Sub

Sub LoadBas(P As VBProject, Srcp$)
Dim F$(): F = BasFfny(Srcp)
Dim I: For Each BasItm In Itr(F)
    P.VBComponents.Import I
Next
End Sub

Private Function BasFfny(Srcp$) As String()
Dim F$(): F = Ffny(Srcp)
Stop
Dim I: For Each I In Itr(F)
    If IsBasFfn(I) Then
        PushI BasFfny, I
    End If
Next
End Function
Private Function IsBasFfn(Ffn) As Boolean
IsBasFfn = HasSfx(Ffn, ".bas")
End Function


Sub GenFbaP()
GenFbazP CPj
End Sub
Sub GenFbazP(P As VBProject)
Dim Acs As New Access.Application, OPj As VBProject
Dim SPth$: SPth = SrcpzP(P)
Dim OFba$: OFba = DistFba(SPth)
                  DltFfnIf OFba
                  CrtFb OFba             '<== Crt OFba
                  ExpPj P              '<== Exp
                  OpnFb Acs, OFba
        Set OPj = PjzAcs(Acs)
                  AddRfzS OPj, RfSrczSrcp(SPth)   '<== Add Rf
                  LoadBas OPj, SPth       '<== Load Bas
Dim Frm$(): Frm = FrmFfny(SPth)
Dim F: For Each F In Itr(Frm)
    Dim N$: N = RmvExt(RmvExt(F))
    Acs.LoadFromText acForm, N, F       '<== Load Frm
Next
QuitAcs Acs
End Sub

Sub GenFxaP()
GenFxazP CPj
End Sub

Sub GenFxazP(Pj As VBProject)
Dim SPth$:                     SPth = Srcp(Pj)
Dim OFxa$:               OFxa = DistFxa(SPth)
                                ExpPj Pj
                                CrtFxa OFxa
Dim OPj As VBProject: Set OPj = PjzFxa(OFxa)
                                AddRfzS OPj, RfSrczSrcp(SPth)
                                LoadBas OPj, SPth
End Sub
Private Sub ZZ()
QIde_Bld_GenFxa:
End Sub

