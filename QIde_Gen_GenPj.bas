Attribute VB_Name = "QIde_Gen_GenPj"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf_Fba."
Private Const Asm$ = "QIde"
':SrcRoot: :Pth #Src-Root#          is a :Pth.  Its :Fdr eq ".Src"
':Srcp:    :Pth #Src-Path#          is a :Pth.  Its :Fdr is a PjFn and par fdr is :SrcRoot
':Disp:    :Pth #Distribution-Path# is a :Pth.  d
':InstPth: :Pth #Instance-Path#     of a @pth is any :TimNm :Fdr under @pth
':TimNm:   :Nm

Function Fxa$(FxaNm, Srcp)
Fxa = Distp(Srcp) & FxaNm & ".xlam"
End Function

Function Fba$(FbaNm, Srcp)
Fba = EnsPth(Srcp & "Dist") & FbaNm & ".accdb"
End Function

Private Sub Z_CompressFxa()
CompressFxa Pjf(CPj)
End Sub

Sub CompressFxa(Fxa$)
ExpPj PjzPjf(Exl.Vbe, Fxa)
Dim Srcp$: Srcp = SrcpzPjf(Fxa)
GenFxazSrcp Srcp
'BackupFfn Fxa, Srcp
End Sub

Private Function SrcRoot$(Srcp$)
'Ret: :SrcRoot @@
SrcRoot = ParPth(Srcp)
End Function

Function DistpP$() 'Distribution Path
DistpP = Distp(SrcpP)
End Function

Private Function Distp$(Srcp) 'Distribution Path
Distp = AddFdrEns(UpPth(Srcp, 2), ".Dist")
End Function

Private Function DistFba$(Srcp)
DistFba = PjfzSrcp(Srcp, ".accdb")
End Function

Private Function PjfzSrcp(Srcp, Ext) '
Dim P$:   P = Distp(Srcp)
Dim F1$: F1 = RplExt(Fdr(ParPth(P)), Ext)
Dim F2$: F2 = NxtFfnzNotIn(F1, PjfnAyV)
Dim F$:   F = NxtFfnzAva(P & F2)
   PjfzSrcp = F
End Function

Private Sub Z_FxazSrcp()
Dim Srcp$
GoSub T0
Exit Sub
T0:
    Srcp = SrcpP
    Ept = "C:\Users\user\Documents\Projects\Vba\QLib\.Dist\QLib(002).xlam"
    GoTo Tst
Tst:
    Act = FxazSrcp(Srcp)
    C
    Return
End Sub

Private Function FxazSrcp$(Srcp)
FxazSrcp = PjfzSrcp(Srcp, ".xlam")
End Function

Private Sub LoadBas(P As VBProject, Srcp$)
Dim F$(): F = BasFfny(Srcp)
Dim I: For Each I In Itr(F)
    P.VBComponents.Import I
Next
End Sub
Private Sub LoadBas3(P As VBProject, Srcp$)
Dim F$(): F = BasFfny(Srcp)
Dim J%, I: For Each I In Itr(F)
    P.VBComponents.Import I
    J = J + 1
    If J > 3 Then Exit Sub
Next
End Sub

Private Function BasFfny(Srcp$) As String()
Dim F$(): F = Ffny(Srcp)
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

Private Sub GenFbazP(P As VBProject)
Dim Acs As New Access.Application, OPj As VBProject
Dim SPth$:     SPth = SrcpzP(P)
Dim OFba$:     OFba = DistFba(SPth)
:                     DltFfnIf OFba
:                     CrtFb OFba                    ' <== Crt OFba
:                     ExpPj P                       ' <== Exp
:                     OpnFb Acs, OFba
            Set OPj = PjzAcs(Acs)
:                     AddRfzS OPj, RfSrczSrcp(SPth) ' <== Add Rf
:                     LoadBas OPj, SPth             ' <== Load Bas
Dim Frm$():     Frm = FrmFfny(SPth)
Dim F: For Each F In Itr(Frm)
    Dim N$: N = RmvExt(RmvExt(F))
:               Acs.LoadFromText acForm, N, F ' <== Load Frm
Next
#If False Then
'Following code is not able to save
Dim Vbe As Vbe: Set Vbe = Acs.Vbe
Dim C As VBComponent: For Each C In Acs.Vbe.ActiveVBProject.VBComponents
    C.Activate
    BoSavzV(Vbe).Execute
    Acs.Eval "DoEvents"
Next
#End If
MsgBox "Go access to save....."
QuitAcs Acs
Inf CSub, "Fba is created", "Fba", OFba
End Sub

Sub GenFxaP()
GenFxazP CPj
End Sub
Private Sub GenFxazSrcp(Srcp$)

End Sub

Private Sub GenFxazP(Pj As VBProject)
Dim SPth$:               SPth = Srcp(Pj)
Dim OFxa$:               OFxa = FxazSrcp(SPth)
:                               ExpPj Pj                                 ' <== Export
:                               CrtFxa OFxa                              ' <== Crt
Dim OPj As VBProject: Set OPj = PjzFxa(OFxa)
:                               AddRfzS OPj, RfSrczSrcp(SPth)            ' <== Add Rf
:                               LoadBas OPj, SPth                        ' <== Load Bas
:                               Inf CSub, "Fxa is created", "Fxa", OFxa
End Sub

Private Sub Z()
QIde_Gen_GenPj:
End Sub

'
