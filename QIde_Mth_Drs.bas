Attribute VB_Name = "QIde_Mth_Drs"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Drs."
Private Const Asm$ = "QIde"

Function DMthczFxa(Fxa$, Optional Xls As Excel.Application) As Drs
Dim A As Excel.Application: Set A = DftXls(Xls)
DMthczFxa = DMthczP(PjzFxa(Fxa))
If IsNothing(Xls) Then QuitXls Xls
End Function

Function DMthczM(M As CodeModule) As Drs
DMthczM = DMthczS(Src(M))
End Function

Function DMthcP() As Drs
DMthcP = DMthczP(CPj)
End Function

Function DMthczP(P As VBProject) As Drs
Dim Pjn$: Pjn = P.Name
Dim C As VBComponent: For Each C In P.VBComponents
    Dim A As Drs: A = DMthczM(C.CodeModule)
    Dim B As Drs: B = InsColzDrsC3(A, "Pjn MdTy Mdn", Pjn, ShtCmpTy(C.Type), C.Name)
    Dim O As Drs: O = AddDrs(O, B)
Next
DMthczP = O
End Function

Function DMthczPjf(Pjf) As Drs
Dim V As Vbe, App, P As VBProject, PjDte As Date
OpnPjf Pjf ' Either Excel.Application or Access.Application
Set V = VbezPjf(Pjf)
Set P = PjzPjf(V, Pjf)
Select Case True
Case IsFb(Pjf):  PjDte = PjDtezAcs(CvAcs(App))
Case IsFxa(Pjf): PjDte = DtezFfn(Pjf)
Case Else: Stop
End Select
DMthczPjf = DrsAddCol(DMthczP(P), "PjDte", PjDte)
If IsFb(Pjf) Then
    CvAcs(App).CloseCurrentDatabase
End If
End Function

Function DMthczPjfy(Pjfy$()) As Drs
Dim F
For Each F In Pjfy
    ApdDrs DMthczPjfy, DMthczPjf(F)
Next
End Function

Function DMthczV(V As Vbe) As Drs
Dim P As VBProject: For Each P In V.VBProjects
    Dim A As Drs: A = DMthczP(P)
    Dim O As Drs: O = AddDrs(O, A)
Next
DMthczV = O
End Function

Function Dr_MthLin(MthLin) As Variant()
'If Not HitMthLin(MthLin, B) Then Exit Function
Dim X As MthLinRec
X = MthLinRec(MthLin)
With X
Dr_MthLin = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function
