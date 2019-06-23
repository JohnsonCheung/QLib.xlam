Attribute VB_Name = "QIde_B_DMth"
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
    Dim O As Drs: O = DrszAdd(O, B)
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
    Dim O As Drs: O = DrszAdd(O, A)
Next
DMthczV = O
End Function

Function DrMthLin(MthLin) As Variant()
With MthLinRec(MthLin)
DrMthLin = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function

Function WsMthP() As Worksheet
Dim O As Worksheet: Set O = WszDrs(DMthP)
Dim L As ListObject: Set L = FstLo(O)
SetWdtLc L, "MthLin", 80
SetWdtLc L, "Mdn", 15
SetWdtLc L, "Mthn", 20
Set WsMthP = ShwWs(O)
End Function
Function DMthP() As Drs
Static A As Drs
If NoReczDrs(A) Then A = DMthzP(CPj)
DMthP = A
End Function

Function DMthzP(P As VBProject) As Drs
Dim C As VBComponent, ODy(), Dy(), Pjn$
Pjn = P.Name
For Each C In P.VBComponents
    Dy = DMth(C.CodeModule).Dy
    Dy = InsColzDyAv(Dy, Av(Pjn, ShtCmpTy(C.Type), C.Name))
    PushIAy ODy, Dy
Next
DMthzP = DrszFF("Pjn MdTy Mdn L Mdy Ty Mthn MthLin", ODy)
End Function

Function DMth(M As CodeModule) As Drs
'Ret : L Mdy Ty Mthn MthLin ! Mdy & Ty are Sht @@
DMth = DMthzS(Src(M))
End Function
Function DMthzM(M As CodeModule) As Drs
'Ret : L Mdy Ty Mthn MthLin ! Mdy & Ty are Sht @@
DMthzM = DMthzS(Src(M))
End Function

Function DMthe(M As CodeModule) As Drs
'Ret : L E Mdy Ty Mthn MthLin ! Mdy & Ty are Sht. L is Lno E is ELno @@d
DMthe = DMthezS(Src(M))
End Function

Function DMthc(M As CodeModule) As Drs
'Ret : L E Mdy Ty Mthn MthLin MthLy! Mdy & Ty are Sht. L is Lno E is ELno @@
DMthc = DMthczS(Src(M))
End Function

Function DMtheM() As Drs
DMtheM = DMthe(CMd)
End Function
Function DMthcM() As Drs
DMthcM = DMthczM(CMd)
End Function

Function DMthczS(Src$()) As Drs
Dim A As Drs: A = DMthzS(Src)
Dim Dy(), Dr
For Each Dr In Itr(A.Dy)
    Dim F&: F = Dr(0) - 1
    Dim E&: E = EndLix(Src, F) + 1
    Dim T&: T = E - 1
    Dim MthLy$(): MthLy = AwFT(Src, F, T)
    Dr = InsEle(Dr, E, 1)
    PushI Dr, MthLy
    PushI Dy, Dr
Next
DMthczS = DrszFF("L E Mdy Ty Mthn MthLin MthLy", Dy)
End Function

