Attribute VB_Name = "QIde_DoMth_DoMth"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Drs."
Private Const Asm$ = "QIde"

Function DoMthczFxa(Fxa$, Optional Exl As Excel.Application) As Drs
Dim A As Excel.Application: Set A = DftXls(Exl)
DoMthczFxa = DoMthczP(PjzFxa(Fxa))
If IsNothing(Exl) Then QuitXls Exl
End Function

Function DoMthczM(M As CodeModule) As Drs
DoMthczM = DoMthczS(Src(M))
End Function

Function DoMthcP() As Drs
DoMthcP = DoMthczP(CPj)
End Function

Function DoMthczP(P As VBProject) As Drs
Dim Pjn$: Pjn = P.Name
Dim C As VBComponent: For Each C In P.VBComponents
    Dim A As Drs: A = DoMthczM(C.CodeModule)
    Dim B As Drs: B = InsColzDrsC3(A, "Pjn MdTy Mdn", Pjn, ShtCmpTy(C.Type), C.Name)
    Dim O As Drs: O = AddDrs(O, B)
Next
DoMthczP = O
End Function

Function DoMthczPjf(Pjf) As Drs
Dim V As Vbe, App, P As VBProject, PjDte As Date
OpnPjf Pjf ' Either Excel.Application or Access.Application
Set V = VbezPjf(Pjf)
Set P = PjzPjf(V, Pjf)
Select Case True
Case IsFb(Pjf):  PjDte = PjDtezAcs(CvAcs(App))
Case IsFxa(Pjf): PjDte = DtezFfn(Pjf)
Case Else: Stop
End Select
DoMthczPjf = AddCol(DoMthczP(P), "PjDte", PjDte)
If IsFb(Pjf) Then
    CvAcs(App).CloseCurrentDatabase
End If
End Function

Function DoMthczPjfy(Pjfy$()) As Drs
Dim F
For Each F In Pjfy
    ApdDrs DoMthczPjfy, DoMthczPjf(F)
Next
End Function

Function DoMthczV(V As Vbe) As Drs
Dim P As VBProject: For Each P In V.VBProjects
    Dim A As Drs: A = DoMthczP(P)
    Dim O As Drs: O = AddDrs(O, A)
Next
DoMthczV = O
End Function

Function DrMthLin(MthLin) As Variant()
With MthLinRec(MthLin)
DrMthLin = Array(.ShtMdy, .ShtTy, .Nm, .ShtRetTy, FmtPm(.Pm, IsNoBkt:=True), .Rmk)
End With
End Function

Function WsMthP() As Worksheet
Dim O As Worksheet: Set O = WszDrs(DoPubMth)
Dim L As ListObject: Set L = FstLo(O)
SetWdtLc L, "MthLin", 80
SetWdtLc L, "Mdn", 15
SetWdtLc L, "Mthn", 20
Set WsMthP = ShwWs(O)
End Function

Function DoPubMth() As Drs
DoPubMth = DwEq(DoMthP, "Mdy", "Pub")
End Function

Function MthQnzMthn$(Mthn)
Dim D As Drs: D = DwEq(DoMthP, "Mthn", Mthn)
Select Case Si(D.Dy)
Case 0: InfLin CSub, "No such Mthn[" & Mthn & "]"
Case 1:
    Dim IxMdn%: IxMdn = IxzAy(D.Fny, "Mdn")
    MthQnzMthn = D.Dy(0)(IxMdn) & "." & Mthn
Case Else
    InfLin CSub, "No then one Md has Mthn[" & Mthn & "]"
    IxMdn = IxzAy(D.Fny, "Mdn")
    Dim Dr: For Each Dr In D.Dy
        Debug.Print Dr(IxMdn) & "." & Mthn
    Next
End Select
End Function

Function DoMthP() As Drs
DoMthP = DoMthzP(CPj)
End Function

Function DoMthzP(P As VBProject) As Drs
Dim C As VBComponent, ODy(), Dy(), Pjn$
Pjn = P.Name
For Each C In P.VBComponents
    Dy = DoMthzM(C.CodeModule).Dy
    Dy = InsColzDyAv(Dy, Av(Pjn, ShtCmpTy(C.Type), C.Name))
    PushIAy ODy, Dy
Next
DoMthzP = Drs(FoMth, ODy)
End Function

Function FoMth() As String()
FoMth = SyzSS("Pjn MdTy Mdn L Mdy Ty Mthn MthLin")
End Function

Function DoMthzM(M As CodeModule) As Drs
'Ret : L Mdy Ty Mthn MthLin ! Mdy & Ty are Sht @@
DoMthzM = DoMthzS(Src(M))
End Function

Function DoMthezM(M As CodeModule) As Drs
DoMthezM = DoMthe(Src(M))
End Function

Function DoMthc(M As CodeModule) As Drs
'Ret : L E Mdy Ty Mthn MthLin MthLy! Mdy & Ty are Sht. L is Lno E is ELno @@
DoMthc = DoMthczS(Src(M))
End Function

Function DoMtheM() As Drs
DoMtheM = DoMthezM(CMd)
End Function

Function DoMthcM() As Drs
DoMthcM = DoMthczM(CMd)
End Function

Function DoMthczS(Src$()) As Drs
Dim A As Drs: A = DoMthzS(Src)
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
DoMthczS = DrszFF("L E Mdy Ty Mthn MthLin MthLy", Dy)
End Function

