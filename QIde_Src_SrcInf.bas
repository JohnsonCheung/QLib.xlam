Attribute VB_Name = "QIde_Src_SrcInf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Src."
Private Const Asm$ = "QIde"
Private P As VBProject
Private C As VBComponent
Private Sub Y__SrcDcl()
BrwAy DclLy(Y_Src)
End Sub

Private Sub Y__FstMthIx()
Dim Act%
Act = FstMthIxzS(Y_Src)
Ass Act = 2
End Sub
Function SrczMdn(Mdn) As String()
SrczMdn = Src(Md(Mdn))
End Function
Private Sub Y__MthTopRmIx_SrcFm()
Dim ODry()
    Dim Src$(): Src = SrczMdn("IdeSrcLin")
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, Lin, I
    For Each I In Src
        Lin = I
        IsMth = ""
        RmkLx = ""
        If IsMthLin(Lin) Then
            IsMth = "*Mth"
            RmkLx = TopRmkIx(Src, Lx)

        End If
        Dr = Array(IsMth, RmkLx, Lin)
        Push ODry, Dr
        Lx = Lx + 1
    Next
BrwDrs DrszFF("Mth RmkLx Lin", ODry)
End Sub

Private Property Get Y_Src() As String()
Y_Src = Src(Md("IdeSrc"))
End Property

Private Property Get Y_SrcLin()
Y_SrcLin = "Private Sub IsMthLin()"
End Property

Sub AsgMthDr(MthDr, OMdy$, OTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OTopRmk$)
AsgAp MthDr, OMdy, OTy, ONm, OPrm, ORet, OLinRmk, OLines, OTopRmk
End Sub

Private Sub Z_Dry_MthLinzS()
BrwDry Dry_MthLinzS(CSrc)
End Sub

Function Dry_MthLinzS(Src$()) As Variant()
Dim L
For Each L In Itr(Src)
    PushISomSi Dry_MthLinzS, Dr_MthLin(CStr(L))
Next
End Function

Function MthDNyzP(P As VBProject) As String()
If P.Protection = vbext_pp_locked Then Exit Function
For Each C In P.VBComponents
    PushIAy MthDNyzP, MthDNyzM(C.CodeModule)
Next
End Function

Function CSrcLines$()
CSrcLines = SrcLineszM(CMd)
End Function

Function SrcLinesP$()
SrcLinesP = SrcLineszP(CPj)
End Function

Function SrcLineszP$(P As VBProject)
SrcLineszP = JnCrLf(SrczP(P))
End Function

Function SrczMd(A As CodeModule) As String()
SrczMd = Src(A)
End Function
Function CSrc() As String()
CSrc = Src(CMd)
End Function
Function Src(A As CodeModule) As String()
Src = SplitCrLf(SrcLineszM(A))
End Function

Function SrcV() As String()
SrcV = SrczV(CVbe)
End Function


Function SrczP(P As VBProject) As String()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy SrczP, Src(C.CodeModule)
Next
End Function
Function SrczV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy SrczV, SrczP(P)
Next
End Function

Function NMthzS%(Src$())
NMthzS = Si(MthIxy(Src))
End Function

Function NUsrTySrc%(A$())
If Si(A) = 0 Then Exit Function
Dim I, O%
For Each I In A
'   If SrcLin_IsTy(I) Then O = O + 1
Next
NUsrTySrc = O
End Function

