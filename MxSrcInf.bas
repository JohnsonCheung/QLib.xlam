Attribute VB_Name = "MxSrcInf"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrcInf."
Function SrczMdn(Mdn) As String()
SrczMdn = Src(Md(Mdn))
End Function
Sub Y__MthTopRmIx_SrcFm()
Dim ODy()
    Dim Src$(): Src = SrczMdn("IdeSrcLin")
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, Lin, I
    For Each I In Src
        Lin = I
        IsMth = ""
        RmkLx = ""
        If IsLinMth(Lin) Then
            IsMth = "*Mth"
            RmkLx = TopRmkIx(Src, Lx)

        End If
        Dr = Array(IsMth, RmkLx, Lin)
        Push ODy, Dr
        Lx = Lx + 1
    Next
BrwDrs DrszFF("Mth RmkLx Lin", ODy)
End Sub

Function Y_Src() As String()
Y_Src = Src(Md("IdeSrc"))
End Function

Property Get Y_SrcLin()
Y_SrcLin = "Sub IsLinMth()"
End Property

Sub AsgMthDr(MthDr, OMdy$, OTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OTopRmk$)
AsgAp MthDr, OMdy, OTy, ONm, OPrm, ORet, OLinRmk, OLines, OTopRmk
End Sub

Function CSrcl$()
CSrcl = Srcl(CMd)
End Function

Function SrcLP$()
SrcLP = SrcLzP(CPj)
End Function

Function SrcLzP$(P As VBProject)
SrcLzP = JnCrLf(SrczP(P))
End Function

Function SrczMd(M As CodeModule) As String()
SrczMd = Src(M)
End Function
Function CSrc() As String()
':CSrc: :Src #Cur-Src#
CSrc = Src(CMd)
End Function
Sub Z_VbExmLy()
Vc VbExmLy(SrczP(CPj))
End Sub
Function VbExmLy(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLinVbExmRmk(L) Then PushI VbExmLy, L
Next
End Function
Function IsLinVbExmRmk(Lin) As Boolean
Dim L$: L = LTrim(Lin)
If Not ShfPfx(L, "'") Then Exit Function
L = LTrim(L)
If Not ShfPfx(L, "!") Then Exit Function
IsLinVbExmRmk = True
End Function
Function Rmkl$(VbExmRmk$())
Dim O$()
Dim L: For Each L In Itr(VbExmRmk)
    PushNB O, RmkLin(L)
Next
Rmkl = JnCrLf(O)
End Function

Function RmkLin$(VbExmLin)
Dim L$: L = LTrim(VbExmLin)
If Not ShfPfx(L, "'") Then Exit Function
L = LTrim(L)
If Not ShfPfx(L, "!") Then Exit Function
RmkLin = Trim(L)
End Function
Function VbRmk(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLinVbRmk(L) Then PushI VbRmk, L
Next
End Function

Function SrczFc(M As CodeModule, Fc As Fc) As String()
SrczFc = SplitCrLf(M.Lines(Fc.FmLno, Fc.Cnt))
End Function

Function SrcwSngDblQ(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If HasSngDblQ(L) Then PushI SrcwSngDblQ, L
Next
End Function

Function Src(M As CodeModule) As String()
Src = SplitCrLf(Srcl(M))
End Function

Function SrczM(M As CodeModule) As String()
SrczM = SplitCrLf(Srcl(M))
End Function

Function SrcV() As String()
SrcV = SrczV(CVbe)
End Function

Function WrdAyP() As String()
WrdAyP = WrdAyzP(CPj)
End Function

Function WrdAyzP(P As VBProject) As String()
Dim L: For Each L In SrczP(P)
    PushIAy WrdAyzP, WrdAy(L)
Next
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

Function NTySrc%(A$())
If Si(A) = 0 Then Exit Function
Dim I, O%
For Each I In A
'   If SrcLin_IsTy(I) Then O = O + 1
Next
NTySrc = O
End Function


Property Get NSrcLin&()
NSrcLin = NSrcLinzP(CPj)
End Property

Function NSrcLinzP&(P As VBProject)
Dim O&, C As VBComponent
If P.Protection = vbext_pp_locked Then Exit Function
For Each C In P.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinzP = O
End Function
