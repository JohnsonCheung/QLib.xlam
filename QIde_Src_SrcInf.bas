Attribute VB_Name = "QIde_Src_SrcInf"
Option Explicit
Private Const CMod$ = "MIde_Src."
Private Const Asm$ = "QIde"
Private P As VBProject
Private C As VBComponent
Private Sub ZZ_SrcDcl()
BrwAy DclLy(ZZSrc)
End Sub

Private Sub ZZ_FstMthIx()
Dim Act%
Act = FstMthIx(ZZSrc)
Ass Act = 2
End Sub
Function SrczMdNm(MdNm$) As String()
SrczMdNm = Src(Md(MdNm))
End Function
Private Sub ZZ_MthTopRmIx_SrcFm()
Dim ODry()
    Dim Src$(): Src = SrczMdNm("IdeSrcLin")
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, Lin$, I
    For Each I In Src
        Lin = I
        IsMth = ""
        RmkLx = ""
        If IsMthLin(Lin) Then
            IsMth = "*Mth"
            RmkLx = MthTopRmkIx(Src, Lx)

        End If
        Dr = Array(IsMth, RmkLx, Lin)
        Push ODry, Dr
        Lx = Lx + 1
    Next
BrwDrs DrszFF("Mth RmkLx Lin", ODry)
End Sub

Private Property Get ZZSrc() As String()
ZZSrc = Src(Md("IdeSrc"))
End Property

Private Property Get ZZSrcLin$()
ZZSrcLin = "Private Sub IsMthLin()"
End Property
Private Sub Z_MthNyzSrc()
BrwAy MthNyzSrc(SrczMdNm("AAAMod"))
End Sub

Sub AsgMthDr(MthDr, OMdy$, OTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OTopRmk$)
AsgAp MthDr, OMdy, OTy, ONm, OPrm, ORet, OLinRmk, OLines, OTopRmk
End Sub

Private Sub Z_MthLinDryzSrc()
BrwDry MthLinDryzSrc(CurSrc)
End Sub

Function MthLinDryzSrc(Src$()) As Variant()
Dim L
For Each L In Itr(Src)
    PushISomSi MthLinDryzSrc, MthLinDr(CStr(L))
Next
End Function

Function MthDNyzPj(A As VBProject, Optional WhStr$) As String()
If A.Protection = vbext_pp_locked Then Exit Function
For Each C In A.VBComponents
    PushIAy MthDNyzPj, MthDNyzMd(C.CodeModule)
Next
End Function

Function CurSrcLines$()
CurSrcLines = SrcLineszMd(CurMd)
End Function

Function SrcLinesInPj$()
SrcLinesInPj = SrcLineszPj(CurPj)
End Function

Function SrcLineszPj$(A As VBProject)
SrcLineszPj = JnCrLf(SrczPj(A))
End Function

Function SrcLineszMd$(A As CodeModule)
If A.CountOfLines = 0 Then Exit Function
SrcLineszMd = A.Lines(1, A.CountOfLines)
End Function

Function SrczMd(A As CodeModule) As String()
SrczMd = Src(A)
End Function

Function Src(A As CodeModule) As String()
Src = SplitCrLf(SrcLineszMd(A))
End Function

Function SrcV() As String()
SrcV = SrczVbe(CurVbe)
End Function

Function SrczPj(A As VBProject) As String()
If A.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In A.VBComponents
    PushIAy SrczPj, Src(C.CodeModule)
Next
End Function

Function SrczVbe(A As Vbe) As String()
Dim P
For Each P In A.VBProjects
    PushIAy SrczVbe, SrczPj(CvPj(P))
Next
End Function

Property Get CurSrc() As String()
CurSrc = Src(CurMd)
End Property

Function NMthzSrc%(A$())
NMthzSrc = Si(MthIxAy(A))
End Function

Function NUsrTySrc%(A$())
If Si(A) = 0 Then Exit Function
Dim I, O%
For Each I In A
'   If SrcLin_IsTy(I) Then O = O + 1
Next
NUsrTySrc = O
End Function

