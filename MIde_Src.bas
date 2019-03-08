Attribute VB_Name = "MIde_Src"
Option Explicit
Private P As VBProject
Private C As VBComponent
Private Sub ZZ_SrcDcl()
BrwStr DclLy(ZZSrc)
End Sub

Private Sub ZZ_FstMthIx()
Dim Act%
Act = FstMthIx(ZZSrc)
Ass Act = 2
End Sub
Function SrcMdNm(MdNm$) As String()
SrcMdNm = Src(Md(MdNm))
End Function
Private Sub ZZ_MthTopRmIx_SrcFm()
Dim ODry()
    Dim S$(): S = SrcMdNm("IdeSrcLin")
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, L
    For Each L In S
        IsMth = ""
        RmkLx = ""
        If IsMthLin(L) Then
            IsMth = "*Mth"
            RmkLx = MthTopRmkIx(S, Lx)

        End If
        Dr = Array(IsMth, RmkLx, L)
        Push ODry, Dr
        Lx = Lx + 1
    Next
BrwDrs Drs("Mth RmkLx Lin", ODry)
End Sub

Private Property Get ZZSrc() As String()
ZZSrc = Src(Md("IdeSrc"))
End Property

Private Property Get ZZSrcLin$()
ZZSrcLin = "Private Sub IsMthLin()"
End Property
Private Sub Z_MthNyzSrc()
BrwAy MthNyzSrc(SrcMdNm("AAAMod"))
End Sub

Sub AsgMthDr(MthDr, OMdy$, OTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OTopRmk$)
AsgAp MthDr, OMdy, OTy, ONm, OPrm, ORet, OLinRmk, OLines, OTopRmk
End Sub

Private Sub Z_MthLinDryzSrc()
BrwDry MthLinDryzSrc(SrcMd)
End Sub

Function MthLinDryzSrc(Src$()) As Variant()
Dim L
For Each L In Itr(Src)
    PushISomSz MthLinDryzSrc, MthLinDr(L)
Next
End Function

Function MthDNyPj(A As VBProject, Optional WhStr$) As String()
If A.Protection = vbext_pp_locked Then Exit Function
For Each C In A.VBComponents
    PushIAy MthDNyPj, MthDNyMd(C.CodeModule)
Next
End Function
Function LinesPj$(A As VBProject)
LinesPj = JnCrLf(SrczPj(A))
End Function


Function NMthzSrc%(A$())
NMthzSrc = Sz(MthIxAy(A))
End Function

Function NUsrTySrc%(A$())
If Sz(A) = 0 Then Exit Function
Dim I, O%
For Each I In A
'   If SrcLin_IsTy(I) Then O = O + 1
Next
NUsrTySrc = O
End Function

