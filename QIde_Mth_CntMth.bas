Attribute VB_Name = "QIde_Mth_CntMth"
Option Explicit
Private Const CMod$ = "MIde_Mth_Cnt."
Private Const Asm$ = "QIde"
Const MthCntPP$ = "NPubSub NPubFun NPubPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp"
Type CntgMth
    Mdn As String
    NPubSub As Integer
    NPubFun As Integer
    NPubPrp As Integer
    NPrvSub As Integer
    NPrvFun As Integer
    NPrvPrp As Integer
    NFrdSub As Integer
    NFrdFun As Integer
    NFrdPrp As Integer
End Type
Function CntgMth(Mdn, NPubSub%, NPubFun%, NPubPrp%, NPrvSub%, NPrvFun%, NPrvPrp%, NFrdSub%, NFrdFun%, NFrdPrp%) As MthCnt
With CntgMth
.Mdn = Mdn
.NPubSub = NPubSub
.NPubFun = NPubFun
.NPubPrp = NPubPrp
.NPrvSub = NPrvSub
.NPrvFun = NPrvFun
.NPrvPrp = NPrvPrp
.NFrdSub = NFrdSub
.NFrdFun = NFrdFun
.NFrdPrp = NFrdPrp
End With
End Function
Property Get NMth%(A As CntgMth)
NMth = NPubSub + NPubFun + NPubPrp + NPrvSub + NPrvFun + NPrvPrp + NFrdSub + NFrdFun + NFrdPrp
End Property
Function CntgMthLin(A As CntgMth, Optional Hdr As EmHdr)
Dim Pfx$: If Hdr = EiWiHdr Then Pfx = "Pub* | Prv* | Frd* : *{Sub Fun Frd} "
CntgMthLin = Pfx & Mdn & " | " & N & " | " & NPubSub & " " & NPubFun & " " & NPubPrp & " | " & NPrvSub & " " & NPrvFun & " " & NPrvPrp & " | " & NFrdSub & " " & NFrdFun & " " & NFrdPrp
End Function


Function NMthzS%(Src$())
NMthzS = Si(MthIxy(Src))
End Function

Function NMthPj%()
NMthPj = NMthzP(CPj)
End Function

Function NMthMd%()
NMthMd = NMthzMd(CMd)
End Function

Function NMthzP%(Pj As VBProject)
Dim O%, C As VBComponent
For Each C In Pj.VBComponents
    O = O + NMthzS(Src(C.CodeModule))
Next
NMthzP = O
End Function

Function MthDotCmlNyInVbe(Optional WhStr$) As String()
MthDotCmlNyInVbe = MthDotCmlNyzV(CVbe, WhStr)
End Function
Private Function MthDotCmlNyzV(A As Vbe, Optional WhStr$) As String()
Dim Mthn
For Each Mthn In MthNyzV(A, WhStr)
    PushI MthDotCmlNyzV, DotCml(Mthn)
Next
End Function
Function MthCmlGpAsetInVbe(Optional WhStr$) As Aset
Set MthCmlGpAsetInVbe = MthCmlGpAsetzV(CVbe, WhStr)
End Function

Function MthCmlGpAsetzV(A As Vbe, Optional WhStr$) As Aset
Dim Mthn
Set MthCmlGpAsetzV = New Aset
For Each Mthn In Itr(MthNyzV(A, WhStr))
    MthCmlGpAsetzV.PushAy CmlGp(Mthn)
Next
End Function

Function MthCmlAsetzP(P As VBProject, Optional WhStr$) As Aset
Set MthCmlAsetzP = CmlAset(MthnyzP(P))
End Function

Function CntgMthzM(A As CodeModule) As MthCnt
Dim NPubSub%, NPubFun%, NPubPrp%, NPrvSub%, NPrvFun%, NPrvPrp%, NFrdSub%, NFrdFun%, NFrdPrp%
Dim MthLin
For Each MthLin In Itr(MthLinyzSrc(Src(A)))
    With Mthn3(MthLin)
        Select Case True
        Case .IsPub And .IsSub: NPubSub = NPubSub + 1
        Case .IsPub And .IsFun: NPubFun = NPubFun + 1
        Case .IsPub And .IsPrp: NPubPrp = NPubPrp + 1
        Case .IsPrv And .IsSub: NPrvSub = NPrvSub + 1
        Case .IsPrv And .IsFun: NPrvFun = NPrvFun + 1
        Case .IsPrv And .IsPrp: NPrvPrp = NPrvPrp + 1
        Case .IsFrd And .IsSub: NFrdSub = NFrdSub + 1
        Case .IsFrd And .IsFun: NFrdFun = NFrdFun + 1
        Case .IsFrd And .IsPrp: NFrdPrp = NFrdPrp + 1
        Case Else: Thw CSub, "Invalid Mthn3", "MthLin Mthn3", MthLin, .Lin
        End Select
    End With
Next
Set MthCnt = New MthCnt
MthCnt.Init Mdn(A), NPubSub, NPubFun, NPubPrp, NPrvSub, NPrvFun, NPrvPrp, NFrdSub, NFrdFun, NFrdPrp
End Function
Function MthCntMd() As MthCnt
Set MthCntMd = MthCnt(CMd)
End Function
Sub CntMthP()
CntMthzP CPj
End Sub
Sub CntMthzP(A As VBProject)
End Sub
Function CntgMthszP(P As VBProject) As CntgMths
Dim C As VBComponent
For Each C In P.VBComponents
    PushCntgMth CntgMthszP, CntgMthzM(C.CodeModule)
Next
End Function
Function PushCntgMth(O As CntgMths, M As CntgMth)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Function

Function LyzCntgMths(A As CntgMths) As String()
Dim J&
For J = 0 To A.N - 1
    PushIAy LyzCntgMths, CntgMthLin(A.Ay(J))
Next
End Function

Function CntgMths(P As VBProject) As CntgMths
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    PushObj MthCntAy, MthCnt(C.CodeModule)
Next
End Function


Function NMthzMd%(A As CodeModule, Optional WhStr$)
NMthzMd = NMthzS(Src(A), WhStr)
End Function

Function NSrcLinPj&(P As VBProject)
Dim O&, C As VBComponent
For Each C In P.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinPj = O
End Function

Function NPMthMd%(A As CodeModule)
NPMthMd = NMthzS(Src(A), "-Pub")
End Function
Function NPMthVbe%(A As Vbe)
Dim O%, P As VBProject
For Each P In A.VBProjects
    O = O + NPMthPj(P)
Next
NPMthVbe = O
End Function
Property Get NPMth%()
NPMth = NPMthVbe(CVbe)
End Property

Function NPMthPj%(P As VBProject)
Dim O%, C As VBComponent
For Each C In P.VBComponents
    O = O + NPMthMd(C.CodeModule)
Next
NPMthPj = O
End Function

