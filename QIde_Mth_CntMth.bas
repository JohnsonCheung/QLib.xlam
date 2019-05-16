Attribute VB_Name = "QIde_Mth_CntMth"
Option Explicit
Private Const CMod$ = "MIde_Mth_Cnt."
Private Const Asm$ = "QIde"
Const CntgMthPP$ = "NPubSub NPubFun NPubPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp"
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
Type CntgMths: N As Long: Ay() As CntgMth: End Type
Function CntgMth(Mdn, NPubSub%, NPubFun%, NPubPrp%, NPrvSub%, NPrvFun%, NPrvPrp%, NFrdSub%, NFrdFun%, NFrdPrp%) As CntgMth
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
With A
NMth = .NPubSub + .NPubFun + .NPubPrp + .NPrvSub + .NPrvFun + .NPrvPrp + .NFrdSub + .NFrdFun + .NFrdPrp
End With
End Property
Function FmtCntgMth(A As CntgMth, Optional Hdr As EmHdr)
Dim Pfx$: If Hdr = EiWiHdr Then Pfx = "Pub* | Prv* | Frd* : *{Sub Fun Frd} "
With A
Dim N%: N = NMth(A)
FmtCntgMth = JnAp(" | ", Pfx, .Mdn, N) & " | " & JnSpcAp(.NPubSub, .NPubFun, .NPubPrp, .NPrvSub, .NPrvFun, .NPrvPrp, .NFrdSub, .NFrdFun, .NFrdPrp)
End With
End Function

Function NMthzS%(Src$())
NMthzS = Si(MthIxy(Src))
End Function

Function NMthP%()
NMthP = NMthzP(CPj)
End Function

Function NMthM%()
NMthM = NMthzM(CMd)
End Function

Function NMthzP%(Pj As VBProject)
Dim O%, C As VBComponent
For Each C In Pj.VBComponents
    O = O + NMthzS(Src(C.CodeModule))
Next
NMthzP = O
End Function

Function MthDotCmlNyInVbe() As String()
MthDotCmlNyInVbe = MthDotCmlNyzV(CVbe)
End Function
Private Function MthDotCmlNyzV(A As Vbe) As String()
Dim Mthn
For Each Mthn In MthnyzV(A)
    PushI MthDotCmlNyzV, DotCml(Mthn)
Next
End Function
Function MthCmlGpAsetInVbe() As Aset
Set MthCmlGpAsetInVbe = MthCmlGpAsetzV(CVbe)
End Function

Function MthCmlGpAsetzV(A As Vbe) As Aset
Dim Mthn
Set MthCmlGpAsetzV = New Aset
For Each Mthn In Itr(MthnyzV(A))
    MthCmlGpAsetzV.PushAy CmlGp(Mthn)
Next
End Function

Function MthCmlAsetzP(P As VBProject) As Aset
Set MthCmlAsetzP = CmlAset(MthnyzP(P))
End Function

Function CntgMthzM(A As CodeModule) As CntgMth
Dim L
Dim Pub As Boolean, Prv As Boolean, Frd As Boolean
Dim Sbr As Boolean, Fun As Boolean, Prp As Boolean
For Each L In Itr(MthLinyzS(Src(A)))
    With Mthn3zL(L)
        Select Case .ShtMdy
        Case "Prv": Prv = True
        Case "Pub": Pub = True
        Case "Frd": Frd = True
        Case Else: Thw CSub, "Out of valid value: Prv PUb Frd", "ShtMdy", .ShtMdy
        End Select
        Select Case ShtMthKdzShtMthTy(.ShtTy)
        Case "Fun": Fun = True
        Case "Sub": Sbr = True
        Case "Prp": Prp = True
        Case Else: Thw CSub, "Out of valid value: Sub Fun Prp", "ShtMdy", .ShtMdy
        End Select
    End With
    With CntgMthzM
        Select Case True
        Case Pub And Sbr: .NPubSub = .NPubSub + 1
        Case Pub And Fun: .NPubFun = .NPubFun + 1
        Case Pub And Prp: .NPubPrp = .NPubPrp + 1
        Case Prv And Sbr: .NPrvSub = .NPrvSub + 1
        Case Prv And Fun: .NPrvFun = .NPrvFun + 1
        Case Prv And Prp: .NPrvPrp = .NPrvPrp + 1
        Case Frd And Sbr: .NFrdSub = .NFrdSub + 1
        Case Frd And Fun: .NFrdFun = .NFrdFun + 1
        Case Frd And Prp: .NFrdPrp = .NFrdPrp + 1
        Case Else: Thw CSub, "Invalid Mthn3", "MthLin", L
        End Select
    End With
Next
End Function
Function CntgMthM() As CntgMth
CntgMthM = CntgMthzM(CMd)
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

Function FmtCntgMths(A As CntgMths) As String()
Dim J&
For J = 0 To A.N - 1
    PushIAy FmtCntgMths, FmtCntgMth(A.Ay(J))
Next
End Function

Function CntgMths(P As VBProject) As CntgMths
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    PushCntgMth CntgMths, CntgMthzM(C.CodeModule)
Next
End Function

Function NMthzM%(A As CodeModule)
NMthzM = NMthzS(Src(A))
End Function

Function NSrcLinPj&(P As VBProject)
Dim O&, C As VBComponent
For Each C In P.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinPj = O
End Function
Function PMthLinAy(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsPMthLin(L) Then PushI PMthLinAy, L
Next
End Function

Function PMthLinItr(Src$())
Asg Itr(PMthLinAy(Src)), PMthLinItr
End Function

Function NPMthzS%(Src$())
NPMthzS = NItr(PMthLinItr(Src))
End Function

Function NPMthzM%(A As CodeModule)
NPMthzM = NPMthzS(Src(A))
End Function

Function NPMthzV%(A As Vbe)
Dim O%, P As VBProject
For Each P In A.VBProjects
    O = O + NPMthzP(P)
Next
NPMthzV = O
End Function

Property Get NPMthV%()
NPMthV = NPMthzV(CVbe)
End Property

Function NPMthzP%(P As VBProject)
Dim O%, C As VBComponent
For Each C In P.VBComponents
    O = O + NPMthzM(C.CodeModule)
Next
NPMthzP = O
End Function
