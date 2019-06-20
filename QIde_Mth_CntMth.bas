Attribute VB_Name = "QIde_Mth_CntMth"
Option Compare Text
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
Function DMthCntP() As Drs
DMthCntP = DMthCntzP(CPj)
End Function

Function DMthCntzP(P As VBProject) As Drs
Dim C As VBComponent, Dry(): For Each C In P.VBComponents
    PushI Dry, DrMthCnt(C.CodeModule, C.Name)
Next
DMthCntzP = DrszFF("Mdn NLines NMth NPubSub NPubFun NPubPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp", Dry)
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

Private Function DrMthCnt(M As CodeModule, Mdn$) As Variant()
Dim S$(): S = Src(M)
Dim Mth$(): Mth = MthLinAyzS(S)
Dim L: For Each L In Itr(Mth)
    With Mthn3zL(L)
        Dim Prv As Boolean: Prv = False
        Dim Pub As Boolean: Pub = False
        Dim Frd As Boolean: Frd = False
        Dim Fun As Boolean: Fun = False
        Dim Sbr As Boolean: Sbr = False
        Dim Prp As Boolean: Prp = False
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
            
    Select Case True
        Case Pub And Sbr: Dim NPubSub%: NPubSub = NPubSub + 1
        Case Pub And Fun: Dim NPubFun%: NPubFun = NPubFun + 1
        Case Pub And Prp: Dim NPubPrp%: NPubPrp = NPubPrp + 1
        Case Prv And Sbr: Dim NPrvSub%: NPrvSub = NPrvSub + 1
        Case Prv And Fun: Dim NPrvFun%: NPrvFun = NPrvFun + 1
        Case Prv And Prp: Dim NPrvPrp%: NPrvPrp = NPrvPrp + 1
        Case Frd And Sbr: Dim NFrdSub%: NFrdSub = NFrdSub + 1
        Case Frd And Fun: Dim NFrdFun%: NFrdFun = NFrdFun + 1
        Case Frd And Prp: Dim NFrdPrp%: NFrdPrp = NFrdPrp + 1
        Case Else: Thw CSub, "Invalid Mthn3", "MthLin", L
    End Select
    Dim NMth%: NMth = NMth + 1
Next
Dim NLin&: NLin = Si(S)
DrMthCnt = Array(Mdn, NLin, NMth, NPubSub, NPubFun, NPubPrp, NPrvSub, NPrvPrp, NFrdSub, NFrdFun, NFrdPrp)
End Function

Sub CntMthP()
CntMthzP CPj
End Sub
Sub CntMthzP(A As VBProject)
End Sub

Function NMthzM%(M As CodeModule)
NMthzM = NMthzS(Src(M))
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
    If IsLinPubMth(L) Then PushI PMthLinAy, L
Next
End Function

Function PMthLinItr(Src$())
Asg Itr(PMthLinAy(Src)), PMthLinItr
End Function

Function NPMthzS%(Src$())
NPMthzS = NItr(PMthLinItr(Src))
End Function

Function NPMthzM%(M As CodeModule)
NPMthzM = NPMthzS(Src(M))
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
