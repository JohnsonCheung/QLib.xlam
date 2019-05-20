Attribute VB_Name = "QIde_Mth_LinAy_Drs"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Liny_Drs."

Function MthDnzDiLinV() As Dictionary
Set MthDnzDiLinV = MthDnzDiLinzV(CVbe)
End Function

Private Function MthDnzDiLinzV(A As Vbe) As Dictionary
Dim O As New Dictionary, I
For Each I In A.VBProjects
    PushDic O, MthDnzDiLinzP(CvPj(I))
Next
Set MthDnzDiLinzV = O
End Function

Function MthDnzDiLinzP(P As VBProject) As Dictionary
Dim O As New Dictionary, I, Pfx$, M As CodeModule
For Each I In MdItr(P)
    Set M = I
    PushDic O, DicAddKeyPfx(MthDnzDiLinzS(Src(M)), MdDNm(M) & ".")
Next
Set MthDnzDiLinzP = O
End Function

Function MthDnzDiLinzS(Src$()) As Dictionary
Dim L$, I
Set MthDnzDiLinzS = New Dictionary
For Each I In Itr(MthLinyzS(Src))
    L = I
    MthDnzDiLinzS.Add MthDn(L), L
Next
End Function

Function MthLinyzP(P As VBProject) As String()
Dim I
For Each I In MdItr(P)
    PushIAy MthLinyzP, MthLinyzS(Src(CvMd(I)))
Next
End Function

Function MthLinyV() As String()
MthLinyV = MthLinyzV(CVbe)
End Function

Function MthLinyzV(V As Vbe) As String()
Dim P As VBProject
For Each P In V.VBProjects
    PushIAy MthLinyzV, MthLinyzP(P)
Next
End Function

Function MthLnxszM(A As CodeModule) As Lnxs
MthLnxszM = MthLnxszS(Src(A))
End Function

Function MthLnxszS(Src$()) As Lnxs
Dim O$(), J&, B As WhMth
For J = 0 To UB(Src)
    If HitMthLin(Src(J), B) Then
        PushLnx MthLnxszS, Lnx(ContLin(Src, J), J)
    End If
Next
End Function

Function MthLinyzMd(A As CodeModule) As String()
MthLinyzMd = MthLinyzS(Src(A))
End Function

Function MthLinyzS(Src$()) As String()
Dim O$(), J&
For J = 0 To UB(Src)
    If IsMthLin(Src(J)) Then
        PushI MthLinyzS, ContLin(Src, J)
    End If
Next
End Function

Function Fny_Mthn() As String()
Fny_Mthn = SyzSS("Mthn Ty Mdy Md MdTy Pj")
End Function
Private Function MthQ1LyzMthQLy(MthQLy$()) As String()
Dim MthQLin
For Each MthQLin In Itr(MthQLy)
Next
End Function
Function MthQ1LyInVbe() As String()
MthQ1LyInVbe = MthQ1LyzMthQLy(MthQLyzV(CVbe))
End Function

Function MthQLyV() As String()
Static X
If IsEmpty(X) Then
    X = MthQLyzV(CVbe)
End If
MthQLyV = X
End Function

Function MthQLyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MthQLyzV, MthQLyzP(P)
Next
End Function

Function MthQLyzM(A As CodeModule) As String()
Dim P$
P = PjnzM(A) & "." & ShtCmpTy(A.Parent.Type) & "." & A.Parent.Name & "."
MthQLyzM = AddPfxzAy(MthLinyzS(Src(A)), P)
End Function

Function MthQLyzP(P As VBProject) As String()
Dim C As VBComponent
If P.Protection = vbext_pp_locked Then Exit Function
For Each C In P.VBComponents
    PushIAy MthQLyzP, MthQLyzM(C.CodeModule)
Next
End Function
