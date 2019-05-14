Attribute VB_Name = "QIde_Mth_LinAy_Drs"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Liny_Drs."

Function MthDnzDiLinV(Optional WhStr$) As Dictionary
Set MthDnzDiLinV = MthDnzDiLinzV(CVbe, WhStr)
End Function

Private Function MthDnzDiLinzV(A As Vbe, Optional WhStr$) As Dictionary
Dim O As New Dictionary, I
For Each I In PjItr(A, WhStr)
    PushDic O, MthDnzDiLinzP(CvPj(I), WhStr)
Next
Set MthDnzDiLinzV = O
End Function

Function MthDnzDiLinzP(P As VBProject, Optional WhStr$) As Dictionary
Dim O As New Dictionary, I, Pfx$, M As CodeModule
For Each I In MdItr(P, WhStr)
    Set M = I
    PushDic O, DicAddKeyPfx(MthDnzDiLinzS(Src(M)), MdDNm(M) & ".")
Next
Set MthDnzDiLinzP = O
End Function

Function MthDnzDiLinzS(Src$(), Optional WhStr$) As Dictionary
Dim L$, I
Set MthDnzDiLinzS = New Dictionary
For Each I In Itr(MthLinyzSrc(Src, WhStr))
    L = I
    MthDnzDiLinzS.Add MthDn(L), L
Next
End Function

Function MthLinyzP(P As VBProject, Optional WhStr$) As String()
Dim I
For Each I In MdItr(P, WhStr)
    PushIAy MthLinyzP, MthLinyzSrc(Src(CvMd(I)), WhStr)
Next
End Function

Function MthLinyV(Optional WhStr$) As String()
MthLinyV = MthLinyzV(CVbe, WhStr)
End Function

Function MthLinyzV(V As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In V.VBProjects
    PushIAy MthLinyzV, MthLinyzP(P, WhStr)
Next
End Function

Function MthLnxszM(A As CodeModule, Optional WhStr$) As Lnxs
MthLnxszM = MthLnxszSrc(Src(A), WhStr)
End Function

Function MthLnxszSrc(Src$(), Optional WhStr$) As Lnxs
Dim O$(), J&, B As WhMth
Set B = WhMthzStr(WhStr)
For J = 0 To UB(Src)
    If HitMthLin(Src(J), B) Then
        PushLnx MthLnxszSrc, Lnx(ContLin(Src, J, OneLin:=True), J)
    End If
Next
End Function

Function MthLinyzMd(A As CodeModule, Optional WhStr$) As String()
MthLinyzMd = MthLinyzSrc(Src(A), WhStr)
End Function

Function MthLinyzSrc(Src$(), Optional WhStr$) As String()
Dim O$(), J&, B As WhMth
Set B = WhMthzStr(WhStr)
For J = 0 To UB(Src)
    If HitMthLin(Src(J), B) Then
        PushI MthLinyzSrc, ContLin(Src, J, OneLin:=True)
    End If
Next
End Function

Function FnyOfMthn() As String()
FnyOfMthn = SyzSS("Mthn Ty Mdy Md MdTy Pj")
End Function
Private Function MthQ1LyzMthQLy(MthQLy$()) As String()
Dim MthQLin
For Each MthQLin In Itr(MthQLy)
Next
End Function
Function MthQ1LyInVbe(Optional WhStr$) As String()
MthQ1LyInVbe = MthQ1LyzMthQLy(MthQLyzV(CVbe, WhStr))
End Function

Function MthQLyV(Optional WhStr$) As String()
Static X
If IsEmpty(X) Then
    X = MthQLyzV(CVbe, WhStr)
End If
MthQLyV = X
End Function

Function MthQLyzV(A As Vbe, Optional WhStr$) As String()
Dim P
For Each P In PjItr(A, WhStr)
    PushIAy MthQLyzV, MthQLyzP(CvPj(P), WhStr)
Next
End Function

Function MthQLyzM(A As CodeModule, Optional WhStr$) As String()
Dim P$
P = PjnzM(A) & "." & ShtCmpTy(A.Parent.Type) & "." & A.Parent.Name & "."
MthQLyzM = AddPfxzAy(MthLinyzSrc(Src(A), WhStr), P)
End Function

Function MthQLyzP(P As VBProject, Optional WhStr$) As String()
Dim C
If P.Protection = vbext_pp_locked Then Exit Function
For Each C In CmpItr(P, WhStr)
    PushIAy MthQLyzP, MthQLyzM(CvCmp(C).CodeModule, WhStr)
Next
End Function
