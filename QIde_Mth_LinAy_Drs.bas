Attribute VB_Name = "QIde_Mth_LinAy_Drs"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_LinAy_Drs."
Function MthLinDic(Optional WhStr$) As Dictionary
Set MthLinDic = MthLinDiczVbe(CurVbe, WhStr)
End Function
Private Function MthLinDiczVbe(A As Vbe, Optional WhStr$) As Dictionary
Dim O As New Dictionary, I
For Each I In PjItr(A, WhStr)
    PushDic O, MthLinDiczPj(CvPj(I), WhStr)
Next
Set MthLinDiczVbe = O
End Function
Function MthLinDiczPj(A As VBProject, Optional WhStr$) As Dictionary
Dim O As New Dictionary, I, Pfx$, M As CodeModule
For Each I In MdItr(A, WhStr)
    Set M = I
    PushDic O, DicAddKeyPfx(MthLinDiczSrc(Src(M)), MdDNm(M) & ".")
Next
Set MthLinDiczPj = O
End Function

Function MthLinDiczSrc(Src$(), Optional WhStr$) As Dictionary
Dim L$, I
Set MthLinDiczSrc = New Dictionary
For Each I In Itr(MthLinSyzSrc(Src, WhStr))
    L = I
    MthLinDiczSrc.Add MthDNm(L), L
Next
End Function

Function MthLinSyzPj(A As VBProject, Optional WhStr$) As String()
Dim I
For Each I In MdItr(A, WhStr)
    PushIAy MthLinSyzPj, MthLinSyzSrc(Src(CvMd(I)), WhStr)
Next
End Function

Function MthLinSyInVbe(Optional WhStr$) As String()
MthLinSyInVbe = MthLinSyzVbe(CurVbe, WhStr)
End Function

Function MthLinSyzVbe(V As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In V.VBProjects
    PushIAy MthLinSyzVbe, MthLinSyzPj(P, WhStr)
Next
End Function

Function MthLnxAyzMd(A As CodeModule, Optional WhStr$) As Lnx()
MthLnxAyzMd = MthLnxAyzSrc(Src(A), WhStr)
End Function

Function MthLnxAyzSrc(Src$(), Optional WhStr$) As Lnx()
Dim O$(), J&, B As WhMth
Set B = WhMthzStr(WhStr)
For J = 0 To UB(Src)
    If HitMthLin(Src(J), B) Then
        PushObj MthLnxAyzSrc, Lnx(J, ContLin(Src, J, OneLin:=True))
    End If
Next
End Function

Function MthLinSyzMd(A As CodeModule, Optional WhStr$) As String()
MthLinSyzMd = MthLinSyzSrc(Src(A), WhStr)
End Function

Function MthLinSyzSrc(Src$(), Optional WhStr$) As String()
Dim O$(), J&, B As WhMth
Set B = WhMthzStr(WhStr)
For J = 0 To UB(Src)
    If HitMthLin(Src(J), B) Then
        PushI MthLinSyzSrc, ContLin(Src, J, OneLin:=True)
    End If
Next
End Function

Function MthNmFny() As String()
MthNmFny = SyzSsLin("MthNm Ty Mdy Md MdTy Pj")
End Function
Private Function MthQ1LyzMthQLy(MthQLy$()) As String()
Dim MthQLin
For Each MthQLin In Itr(MthQLy)
Next
End Function
Function MthQ1LyInVbe(Optional WhStr$) As String()
MthQ1LyInVbe = MthQ1LyzMthQLy(MthQLyzVbe(CurVbe, WhStr))
End Function

Function MthQLyV(Optional WhStr$) As String()
Static X
If IsEmpty(X) Then
    X = MthQLyzVbe(CurVbe, WhStr)
End If
MthQLyV = X
End Function

Function MthQLyzVbe(A As Vbe, Optional WhStr$) As String()
Dim P
For Each P In PjItr(A, WhStr)
    PushIAy MthQLyzVbe, MthQLyzPj(CvPj(P), WhStr)
Next
End Function

Function MthQLyzMd(A As CodeModule, Optional WhStr$) As String()
Dim P$
P = PjNmzMd(A) & "." & ShtCmpTy(A.Parent.Type) & "." & A.Parent.Name & "."
MthQLyzMd = AddPfxzSy(MthLinSyzSrc(Src(A), WhStr), P)
End Function

Function MthQLyzPj(A As VBProject, Optional WhStr$) As String()
Dim C
If A.Protection = vbext_pp_locked Then Exit Function
For Each C In CmpItr(A, WhStr)
    PushIAy MthQLyzPj, MthQLyzMd(CvCmp(C).CodeModule, WhStr)
Next
End Function
