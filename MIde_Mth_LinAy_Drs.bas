Attribute VB_Name = "MIde_Mth_LinAy_Drs"
Option Explicit
Const CMod$ = "MIde_MthLinAy."
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
Dim L
Set MthLinDiczSrc = New Dictionary
For Each L In Itr(MthLinAyzSrc(Src, WhStr))
    MthLinDiczSrc.Add MthDNm(L), L
Next
End Function
Function MthLinAyzPj(A As VBProject, Optional WhStr$) As String()
Dim I
For Each I In MdItr(A, WhStr)
    PushIAy MthLinAyzPj, MthLinAyzSrc(Src(CvMd(I)), WhStr)
Next
End Function

Function MthLinAyOfVbe(Optional WhStr$) As String()
MthLinAyOfVbe = MthLinAyzVbe(CurVbe, WhStr)
End Function

Function MthLinAyzVbe(V As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In V.VBProjects
    PushIAy MthLinAyzVbe, MthLinAyzPj(P, WhStr)
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

Function MthLinAyzMd(A As CodeModule, Optional WhStr$) As String()
MthLinAyzMd = MthLinAyzSrc(Src(A), WhStr)
End Function

Function MthLinAyzSrc(Src$(), Optional WhStr$) As String()
Dim O$(), J&, B As WhMth
Set B = WhMthzStr(WhStr)
For J = 0 To UB(Src)
    If HitMthLin(Src(J), B) Then
        'If HasPfx(Src(J), "Function AyDic_RsKF") Then Stop
        PushI MthLinAyzSrc, ContLin(Src, J, OneLin:=True)
    End If
Next
End Function

Function MthNmFny() As String()
MthNmFny = SySsl("MthNm Ty Mdy Md MdTy Pj")
End Function
Private Function MthQ1LyzMthQLy(MthQLy$()) As String()
Dim MthQLin
For Each MthQLin In Itr(MthQLy)
Next
End Function
Function MthQ1LyOfVbe(Optional WhStr$) As String()
MthQ1LyOfVbe = MthQ1LyzMthQLy(MthQLyzVbe(CurVbe, WhStr))
End Function

Function MthQLyOfVbe(Optional WhStr$) As String()
MthQLyOfVbe = MthQLyzVbe(CurVbe, WhStr)
End Function

Function MthQLyzVbe(A As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In PjItr(A, WhStr)
    PushIAy MthQLyzVbe, MthQLyzPj(P, WhStr)
Next
End Function

Function MthQLyzMd(A As CodeModule) As String()
Dim P$
P = A.Parent.Collection.Parent.Name & "." & ShtCmpTy(A.Parent.Type) & "." & A.Parent.Name & "."
MthQLyzMd = AyAddPfx(MthLinAyzSrc(Src(A)), P)
End Function

Function MthQLyzPj(A As VBProject, Optional WhStr$) As String()
Dim C
If A.Protection = vbext_pp_locked Then Exit Function
For Each C In CmpItr(A, WhStr)
    PushIAy MthQLyzPj, MthQLyzMd(CvCmp(C).CodeModule)
Next
End Function
