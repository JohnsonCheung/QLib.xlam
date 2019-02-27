Attribute VB_Name = "MIde_Mth_Nm_Ay"
Option Explicit

Function MthNset(Optional WhStr$) As Aset
Set MthNset = AsetzAy(MthNyPj(WhStr))
End Function

Function MthNyzSrcFm(Src$(), FmMthIxAy&()) As String()
Dim Ix
For Each Ix In Itr(FmMthIxAy)
    PushI MthNyzSrcFm, MthNm(Src(Ix))
Next
End Function

Function MthNyVbe(Optional WhStr$) As String()
MthNyVbe = MthNyzVbe(CurVbe, WhStr$)
End Function

Function MthAsetPj(Optional WhStr$) As Aset
Set MthAsetPj = AsetzAy(MthNyPj(WhStr))
End Function
Function MthNyPj(Optional WhStr$) As String()
MthNyPj = MthNyzPj(CurPj, WhStr$)
End Function

Function MthNyPubVbe(Optional WhStr$) As String()
MthNyPubVbe = MthNyzVbe(CurVbe, WhStr & " -Pub")
End Function

Function MthNyzPj(A As VBProject, Optional WhStr$) As String()
Dim M, W As WhMth
Set W = WhMthzStr(WhStr)
For Each M In MdItrzPj(A, WhStr)
    PushIAy MthNyzPj, MthNyzMd(CvMd(M), W)
Next
End Function

Function MthQNyVbe(A As Vbe, Optional WhStr$) As String()
Dim I
For Each I In PjItr(A, WhStr)
    PushAy MthQNyVbe, MthNyzPj(CvPj(I), WhStr)
Next
End Function

Function PubMthDNySrcVbe(A As Vbe, Optional WhStr$) As String()
PubMthDNySrcVbe = MthDNyVbe(A, WhStr & " -Pub")
End Function

Function MthDNySrcNm(A As Vbe, MthNm$) As String()
Dim P As VBProject, M, Md As CodeModule
For Each P In A.VBProjects
    For Each M In P.VBComponents
        PushIAy MthDNySrcNm, MthDNyMdMthNm(CvMd(M), MthNm)
    Next
Next
End Function
Function MthDNyMdMthNm(Md As CodeModule, MthNm$) As String()

End Function
Function MthNyFb(Fb) As String()
MthNyFb = MthNyzVbe(VbePjf(Fb))
ClsPjf Fb
End Function


Private Sub Z_MthNyFb()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
'    For Each Fb In AppFbAy
        PushAy O, MthNyFb(Fb)
'    Next
    'Brw O
    Return
X_BrwOne:
'    Brw MthNyFb(AppFbAy()(0))
    Return
End Sub


Private Sub Z_MthNyzSrc()
Brw MthNyzSrc(SrcMd)
End Sub

Function MthNyzSrc(Src$(), Optional B As WhMth) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlankStr MthNyzSrc, MthNm(L, B)
Next
End Function

Function MthNyPubzMd(A As CodeModule, Optional WhStr$) As String()
MthNyPubzMd = MthNyzSrc(Src(A), WhMthzStr(WhStr))
End Function

Private Sub Z()
Z_MthNyFb
Z_MthNyzSrc
MIde_Mth_Nm:
End Sub


Function MthNyzMd(A As CodeModule, Optional B As WhMth) As String()
MthNyzMd = MthNyzSrc(Src(A), B)
End Function



Function ModNyPubMthNm(PubMthNm) As String()
Dim I, A$
A = PubMthNm
For Each I In ModItr
    If HasEle(MthNyzPub(Src(CvMd(I))), A) Then PushI ModNyPubMthNm, MdNm(CvMd(I))
Next
End Function

Private Sub ZZ_MthNyzSrc()
Dim Act$()
   Act = MthNyzSrc(SrcMd)
   BrwAy Act
End Sub


Function PjMthQDNySq(A As VBProject) As Variant()
PjMthQDNySq = MthQDNySq(MthNyzPj(A, True))
End Function

Function PjMthQDNyWs(A As VBProject) As Worksheet
Set PjMthQDNyWs = WsVis(WszSq(PjMthQDNySq(A)))
End Function

Function MthQDNyMd(A As CodeModule) As String()
MthQDNyMd = MthDNySrc(Src(A))
End Function

Function MthDDNyMd(A As CodeModule) As String()
MthDDNyMd = MthDNySrc(Src(A))
End Function

Function MthNyzVbe(A As Vbe, Optional WhStr$) As String()
Dim I
For Each I In PjItr(A, WhStr)
    PushIAy MthNyzVbe, MthNyzPj(CvPj(I), WhStr)
Next
End Function

Function MthAsetVbe(Optional WhStr$) As Aset
Set MthAsetVbe = AsetzAy(MthNyVbe(WhStr))
End Function


Property Get CurMthNyzMd() As String()
CurMthNyzMd = MthNyzMd(CurMd)
End Property

Private Sub Z_MthDDNy()
Dim Md1 As CodeModule
Set Md1 = Md("AAAMod")
BrwPth MthNyzMd(Md1)
BrwAy MthDDNyMd(Md1)
End Sub

