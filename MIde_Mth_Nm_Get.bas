Attribute VB_Name = "MIde_Mth_Nm_Get"
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

Function MthQNyVbe(Optional WhStr$) As String()
MthQNyVbe = MthQNyzVbe(CurVbe, WhStr)
End Function

Function MthQNyzVbe(A As Vbe, Optional WhStr$) As String()
Dim I
For Each I In PjItr(A, WhStr)
    PushAy MthQNyzVbe, MthQNyzPj(CvPj(I), WhStr)
Next
End Function

Function MthQNyzMd(A As CodeModule, Optional WhStr$) As String()
MthQNyzMd = AyAddPfx(MthDNyzSrc(Src(A), WhStr), MdQNmzMd(A) & ".")
End Function

Function MthQNmDryzPj(A As VBProject, Optional WhStr$) As Variant()
Dim QNm
For Each QNm In Itr(MthQNyzPj(A, WhStr))
    PushI MthQNmDryzPj, DrzMthQNm(QNm)
Next
End Function

Function DrzMthQNm(MthQNm) As String()
Dim O$(): O = SplitDot(MthQNm)
If Sz(O) <> 5 Then Thw CSub, "MthQNm should have 4 dot", "MthQNm", MthQNm
DrzMthQNm = O
End Function

Function MthQNyzPj(A As VBProject, Optional WhStr$) As String()
Dim I
For Each I In MdItrzPj(A, WhStr)
    PushAy MthQNyzPj, MthQNyzMd(CvMd(I), WhStr)
Next
End Function

Function MthDNyzVbezPub(A As Vbe, Optional WhStr$) As String()
MthDNyzVbezPub = MthDNyzVbe(A, WhStr & " -Pub")
End Function

Function MthDNyzMthNm(A As Vbe, MthNm$) As String()
Dim P As VBProject, M, Md As CodeModule
For Each P In A.VBProjects
    For Each M In P.VBComponents
        PushIAy MthDNyzMthNm, MthDNyzMdMthNm(CvMd(M), MthNm)
    Next
Next
End Function
Function MthDNyzMdMthNm(Md As CodeModule, MthNm$) As String()

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



Private Sub ZZ_MthNyzSrc()
Dim Act$()
   Act = MthNyzSrc(SrcMd)
   BrwAy Act
End Sub

Function MthDNmSqzPj(A As VBProject) As Variant()
MthDNmSqzPj = MthDNySq(MthNyzPj(A, True))
End Function

Function MthDNmWszPj(A As VBProject) As Worksheet
Set MthDNmWszPj = WsVis(WszSq(MthDNmSqzPj(A)))
End Function

Function MthDNyzMd(A As CodeModule) As String()
MthDNyzMd = MthDNyzSrc(Src(A))
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

Private Sub Z_MthDNy()
Dim Md1 As CodeModule
Set Md1 = Md("AAAMod")
BrwPth MthNyzMd(Md1)
BrwAy MthDNyzMd(Md1)
End Sub


Private Sub Z_MthDNyzSrc()
BrwAy MthDNyzSrc(SrcMd)
End Sub

Function MthDNy(Optional WhStr$) As String()
MthDNy = MthDNyzVbe(CurVbe, WhStr)
End Function

Function MthDNyzVbe(A As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In PjItr(A, WhStr)
    PushIAy MthDNyzVbe, MthDNyPj(P, WhStr)
Next
End Function

Function MthDNyMd(A As CodeModule, Optional WhStr$) As String()
MthDNyMd = MthDNyzSrc(Src(A), WhStr)
End Function

Function MthDNyzSrc(Src$(), Optional WhStr$) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlankStr MthDNyzSrc, MthDNmzLin(L)
Next
End Function


