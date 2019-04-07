Attribute VB_Name = "MIde_Mth_Nm_Get"
Option Explicit
Public Const DocOfDta_MthQVNm$ = "It is a String dervied from Nm.  Q for quoted.  V for verb.  It has 3 Patn: NoVerb-[#xxx], MidVerb-[xxx(vvv)xxx], FstVerb-[(vvv)xxx]."
Public Const DocOfNmRul_FstVerbBeingDo1$ = "The Fun will not return any value"
Public Const DocOfNmRul_FstVerbBeingDo2$ = "The Cmls aft Do is a verb"
Private Sub Z_MthNsetOfVbeWiVerb()
MthNsetOfVbeWiVerb.Srt.Vc
End Sub
Private Sub Z_DryOf_MthNm_Verb_OfVbe()
BrwDry DryOf_MthNm_Verb_OfVbe
End Sub
Function DryOf_MthNm_Verb_OfVbe() As Variant()
Dim MthNm, ODry()
For Each MthNm In Itr(MthNyOfVbe)
    PushI ODry, Sy(MthNm, Verb(MthNm))
Next
DryOf_MthNm_Verb_OfVbe = DrywDist(ODry)
End Function
Private Sub Z_MthNsetOfVbeWoVerb()
MthNsetOfVbeWoVerb.Srt.Vc
End Sub

Property Get MthNyOfVbeWiVerb() As String()
Dim MthNm, J&
For Each MthNm In Itr(MthNyOfVbe)
'    If HasSubStr(MthNm, "Z_ExprDic") Then Stop
    If J Mod 100 = 0 Then Debug.Print J
    If HasVerb(MthNm) Then PushI MthNyOfVbeWiVerb, MthNm
    J = J + 1
Next
End Property
Property Get MthNyOfVbeWoVerb() As String()
Dim MthNm
For Each MthNm In Itr(MthNyOfVbe)
    If Not HasVerb(MthNm) Then PushI MthNyOfVbeWiVerb, MthNm
Next
End Property

Function HasVerb(Nm) As Boolean
HasVerb = Verb(Nm) <> ""
End Function
Property Get MthNsetOfVbeWiVerb() As Aset
Set MthNsetOfVbeWiVerb = AsetzAy(MthNyOfVbeWiVerb)
End Property
Property Get MthNsetOfVbeWoVerb() As Aset
Set MthNsetOfVbeWoVerb = AsetzAy(MthNyOfVbeWoVerb)
End Property
Function MthNsetOfVbe(Optional WhStr$) As Aset
Set MthNsetOfVbe = AsetzAy(MthNyOfVbe(WhStr))
End Function

Function MthNyzSrcFm(Src$(), FmMthIxAy&()) As String()
Dim Ix
For Each Ix In Itr(FmMthIxAy)
    PushI MthNyzSrcFm, MthNm(Src(Ix))
Next
End Function

Function MthNyOfVbe(Optional WhStr$) As String()
MthNyOfVbe = MthNyzVbe(CurVbe, WhStr$)
End Function

Function MthNsetOfPj(Optional WhStr$) As Aset
Set MthNsetOfPj = AsetzAy(MthNyOfPj(WhStr))
End Function

Function MthNyOfPj(Optional WhStr$) As String()
MthNyOfPj = MthNyzPj(CurPj, WhStr$)
End Function

Function MthNyOfPubVbe(Optional WhStr$) As String()
MthNyOfPubVbe = MthNyzVbe(CurVbe, WhStr & " -Pub")
End Function

Function MthNyzPj(A As VBProject, Optional WhStr$) As String()
Dim M, W As WhMth
Set W = WhMthzStr(WhStr)
For Each M In MdItr(A, WhStr)
    PushIAy MthNyzPj, MthNyzMd(CvMd(M), W)
Next
End Function

Function MthQNyOfVbe(Optional WhStr$) As String()
MthQNyOfVbe = MthQNyzVbe(CurVbe, WhStr)
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
    PushI MthQNmDryzPj, MthQNmDr(QNm)
Next
End Function

Function MthQNmDr(MthQNm) As String()
Dim O$(): O = SplitDot(MthQNm)
If Si(O) <> 5 Then Thw CSub, "MthQNm should have 4 dot", "MthQNm", MthQNm
MthQNmDr = O
End Function

Function MthQNyzPj(A As VBProject, Optional WhStr$) As String()
Dim I
For Each I In MdItr(A, WhStr)
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
Property Get MMthNyOfVbe() As String()
MthNyOfVbe
End Property
Function MthNyzFb(Fb) As String()
MthNyzFb = MthNyzVbe(VbePjf(Fb))
ClsPjf Fb
End Function


Private Sub Z_MthNyFb()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
'    For Each Fb In AppFbAy
        PushAy O, MthNyzFb(Fb)
'    Next
    'Brw O
    Return
X_BrwOne:
'    Brw MthNyFb(AppFbAy()(0))
    Return
End Sub


Private Sub Z_MthNyzSrc()
Brw MthNyzSrc(CurSrc)
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
   Act = MthNyzSrc(CurSrc)
   BrwAy Act
End Sub

Function MthDNmSqzPj(A As VBProject) As Variant()
MthDNmSqzPj = MthDNySq(MthNyzPj(A, True))
End Function

Function MthDNmWszPj(A As VBProject) As Worksheet
Set MthDNmWszPj = WsVis(WszSq(MthDNmSqzPj(A)))
End Function

Function MthNyzVbe(A As Vbe, Optional WhStr$) As String()
Dim I
For Each I In PjItr(A, WhStr)
    PushIAy MthNyzVbe, MthNyzPj(CvPj(I), WhStr)
Next
End Function

Function MthAsetVbe(Optional WhStr$) As Aset
Set MthAsetVbe = AsetzAy(MthNyOfVbe(WhStr))
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
BrwAy MthDNyzSrc(CurSrc)
End Sub

Function MthDNyOfVbe(Optional WhStr$) As String()
MthDNyOfVbe = MthDNyzVbe(CurVbe, WhStr)
End Function

Function MthDNyzVbe(A As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In PjItr(A, WhStr)
    PushIAy MthDNyzVbe, MthDNyzPj(P, WhStr)
Next
End Function

Function MthDNyzMd(A As CodeModule, Optional WhStr$) As String()
MthDNyzMd = MthDNyzSrc(Src(A), WhMthzStr(WhStr))
End Function
Function MthDNyzSrc(Src$(), Optional WhStr$) As String()
Dim L, B As WhMth
Set B = WhMthzStr(WhStr)
For Each L In Itr(Src)
    PushNonBlankStr MthDNyzSrc, MthDNm(L, B)
Next
End Function

