Attribute VB_Name = "QIde_Mth_Lines"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Lines."

'aaa
Private Property Get XX1()

End Property

'BB
Private Property Let XX1(V)

End Property

Function MthLinesByPubMthNm$(PubMthNm$)
Const CSub$ = CMod & "MthLinesByPubMthNm"
Dim A$: A = PubMthNm
Dim B$(): B = ModNyzPubMthNm(A)
If Si(B) <> 1 Then
    Thw CSub, "Should be 1 module found", "PubMthNm [#Mod having PubMthNm] ModNy-Found", PubMthNm, Si(B), B
End If
MthLinesByPubMthNm = MthLinesBySrcNm(SrczMdNm(B(0)), PubMthNm, WiTopRmk:=True)
End Function
'

'
Property Get CurMthLines$()
CurMthLines = MthLinesByMdMth(CurMd, CurMthNm$, WiTopRmk:=True)
End Property

Sub VcMthLinesAyInPj()
Vc FmtCntDic(MthLinesAyInPj(WiTopRmk:=True))
End Sub
Function MthLinesAyInPj(Optional WiTopRmk As Boolean) As String()
MthLinesAyInPj = MthLinesAyByPj(CurPj, WiTopRmk)
End Function

Function MthLinesAyByPj(A As VBProject, Optional WiTopRmk As Boolean) As String()
Dim I
For Each I In MdItr(A)
    PushIAy MthLinesAyByPj, MthLinesAyByMd(CvMd(I), WiTopRmk)
Next
End Function

Function MthLinesAyByMd(A As CodeModule, Optional WiTopRmk As Boolean) As String()
MthLinesAyByMd = MthLinesAyBySrc(Src(A), WiTopRmk)
End Function

Function MthLinesAyBySrc(Src$(), Optional WiTopRmk As Boolean) As String()
Dim I, Ix&
For Each I In Itr(MthIxAy(Src))
    Ix = I
    PushI MthLinesAyBySrc, MthLinesBySrcFm(Src, Ix, WiTopRmk)
Next
End Function

Function MthLines$(MthNm$, Optional WoTopRmk As Boolean)
MthLines = MthLineszNm(MthNm$, WoTopRmk)
End Function
Function MthLineszNm$(MthNm$, Optional WoTopRmk As Boolean)

End Function

Function MthLinesByMdMth$(Md As CodeModule, MthNm$, Optional WiTopRmk As Boolean)
MthLinesByMdMth = MthLinesBySrcNm(Src(Md), MthNm, WiTopRmk)
End Function

Function MthLinesByMdNmTy$(Md As CodeModule, MthNm$, ShtMthTy$, Optional WiTopRmk As Boolean)
MthLinesByMdNmTy = MthLinesBySrcNmTy(Src(Md), MthNm, ShtMthTy, WiTopRmk)
End Function

Function MthLyByMdMth(Md As CodeModule, MthNm$, Optional WiTopRmk As Boolean) As String()
MthLyByMdMth = MthLyBySrcNm(Src(Md), MthNm, WiTopRmk)
End Function

Function MthLinesBySrcFm$(Src$(), MthFmIx&, Optional WiTopRmk As Boolean)
MthLinesBySrcFm = JnCrLf(MthLyBySrcFm(Src, MthFmIx, WiTopRmk))
End Function

Function MthLyBySrcFm(Src$(), MthFmIx&, Optional WiTopRmk As Boolean) As String()
Dim TopLy$()
If WiTopRmk Then
    TopLy = MthTopRmkLy(Src, MthFmIx)
End If
Dim ToIx&: ToIx = MthToIx(Src, MthFmIx)
Dim MthLy$(): MthLy = AywFT(Src, MthFmIx, ToIx)
MthLyBySrcFm = AddSyAp(TopLy, MthLy)
End Function

Function MthLinesBySrcNm$(Src$(), MthNm$, Optional WiTopRmk As Boolean)
MthLinesBySrcNm = JnCrLf(MthLyBySrcNm(Src, MthNm, WiTopRmk))
End Function

Function MthLinesBySrcNmTy$(Src$(), N, ShtMthTy$, Optional WiTopRmk As Boolean)
MthLinesBySrcNmTy = JnCrLf(MthLyBySrcNmTy(Src, N, ShtMthTy, WiTopRmk))
End Function

Function MthLyBySrcNm(Src$(), MthNm$, Optional WiTopRmk As Boolean) As String()
Dim I, Ix&
For Each I In Itr(MthIxAyzNm(Src, MthNm))
    Ix = I
    PushI MthLyBySrcNm, MthLinesBySrcFm(Src, Ix, WiTopRmk)
Next
End Function

Function MthLyBySrcNmTy(Src$(), N, ShtMthTy$, Optional WiTopRmk As Boolean) As String()
With MthIxzSrcNmTy(Src, N, ShtMthTy)
    If Not .Som Then Stop: Exit Function
    MthLyBySrcNmTy = MthLyBySrcFm(Src, .Lng, WiTopRmk)
End With
End Function

