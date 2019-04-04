Attribute VB_Name = "MIde_Mth_Lines"
Option Explicit
Const CMod$ = "MIde_Mth_Lines."

'aaa
Private Property Get XX1()

End Property

'BB
Private Property Let XX1(V)

End Property

Function MthLineszPub$(PubMthNm)
Const CSub$ = CMod & "MthLineszPub"
Dim A$: A = PubMthNm
Dim B$(): B = ModNyzPubMthNm(A)
If Si(B) <> 1 Then
    Thw CSub, "Should be 1 module found", "PubMthNm [#Mod having PubMthNm] ModNy-Found", PubMthNm, Si(B), B
End If
MthLineszPub = MthLineszSrcNm(SrczMdNm(B(0)), PubMthNm, WiTopRmk:=True)
End Function

Property Get MthLines$()
MthLines = MthLineszMd(CurMd, CurMthNm$, WiTopRmk:=True)
End Property

Sub BrwMthLinesAyPj()
Vc FmtCntDic(MthLinesAyPj(WiTopRmk:=True))
End Sub
Function MthLinesAyPj(Optional WiTopRmk As Boolean) As String()
MthLinesAyPj = MthLinesAyzPj(CurPj, WiTopRmk)
End Function

Function MthLinesAyzPj(A As VBProject, Optional WiTopRmk As Boolean) As String()
Dim I
For Each I In MdItr(A)
    PushIAy MthLinesAyzPj, MthLinesAyzMd(CvMd(I), WiTopRmk)
Next
End Function

Function MthLinesAyzMd(A As CodeModule, Optional WiTopRmk As Boolean) As String()
MthLinesAyzMd = MthLinesAyzSrc(Src(A), WiTopRmk)
End Function

Function MthLinesAyzSrc(Src$(), Optional WiTopRmk As Boolean) As String()
Dim I
For Each I In Itr(MthIxAy(Src))
    PushI MthLinesAyzSrc, MthLineszSrcFm(Src, I, WiTopRmk)
Next
End Function

Function MthLineszMd$(Md As CodeModule, MthNm, Optional WiTopRmk As Boolean)
MthLineszMd = MthLineszSrcNm(Src(Md), MthNm, WiTopRmk)
End Function

Function MthLineszMdNmTy$(Md As CodeModule, MthNm, ShtMthTy$, Optional WiTopRmk As Boolean)
MthLineszMdNmTy = MthLineszSrcNmTy(Src(Md), MthNm, ShtMthTy, WiTopRmk)
End Function

Function MthLyzMdMth(Md As CodeModule, MthNm, Optional WiTopRmk As Boolean) As String()
MthLyzMdMth = MthLyzSrcNm(Src(Md), MthNm, WiTopRmk)
End Function

Function MthLineszSrcFm$(Src$(), MthFmIx, Optional WiTopRmk As Boolean)
MthLineszSrcFm = JnCrLf(MthLyzSrcFm(Src, MthFmIx, WiTopRmk))
End Function
Function MthLyzSrcFm(Src$(), MthFmIx, Optional WiTopRmk As Boolean) As String()
Dim TopLy$()
If WiTopRmk Then
    TopLy = MthTopRmkLy(Src, MthFmIx)
End If
Dim ToIx&: ToIx = MthToIx(Src, MthFmIx)
Dim MthLy$(): MthLy = AywFT(Src, MthFmIx, ToIx)
MthLyzSrcFm = SyAddAp(TopLy, MthLy)
End Function

Function MthLineszSrcNm$(Src$(), N, Optional WiTopRmk As Boolean)
MthLineszSrcNm = JnCrLf(MthLyzSrcNm(Src, N, WiTopRmk))
End Function

Function MthLineszSrcNmTy$(Src$(), N, ShtMthTy$, Optional WiTopRmk As Boolean)
MthLineszSrcNmTy = JnCrLf(MthLyzSrcNmTy(Src, N, ShtMthTy, WiTopRmk))
End Function


Function MthLyzSrcNm(Src$(), N, Optional WiTopRmk As Boolean) As String()
Dim I
For Each I In Itr(MthIxAyzNm(Src, N))
    PushI MthLyzSrcNm, MthLineszSrcFm(Src, I, WiTopRmk)
Next
End Function

Function MthLyzSrcNmTy(Src$(), N, ShtMthTy$, Optional WiTopRmk As Boolean) As String()
With MthIxzSrcNmTy(Src, N, ShtMthTy)
    If Not .Som Then Stop: Exit Function
    MthLyzSrcNmTy = MthLyzSrcFm(Src, .Lng, WiTopRmk)
End With
End Function

