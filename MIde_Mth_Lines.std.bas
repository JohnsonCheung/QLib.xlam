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
Dim B$(): B = ModNyPubMthNm(A)
If Sz(B) <> 1 Then
    Thw CSub, "Should be 1 module found", "PubMthNm [#Mod having PubMthNm] ModNy-Found", PubMthNm, Sz(B), B
End If
MthLineszPub = MthLineszSrcNm(SrcMdNm(B(0)), PubMthNm, WithTopRmk:=True)
End Function

Property Get MthLines$()
MthLines = MthLineszMd(CurMd, CurMthNm$, WithTopRmk:=True)
End Property

Function MthLineszMd$(Md As CodeModule, MthNm, Optional WithTopRmk As Boolean)
MthLineszMd = MthLineszSrcNm(Src(Md), MthNm, WithTopRmk)
End Function

Function MthLineszMdNmTy$(Md As CodeModule, MthNm, ShtMthTy$, Optional WithTopRmk As Boolean)
MthLineszMdNmTy = MthLineszSrcNmTy(Src(Md), MthNm, ShtMthTy, WithTopRmk)
End Function

Function MthLyzMdMth(Md As CodeModule, MthNm, Optional WithTopRmk As Boolean) As String()
MthLyzMdMth = MthLyzSrcNm(Src(Md), MthNm, WithTopRmk)
End Function

Function MthLineszSrcFm$(Src$(), MthFmIx, Optional WithTopRmk As Boolean)
MthLineszSrcFm = JnCrLf(MthLyzSrcFm(Src, MthFmIx, WithTopRmk))
End Function
Function MthLyzSrcFm(Src$(), MthFmIx, Optional WithTopRmk As Boolean) As String()
Dim TopLy$()
If WithTopRmk Then
    TopLy = MthTopRmkLy(Src, MthFmIx)
End If
Dim ToIx&: ToIx = MthToIx(Src, MthFmIx)
Dim MthLy$(): MthLy = AywFT(Src, MthFmIx, ToIx)
MthLyzSrcFm = SyAddAp(TopLy, MthLy)
End Function

Function MthLineszSrcNm$(Src$(), N, Optional WithTopRmk As Boolean)
MthLineszSrcNm = JnCrLf(MthLyzSrcNm(Src, N, WithTopRmk))
End Function

Function MthLineszSrcNmTy$(Src$(), N, ShtMthTy$, Optional WithTopRmk As Boolean)
MthLineszSrcNmTy = JnCrLf(MthLyzSrcNmTy(Src, N, ShtMthTy, WithTopRmk))
End Function


Function MthLyzSrcNm(Src$(), N, Optional WithTopRmk As Boolean) As String()
Dim I
For Each I In Itr(MthIxAyMth(Src, N))
    PushI MthLyzSrcNm, MthLineszSrcFm(Src, I, WithTopRmk)
Next
End Function

Function MthLyzSrcNmTy(Src$(), N, ShtMthTy$, Optional WithTopRmk As Boolean) As String()
With MthIxzSrcNmTy(Src, N, ShtMthTy)
    If Not .Som Then Stop: Exit Function
    MthLyzSrcNmTy = MthLyzSrcFm(Src, .Lng, WithTopRmk)
End With
End Function

