Attribute VB_Name = "MIde_Cnt_Mth"
Option Explicit
Const MthCntPP$ = "NPubSub NPubFun NPubPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp"
Function NMthzSrc%(Src$())
NMthzSrc = Si(MthIxAy(Src))
End Function

Function NMthPj%()
NMthPj = NMthzPj(CurPj)
End Function

Function NMthMd%()
NMthMd = NMthzMd(CurMd)
End Function
Function NMthzPj%(Pj As VBProject)
Dim O%, C As VBComponent
For Each C In Pj.VBComponents
    O = O + NMthzSrc(Src(C.CodeModule))
Next
NMthzPj = O
End Function

Function MthCmlPj() As Aset
Set MthCmlPj = MthCmlzPj(CurPj)
End Function

Function MthCmlzPj(A As VBProject) As Aset
Set MthCmlzPj = CmlAset(JnSpc(MthNyzPj(A)))
End Function

Function MthCnt(A As CodeModule) As MthCnt
Dim NPubSub%, NPubFun%, NPubPrp%, NPrvSub%, NPrvFun%, NPrvPrp%, NFrdSub%, NFrdFun%, NFrdPrp%
Dim MthLin
For Each MthLin In Itr(MthLinAyzMd(A))
    With MthNm3(MthLin)
        Select Case True
        Case .IsPub And .IsSub: NPubSub = NPubSub + 1
        Case .IsPub And .IsFun: NPubFun = NPubFun + 1
        Case .IsPub And .IsPrp: NPubPrp = NPubPrp + 1
        Case .IsPrv And .IsSub: NPrvSub = NPrvSub + 1
        Case .IsPrv And .IsFun: NPrvFun = NPrvFun + 1
        Case .IsPrv And .IsPrp: NPrvPrp = NPrvPrp + 1
        Case .IsFrd And .IsSub: NFrdSub = NFrdSub + 1
        Case .IsFrd And .IsFun: NFrdFun = NFrdFun + 1
        Case .IsFrd And .IsPrp: NFrdPrp = NFrdPrp + 1
        Case Else: Thw CSub, "Invalid MthNm3", "MthLin MthNm3", MthLin, .Lin
        End Select
    End With
Next
Set MthCnt = New MthCnt
MthCnt.Init MdNm(A), NPubSub, NPubFun, NPubPrp, NPrvSub, NPrvFun, NPrvPrp, NFrdSub, NFrdFun, NFrdPrp
End Function
Function MthCntMd() As MthCnt
Set MthCntMd = MthCnt(CurMd)
End Function
Sub MthCntPjBrw()
BrwDry DryzTLinAy(LyzMthCntAy(MthCntPj))
End Sub
Function MthCntPj() As MthCnt()
MthCntPj = MthCntAy(CurPj)
End Function
Function LyzMthCntAy(A() As MthCnt) As String()
Dim I
For Each I In Itr(A)
    PushI LyzMthCntAy, CvMthCnt(I).Lin
Next
End Function
Function CvMthCnt(A) As MthCnt
Set CvMthCnt = A
End Function
Function MthCntAy(A As VBProject) As MthCnt()
If A.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In A.VBComponents
    PushObj MthCntAy, MthCnt(C.CodeModule)
Next
End Function

