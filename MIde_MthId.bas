Attribute VB_Name = "MIde_MthId"
Option Explicit
Public Const DocOfMthQLin$ = "newtype A3DotLin.  Q for qualified.  Fmt is PjNm.ShtMdTy.MdNm.MthLin"
Public Const DocOfMthId$ = "It is 1-to-3-Dig-Zero-Pad-Integer-Str starting from 1 for each mth within a module.  The sorting is Mdy.Kd.Nm."
Public Const DocOfMthQidLin$ = "newtype A6DotLin.  Q is for qualify.  It is a variant of MthLin of format Pj.MdTy.MdNm.MthId.ShtMthMdy.ShtMthKd.MthRst"
Public Const DocOfMthMLin$ = "M for Modified.  Fmt is [ShtMthMdy.ShtMthKd.MthNmRst].  MthNmRst is MthMLin with MthMdy and MthTy removed."
Public Const DocOfMthSQMLin$ = "It is A5DotStr.  Q for qualified.  M for Modified.  Fmt is [MthSrtKey.PjNm.ShtMdTyNm.MdNm.ShtMthMdy.ShtMthKd.MthNmRst]."
Public Const DocOfMthSrtKey$ = "It is Str.  Fmt is [MthMdy:MthNm]"
Private A_MthQLy$()
Function FmtMthQidLyOfVbe() As String()
FmtMthQidLyOfVbe = DotLyInsSep(MthQidLy(MthQLyOfVbe), 3)
End Function
Private Sub Z_FmtMthQidLyOfVbe()
Vc FmtMthQidLyOfVbe
End Sub
Function MthRetNmRstLinzMthNmRstLin$(MthNmRstLin$, IsRetVal As Boolean)
Dim Pm$, TyChr$, RetTy$, Rmk$, MthNm$

Dim Ret$
    Ret = ShtRetTy(TyChr, RetTy, IsRetVal)
MthRetNmRstLinzMthNmRstLin = JnDotAp(Ret, MthNm, FmtPm(Pm), Rmk)
End Function
Function MthQidLyOfVbe(Optional WhStr$) As String()
MthQidLyOfVbe = MthQidLy(MthQLyOfVbe(WhStr))
End Function

Private Sub Z_MthQidLyOfVbe()
Vc MthQidLyOfVbe
End Sub

Function MthSrtKey$(ShtMthMdy$, MthNm$)
MthSrtKey = ShtMthMdy & ":" & MthNm
End Function
Private Function DicOf_PjMdTyMdNm_To_MthQLy() As Dictionary
Dim K$, MthQLin, O As New Dictionary, MthQLy$()
For Each MthQLin In A_MthQLy
    Dim Ay$(): Ay = SplitDot(MthQLin)
    ReDim Preserve Ay(2)
    K = JnDot(Ay)
    If O.Exists(K) Then
        MthQLy = O(K)
        PushI MthQLy, MthQLin
        O(K) = MthQLy
    Else
        O.Add K, Sy(MthQLin)
    End If
Next
Set DicOf_PjMdTyMdNm_To_MthQLy = O
End Function

Private Function MthSrtKeyzLin$(MthLin) ' MthKey is Mdy.Nm
With MthNm3(MthLin)
MthSrtKeyzLin = .ShtMdy & "." & .Nm
End With
End Function

Private Function MthQidLy(MthQLy$()) As String()
If Si(MthQLy) = 0 Then Exit Function
A_MthQLy = MthQLy
Dim I
For Each I In DicOf_PjMdTyMdNm_To_MthQLy.Items
    PushIAy MthQidLy, MthQidLyzSamMdMthQLy(CvSy(I))
Next
End Function
Function DotLinRmvSegN$(DotLin, Optional SegN% = 1)
Dim Ay$(): Ay = SplitDot(DotLin)
Dim Ay1$(): Ay1 = AyeEleAt(Ay, SegN - 1)
DotLinRmvSegN = JnDot(Ay1)
End Function
Function FstNDotSeg$(DotLin$, Optional NSeg% = 1)
FstNDotSeg = JnDot(AywFstNEle(SplitDot(DotLin), NSeg))
End Function
Function DotLyInsSep(DotLy$(), Optional UpToNSeg% = 1, Optional Sfx$ = "------") As String()
Dim U&: U = UB(DotLy): If U = -1 Then Exit Function
Dim Las$, Cur$, J&
Las = FstNDotSeg(DotLy(0), UpToNSeg)
PushI DotLyInsSep, Las & Sfx
PushI DotLyInsSep, DotLy(0)
For J = 1 To U
    Cur = FstNDotSeg(DotLy(J), UpToNSeg)
    If Cur <> Las Then
        PushI DotLyInsSep, Cur & Sfx
        Las = Cur
    End If
    PushI DotLyInsSep, DotLy(J)
Next
End Function
Function DotLyRmvSegN(DotLy$(), Optional SegN% = 1) As String()
Dim DotLin
For Each DotLin In Itr(DotLy)
    PushI DotLyRmvSegN, DotLinRmvSegN(DotLin, SegN)
Next
End Function
Private Function MthQMLy(MthQLy$()) As String()
Dim O$(), MthQLin
For Each MthQLin In Itr(MthQLy)
    PushI O, MthSQMLin(MthQLin)
Next
MthQMLy = DotLyRmvSegN(CvSy(AyQSrt(O)))
End Function
Private Sub Z_MthSQMLin()
Dim MthQLin$
GoSub T1
Exit Sub
T1:
    MthQLin = "Pj.MdTy.Md.Sub AA() As AA.BB"
    Ept = "Pub:AA.Pj.MdTy.Md.Pub.Sub.AA() As AA.BB"
    GoTo Tst
Tst:
    Act = MthSQMLin(MthQLin)
    C
    Return
End Sub
Private Function MthSQMLin$(MthQLin)
Dim L$: L = MthQLin
Dim ShtMthMdy$, ShtMthTy$, MthNm$, PjNm$, ShtMdTy$, MdNm$, MthNmRst$, Key$
PjNm = ShfDotSeg(L)
ShtMdTy = ShfDotSeg(L)
MdNm = ShfDotSeg(L)
ShtMthMdy = ShfShtMthMdy(L)
ShtMthTy = ShfShtMthTy(L)
MthNm = ShfNm(L)
MthNmRst = MthNm & L
Key = MthSrtKey(ShtMthMdy, MthNm)
MthSQMLin = JnDotAp(Key, PjNm, ShtMdTy, MdNm, ShtMthMdy, ShtMthTy, MthNmRst)
End Function
Private Sub Asg_ShtMthMdy_ShtMthTy_MthNm_MthNmRst(OShtMthMdy$, OShtMthTy$, OMthNm$, OMthNmRst$, MthLin$)
Dim L$: L = MthLin
OShtMthMdy = ShfShtMthMdy(L)
OShtMthTy = ShfShtMthTy(L)
OMthNm = ShfNm(L)
OMthNmRst = OMthNm & L
End Sub
Private Function MthQidLin$(MthQMLin, Id$)
Dim Ay$(): Ay = SplitDot(MthQMLin): If Si(Ay) < 6 Then Thw CSub, "MtQMLin should have at least 5 dots", "MthQMLin", MthQMLin
Dim Ay1$(): Ay1 = AyInsItm(Ay, Id, 3)
MthQidLin = JnDot(Ay1)
End Function

Private Function MthQidLyzSamMdMthQLy(SamMdMthQLy$()) As String() 'Assume the MthQLy are from same module
Dim N%, J&, MthQMLin, Id$
N = NDig(Si(SamMdMthQLy))
J = 0
'Brw SamMdMthQLy
'Stop
For Each MthQMLin In Itr(MthQMLy(SamMdMthQLy))
    J = J + 1
    Id = Pad0(J, N)
    PushI MthQidLyzSamMdMthQLy, MthQidLin(MthQMLin, Id)
Next
End Function

