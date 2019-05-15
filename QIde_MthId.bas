Attribute VB_Name = "QIde_MthId"
Option Explicit
Private Const CMod$ = "MIde_MthId."
Private Const Asm$ = "QIde"
Public Const DoczMthQLin = "newtype A3DotLin.  Q for qualified.  Fmt is Pjn.ShtMdTy.Mdn.MthLin"
Public Const DoczMthId$ = "It is 1-to-3-Dig-Zero-Pad-Integer-Str starting from 1 for each mth within a module.  The sorting is Mdy.Kd.Nm."
Public Const DoczMthQidLin = "newtype A6DotLin.  Q is for qualify.  It is a variant of MthLin of format Pj.MdTy.Mdn.MthId.ShtMthMdy.ShtMthKd.MthRst"
Public Const DoczMthMLin = "M for Modified.  Fmt is [ShtMthMdy.ShtMthKd.MthnRst].  MthnRst is MthMLin with MthMdy and MthTy removed."
Public Const DoczMthSQMLin = "It is A5DotStr.  Q for qualified.  M for Modified.  Fmt is [MthSrtKey.Pjn.ShtShtCmpTyzM.Mdn.ShtMthMdy.ShtMthKd.MthnRst]."
Public Const DoczMthSrtKey$ = "It is Str.  Fmt is [MthMdy:Mthn]"
Function FmtMthQidLyInVbe() As String()
FmtMthQidLyInVbe = DotLyInsSep(MthQidLy(MthQLyV), 3)
End Function
Private Sub Z_FmtMthQidLyInVbe()
Vc FmtMthQidLyInVbe
End Sub
Function MthRetNmRstLinzMthnRstLin(MthnRstLin, IsRetVal As Boolean)
Dim Pm$, TyChr$, RetTy$, Rmk$, Mthn

Dim Ret$
    Ret = ShtRetTy(TyChr, RetTy, IsRetVal)
MthRetNmRstLinzMthnRstLin = JnDotAp(Ret, Mthn, FmtPm(Pm), Rmk)
End Function
Function MthQidLyV() As String()
MthQidLyV = QSrt1(MthQidLy(MthQLyV))
End Function

Private Sub Z_MthQidLyInVbe()
'Vc MthQidLyInVbe
End Sub

Function MthSrtKey$(ShtMthMdy$, Mthn)
MthSrtKey = ShtMthMdy & ":" & Mthn
End Function
Private Function DicOf_PjMdTyMdn_To_MthQLy(MthQLy$()) As Dictionary
Dim K$, MthQLin, I, O As New Dictionary, S$()
For Each I In MthQLy
    MthQLin = I
    Dim Ay$(): Ay = SplitDot(MthQLin)
    ReDim Preserve Ay(2)
    K = JnDot(Ay)
    If O.Exists(K) Then
        S = O(K)
        PushI S, MthQLin
        O(K) = S
    Else
        O.Add K, Sy(MthQLin)
    End If
Next
Set DicOf_PjMdTyMdn_To_MthQLy = O
End Function

Private Function MthSrtKeyzLin(MthLin) ' MthKey is Mdy.Nm
With Mthn3zL(MthLin)
MthSrtKeyzLin = .ShtMdy & "." & .Nm
End With
End Function

Private Function MthQidLy(MthQLy$()) As String()
If Si(MthQLy) = 0 Then Exit Function
Dim I
For Each I In DicOf_PjMdTyMdn_To_MthQLy(MthQLy).Items
    PushIAy MthQidLy, MthQidLyzSamMdMthQLy(CvSy(I))
Next
End Function
Function RmvFstNSegzDotLin(DotLin, Optional FstNSeg% = 1)
Dim Ay$(): Ay = SplitDot(DotLin)
Dim Ay1$(): Ay1 = AyeEleAt(Ay, FstNSeg - 1)
RmvFstNSegzDotLin = JnDot(Ay1)
End Function
Function FstNDotSeg$(DotLin, Optional NSeg% = 1)
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
Dim DotLin, I
For Each I In Itr(DotLy)
    DotLin = I
    PushI DotLyRmvSegN, RmvFstNSegzDotLin(DotLin, SegN)
Next
End Function
Private Function MthQMLy(MthQLy$()) As String()
Dim O$(), MthQLin, I
For Each I In Itr(MthQLy)
    MthQLin = I
    PushI O, MthSQMLin(MthQLin)
Next
MthQMLy = DotLyRmvSegN(CvSy(QSrt1(O)))
End Function
Private Sub Z_MthSQMLin()
Dim MthQLin
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
Private Function MthSQMLin(MthQLin)
Dim L$: L = MthQLin
Dim ShtMthMdy$, ShtMthTy$, Mthn, Pjn$, ShtMdTy$, Mdn, MthnRst$, Key$
Pjn = ShfDotSeg(L)
ShtMdTy = ShfDotSeg(L)
Mdn = ShfDotSeg(L)
ShtMthMdy = ShfShtMdy(L)
ShtMthTy = ShfShtMthTy(L)
Mthn = ShfNm(L)
MthnRst = Mthn & L
Key = MthSrtKey(ShtMthMdy, Mthn)
MthSQMLin = JnDotAp(Key, Pjn, ShtMdTy, Mdn, ShtMthMdy, ShtMthTy, MthnRst)
End Function
Private Sub Asg_ShtMthMdy_ShtMthTy_Mthn_MthnRst(OShtMthMdy$, OShtMthTy$, OMthn, OMthnRst$, MthLin)
Dim L$: L = MthLin
OShtMthMdy = ShfShtMdy(L)
OShtMthTy = ShfShtMthTy(L)
OMthn = ShfNm(L)
OMthnRst = OMthn & L
End Sub
Private Function MthQidLin(MthQMLin, Id$)
Dim Ay$(): Ay = SplitDot(MthQMLin): If Si(Ay) < 6 Then Thw CSub, "MtQMLin should have at least 5 dots", "MthQMLin", MthQMLin
Dim Ay1$(): Ay1 = AyInsEle(Ay, Id, 3)
MthQidLin = JnDot(Ay1)
End Function

Private Function MthQidLyzSamMdMthQLy(SamMdMthQLy$()) As String() 'Assume the MthQLy are from same module
Dim N&, J&, I, MthQMLin, Id$
N = NDig(Si(SamMdMthQLy))
J = 0
'Brw SamMdMthQLy
'Stop
For Each I In Itr(MthQMLy(SamMdMthQLy))
    MthQMLin = I
    J = J + 1
    Id = Pad0(J, N)
    PushI MthQidLyzSamMdMthQLy, MthQidLin(MthQMLin, Id)
Next
End Function

