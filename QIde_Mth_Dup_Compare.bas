Attribute VB_Name = "QIde_Mth_Dup_Compare"
Option Explicit
Private Const CMod$ = "MIde_Mth_Dup_Compare."
Private Const Asm$ = "QIde"
Sub CmpFun(FunNm$, Optional InclEqLines As Boolean)
D FmtCmpFun(FunNm, InclEqLines)
End Sub

Function FmtCmpFun(FunNm, Optional InclSam As Boolean) As String()
'Found all Fun with given name and compare within curVbe if it is same
'Note: Fun is any-Mdy Fun/Sub/Prp-in-Md
Dim O$()
Dim N$(): ' N = FunFNmAy(FunNm)
'DupMthFNy_ShwNotDupMsg N, FunNm
If Si(N) <= 1 Then Exit Function
FmtCmpFun = FmtCmprDic(N, InclSam:=InclSam)
End Function

Private Sub Z_FunCmp()
'FunCmp "Ffn_Dlt"
End Sub

Private Function FmtCmprDic(A, Optional OIx% = -1, Optional OSam% = -1, Optional InclSam As Boolean) As String()
'DupMthFNyGp is Variant/String()-of-MthFNm with all mth-nm is same
'MthFNm is Mthn in FNm-fmt
'          Mth is Prp/Sub/Fun in Md-or-Cls
'          FNm-fmt which is 'Nm:Pj.Md'
'DupMthFNm is 2 or more MthFNy with same Mthn
'Ass DupMthFNyGp_IsVdt(A)
Dim J%, I%
Dim LinesAy$()
Dim UniqLinesAy$()
'    LinesAy = SyzMapAy(A, "FunFNm_MthLines")
    UniqLinesAy = AywDist(LinesAy)
Dim Mthn: Mthn = Brk(A(0), ":").S1
Dim Hdr$(): Hdr = FmtCmprDic__1Hdr(OIx, Mthn, Si(A))
Dim Sam$(): Sam = FmtCmprDic__2Sam(InclSam, OSam, A, LinesAy)
Dim Syn$(): Syn = FmtCmprDic__3Syn(UniqLinesAy, LinesAy, A)
Dim Cmp$(): Cmp = FmtCmprDic__4Cmp(UniqLinesAy, LinesAy, A)
FmtCmprDic = AddAyAp(Hdr, Sam, Syn, Cmp)
End Function

Private Function FmtCmprDic__1Hdr(OIx%, Mthn, Cnt%) As String()
Dim O$(1)
O(0) = "================================================================"
Dim A$
    If OIx >= 0 Then A = FmtQQ("#DupMthNo(?) ", OIx): OIx = OIx + 1
O(1) = A + FmtQQ("DupMthn(?) Cnt(?)", Mthn, Cnt)
FmtCmprDic__1Hdr = O
End Function

Private Function FmtCmprDic__2Sam(InclSam As Boolean, OSam%, DupMthFNyGp, LinesAy$()) As String()
If Not InclSam Then Exit Function
'{DupMthFNyGp} & {LinesAy} have same # of element
Dim O$()
Dim D$(): D = AywDup(LinesAy)
Dim J%, X$()
For J = 0 To UB(D)
    X = FmtCmprDic__2Sam1(OSam, D(J), DupMthFNyGp, LinesAy)
    PushAy O, X
Next
FmtCmprDic__2Sam = O
End Function

Private Function FmtCmprDic__2Sam1(OSam%, Lines$, DupMthFNyGp, LinesAy$()) As String()
Dim A1$()
    If OSam > 0 Then
        Push A1, FmtQQ("#Sam(?) ", OSam)
        OSam = OSam + 1
    End If
Dim A2$()
    Dim J%
    For J = 0 To UB(LinesAy)
        If JnCrLf(J) = Lines Then
            Push A2, "Shw """ & DupMthFNyGp(J) & """"
        End If
    Next
Dim A3$()
    A3 = BoxLyLines(Lines)
FmtCmprDic__2Sam1 = AddAyAp(A1, A2, A3)
End Function

Private Function FmtCmprDic__3Syn(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Si(UniqLinesAy) <= 1 Then Exit Function
Dim B$()
    Dim J%, I%
    Dim Lines
    For Each Lines In UniqLinesAy
        For I = 0 To UB(FunFNyGp)
            If Lines = JnCrLf(I) Then
                Push B, FunFNyGp(I)
                Exit For
            End If
        Next
    Next
'FmtCmprDic__3Syn = MapAyPXSy(B, "FmtQQ", "Sync_Fun ""?""")
End Function

Private Function FmtCmprDic__4Cmp(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Si(UniqLinesAy) <= 1 Then Exit Function
Dim L2$() ' = From L1 with each element with MdDNm added in front
    ReDim L2(UB(UniqLinesAy))
    Dim Fnd As Boolean, DNm$, J%, Lines$, I%
    For J = 0 To UB(UniqLinesAy)
        Lines = UniqLinesAy(J)
        Fnd = False
        For I = 0 To UB(LinesAy)
            If JnCrLf(I) = Lines Then
                DNm = FunFNyGp(I)
                L2(J) = DNm & vbCrLf & Dup("-", Len(DNm)) & vbCrLf & Lines
                Fnd = False
                GoTo Nxt
            End If
        Next
        Stop
Nxt:
    Next
'FmtCmprDic__4Cmp = SplitCrLfAyPad(L2)
End Function

Function MthnCmpFmt(A, Optional InclSam As Boolean) As String()
Dim N$(): ' N = Mthn_DupMthFNy(A)
If Si(N) > 1 Then
    MthnCmpFmt = FmtCmprDic(N, InclSam:=InclSam)
End If
End Function


Function VbeDupMthCmpLy(A As Vbe, B As WhPjMth, Optional InclSam As Boolean) As String()
Stop
Dim N$(): 'N = VbeDupMthFNm(A, B)
Dim Ay(): 'Ay = DupMthFNy_Blk(N)
Dim O$(), J%
'Push O, FmtQQ("Total ? dup function.  ? of them has mth-lines are same", Si(Ay), DupMthFNyBlkAllSameCnt(Ay))
Dim Cnt%, Sam%
For J = 0 To UB(Ay)
    PushAy O, FmtCmprDic(Ay(J), Cnt, Sam, InclSam:=InclSam)
Next
VbeDupMthCmpLy = O
End Function

Private Sub ZZ_VbeDupMthCmpLy()
'Brw VbeDupMthCmpLy(CVbe, WhEmpPjMth)
End Sub

Private Sub ZZ()
Z_FunCmp
MIde_Mth_Dup_Compare:
End Sub
