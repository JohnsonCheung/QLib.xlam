Attribute VB_Name = "QTp_Tp_Blk"
Option Compare Text
Option Explicit
Private Const CMod$ = "MTp_SqyRslt_1_Blks."
Private Const Asm$ = "QTp"
Type Blk: BlkTy As String: SepLin As String: Lnxs As Lnxs: End Type
Type Blks: N As Long: Ay() As Blk: End Type
Function Blks(Tp$, Optional SepLinPfx$ = "==") As Blks
If Trim(Tp) = "" Then Exit Function
Dim M As Blk, L
Dim InBlk As Boolean, J&, IsBrk As Boolean
For Each L In RmvRmkzLy(SplitCrLf(Tp))
    IsBrk = HasPfx(L, SepLinPfx)
    Select Case True
    Case IsBrk: PushBlk Blks, M: M = BlkzSepLin(L)
    Case Else:  If L <> "" Then PushLnx M.Lnxs, Lnx(L, J)
    End Select
    J = J + 1
Next
PushBlk Blks, M
End Function
Private Sub Z_Blks()
Dim Tp$, SepLinPfx$
GoSub ZZ
Exit Sub
ZZ:
    Tp = JnCrLf(SrczP(CPj))
    SepLinPfx = "Function"
    Vc FmtBlks(Blks(Tp, SepLinPfx))
    Return
End Sub
Sub BrwBlks(A As Blks)
B FmtBlks(A)
End Sub
Sub VcBlks(A As Blks)
Vc FmtBlks(A)
End Sub
Function IsBlkOfEmp(A As Blk) As Boolean
If A.BlkTy <> "" Then Exit Function
If A.SepLin <> "" Then Exit Function
'If Not IsLnxsOfEmp(A.Lnxs) Then Exit Function
IsBlkOfEmp = True
End Function
Function BlkzSepLin(SepLin) As Blk
BlkzSepLin.SepLin = SepLin
End Function
Function EmpBlk() As Blk
End Function
Function CntBlk%(A As Blks, BlkTy$)
Dim O&, J&
For J = 0 To A.N - 1
    If A.Ay(J).BlkTy = BlkTy Then O = O + 1
Next
CntBlk = O
End Function

Function FstBlkOrDie(A As Blks, BlkTy$) As Blk
'Blk.Lnxs
'If IsBlkOfEmp(A) Then Thw CSub, "BlkTy not found", "BlkTy Blks", BlkTy, FmtBlks(A)
End Function
Function FstLyOrDiezBlksTy(A As Blks, BlkTy$) As String()
'FstLyOrDiezBlksTy = LyzBlk(LyzBlk(FstBlkOrDie(A, Ty)))
End Function
Function LyzBlk(A As Blk) As String()
LyzBlk = LinAyzLnxs(A.Lnxs)
End Function
Function FstBlk(A As Blks, BlkTy$) As Blk
Dim J&
For J = 0 To A.N - 1
    If A.Ay(J).BlkTy = BlkTy Then FstBlk = A.Ay(J): Exit Function
Next
End Function
Function FmtBlks(A As Blks) As String()
Dim J&
For J = 0 To A.N - 1
    PushIAy FmtBlks, FmtBlk(A.Ay(J), J)
Next
End Function

Function FmtBlk(A As Blk, Optional BlkIx = -1) As String()
Dim P$: If BlkIx >= 0 Then P = "(BlkIx:" & BlkIx & ") "
PushI FmtBlk, P & Quote(A.BlkTy, "BlkTy(*)")
PushI FmtBlk, A.SepLin
PushIAy FmtBlk, FmtLnxs(A.Lnxs)
End Function
Function Blk(BlkTy$, Lnxs As Lnxs) As Blk
Blk.BlkTy = BlkTy
Blk.Lnxs = Lnxs
End Function

Private Function BlkzLy(Ly$()) As Blk
Dim J&, OLnxs As Lnxs, OBlkTy$
For J = 0 To UB(Ly)
    Dim Lin$
    Lin = Ly(J)
    If HasPfx(Lin, "==") Then
        If OLnxs.N > 0 Then
'            PushObj Blk, Gp(Lnxs)
        End If
'        Erase Lnxs
    Else
        PushLnx OLnxs, Lnx(Lin, J)
    End If
Next
BlkzLy = Blk(OBlkTy, OLnxs)
End Function
Function LnxszBlksTy(A As Blks, BlkTy$) As Lnxs
Dim J&
For J = 0 To A.N - 1
'    If A.Ay(J).BlkTy = BlkTy Then LnxszBlksTy = A.Ay(J): Exit Function
Next
End Function
Function BlkswTy(A As Blks, BlkTy$) As Blks
Dim J%
For J = 0 To A.N - 1
    If A.Ay(J).BlkTy = BlkTy Then
        PushBlk BlkswTy, A.Ay(J)
    End If
Next
End Function
Sub PushBlk(O As Blks, M As Blk)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function LyzBlksTy(A As Blks, BlkTy$) As String()
LyzBlksTy = LinAyzLnxs(FstBlk(A, BlkTy).Lnxs)
End Function

Function LyAyzBlksTy(A As Blks, BlkTy$) As Variant()
With BlkswTy(A, BlkTy)
    Dim J&
    For J = 0 To .N - 1
        PushI LyAyzBlksTy, LinAyzLnxs(.Ay(J).Lnxs)
    Next
End With
End Function

Function ErzBlks(A As Blks) As String()
Dim Blk As Blk: Blk = FstBlk(A, "Er")
'If Si(Blk) > 0 Then
    PushIAy ErzBlks, ErzBlk(Blk, "Unexpected Blk, valid block is PM SW SQ RM")
'End If
End Function


