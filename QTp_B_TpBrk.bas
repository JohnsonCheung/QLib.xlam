Attribute VB_Name = "QTp_B_TpBrk"
Option Compare Text
Option Explicit
Private Const CMod$ = "MTp_TpBrk."
Private Const Asm$ = "QTp"
Type Blks
    D As Drs ' BlkTy SepLin Dy:[Lno Lin]
End Type
Type Blk
    DroBlk() As Variant ' BlkTy SepLin Dyo<Lno Lin>  BlkIx
End Type
Type DyoLLin
    Dy() As Variant ' Av<Lno Lin>
End Type
Private Sub Z_ErBrks()
Dim Act As Blks
Dim Ept As Blks
Dim Tp$, BlkTyss$
GoSub T0
Exit Sub
T0:
    Tp = SampSqTp
    BlkTyss = "PM SQ SW RM"
End Sub

Function BlkszTp(Tp$, BlkTyAy$(), Optional SepLinPfx$ = "== ") As Blks
If Trim(Tp) = "" Then Exit Function
Dim ODy() ' Dyo<BlkTy SepLin Dyo<Lno Lin>>
Dim Blk(): Blk = NewBlk ' Av<BlkTy SepLin Dyo<Lno Lin>>
Dim InBlk As Boolean, IsBrk As Boolean
Dim L, J&: For Each L In Itr(RmvRmkzLy(SplitCrLf(Tp)))
    If HasPfx(L, SepLinPfx) Then
        PushI ODy, Blk
    Else
        If L <> "" Then PushI Blk(2), L
    End If
    J = J + 1
Next
BlkszTp = Blks(ODy)
End Function

Function ErBlks(Tp$, BlkTyAy$(), Optional SepLinPfx$ = "== ") As Blks
End Function

Function HasMajPfx(Ly$(), MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To UB(Ly)
    If HasPfx(Ly(J), MajPfx) Then Cnt = Cnt + 1
Next
HasMajPfx = Cnt > (Si(Ly) \ 2)
End Function

Function Blks(DyoBlk()) As Blks
'Fm DyoBlk: Dyo<BlkTy SepLin Dyo<Lno Lin>>
Blks.D = DrszFF("BlkTy SepLin DyoLLin", DyoBlk)
End Function

Function NewBlk() As Variant()
Dim O(): ReDim O(2): O(2) = EmpAv
NewBlk = O
End Function

Private Sub Z_Blks()
Dim Tp$, SepLinPfx$, BlkTyAy$()
GoSub Z
Exit Sub
Z:
    Tp = JnCrLf(SrczP(CPj))
    SepLinPfx = "Function"
    Dim B As Blks: B = BlkszTp(Tp, BlkTyAy, SepLinPfx)
    Vc FmtBlks(B)
    Return
End Sub
Sub BrwBlks(B As Blks)
Brw FmtBlks(B)
End Sub

Sub VcBlks(B As Blks)
Vc FmtBlks(B)
End Sub

Function CntBlk%(B As Blks, BlkTy$)
Dim O&
Dim DroBlk: For Each DroBlk In B.D.Dy
    If DroBlk(0) = BlkTy Then O = O + 1
Next
CntBlk = O
End Function
Function LyzBlk(B As Blk) As String()
LyzBlk = B.DroBlk(2)
End Function
Function LyzBlkzTy(B As Blks, BlkTy$) As String()
LyzBlkzTy = LyzBlk(BlkzTy(B, BlkTy))
End Function

Function FmtBlks(B As Blks) As String()
Dim DroBlk, J&: For Each DroBlk In B.D.Dy
    PushIAy FmtBlks, FmtBlk(Blk(DroBlk), J)
Next
End Function
Function Blk(DroBlk) As Blk
'Fm DroBlk: Av<BlkTy SepLin Dyo<Lno Lin> BlkIx>
If Si(DroBlk) <> 4 Then Stop
Blk.DroBlk = DroBlk
End Function
Function BlkTy$(B As Blk)
BlkTy = B.DroBlk(0)
End Function

Private Function BlkIx$(B As Blk)
BlkIx = B.DroBlk(3)
End Function

Function FmtLLin(DyoLLin()) As String()

End Function

Function SepLin$(B As Blk)
SepLin = B.DroBlk(1)
End Function

Function DyoLLinzBlk(B As Blk) As Variant()
DyoLLinzBlk = B.DroBlk(2)
End Function

Function FmtBlk(B As Blk, Optional BlkIx = -1) As String()
'Fm Blk : Dro<BlkTy SepLin Dyo<L Lin>>
Dim P$: If BlkIx >= 0 Then P = "(BlkIx:" & BlkIx & ") "
PushI FmtBlk, P & Qte(BlkTy(B), "BlkTy(*)")
PushI FmtBlk, SepLin(B)
PushIAy FmtBlk, FmtLLin(DyoLLinzBlk(B))
End Function

Function BlkzTy(B As Blks, BlkTy$, Optional FmIx = 0) As Blk
Dim Dy(): Dy = B.D.Dy
Dim Dr$(), Ix%: For Ix = FmIx To UB(Dy) ' Dro<BlkTy SepLin Dyo<Lno Lin> BlkIx>
    Dr = Dy(Ix)
    If Dr(0) = BlkTy Then
        BlkzTy = Blk(Dr): Exit Function
    End If
Next
End Function

Function LyzBlkTy(B As Blks, BlkTy$) As String()
LyzBlkTy = StrColzDySnd(DyoLLinzBlk(BlkzTy(B, BlkTy)))
End Function

Function LyAyzBlkTy(B As Blks, BlkTy$, Optional FmIx% = 0) As Variant()
Dim D As Drs: D = DwEq(B.D, "BlkTy", BlkTy)
Dim Dr: For Each Dr In D.Dy
    Dim DyoBlk(): DyoBlk = Dr(2)  ' Av<BlkTy SepLin Dyo<Lno Lin>>
    PushI LyAyzBlkTy, StrColzDySnd(DyoBlk)
Next
End Function

Function IsBlkEmp(B As Blk) As Boolean
IsBlkEmp = Si(B.DroBlk) = 0
End Function
Function BlkzNeTy(B As Blks, BlkTyAy$(), Optional FmIx = 0) As Blk

End Function
Function ErzBlkTy(B As Blks, BlkTy$, Optional FmIx = 0) As String()
'Ret: Those Blk @FmIx and eq @BlkTy in @B are considered as error.  Rpt them as :ErLy @@
Dim Dy(): Dy = B.D.Dy
Dim Ix%: For Ix = FmIx To UB(Dy)
    
Next
End Function

Function ErzErBlk(B As Blks, BlkTyss$) As String()
'Ret : Any blk having blk ty not in @BlkTyss are ErBlk, rpt them as :ErLy
If NoReczDrs(B.D) Then Exit Function
Dim Blk As Blk: Blk = BlkzTy(B, "Er")
If IsBlkEmp(Blk) Then Exit Function
PushI ErzErBlk, "Unexpected Blk, valid block is PM SW SQ RM"
While Not IsBlkEmp(Blk)
    PushIAy ErzErBlk, ErzBlkAft(B, Blk)
    Blk = NxtBlk(B, Blk)
Wend
End Function

Function NxtBlk(B As Blks, M As Blk) As Blk
'Ret : Nxt :Blk of @M fm @B.  M.BlkIx
NxtBlk = BlkzTy(B, BlkTy(M), BlkIx(M) + 1)
End Function

Function ErzAftBlk(B As Blks, Aft As Blk) As String()
'Ret : all :Blk @Aft in @B are considered as er, rpt them as :ErLy
Stop
End Function

