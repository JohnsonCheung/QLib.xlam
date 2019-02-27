Attribute VB_Name = "MTp_SqyRslt_1_BlkAy"
Option Explicit
Function BlkAy(SqTp$) As Blk()
Dim Ly$():        Ly = SplitCrLf(SqTp)
Dim G() As Gp:     G = GpAy(Ly)
Dim G1() As Gp:   G1 = GpAyzRmvRmk(G)
BlkAy = BlkAyzGpAy(G1)
End Function

Private Function BlkAyzGpAy(A() As Gp) As Blk()
Dim I
For Each I In Itr(A)
    PushObj BlkAyzGpAy, Blk(CvGp(I))
Next
End Function

Private Function BlkTyzGp$(A As Gp)
Dim Ly$(): Ly = LyzGp(A)
BlkTyzGp = BlkTy(Ly)
End Function

Private Function GpAy(Ly$()) As Gp()
Dim O() As Gp, J&, LnxAy() As Lnx, M As Lnx
For J = 0 To UB(Ly)
    Dim Lin$
    Lin = Ly(J)
    If HasPfx(Lin, "==") Then
        If Sz(LnxAy) > 0 Then
            PushObj GpAy, Gp(LnxAy)
        End If
        Erase LnxAy
    Else
        PushObj LnxAy, Lnx(J, Lin)
    End If
Next
If Sz(LnxAy) > 0 Then
    PushObj GpAy, Gp(LnxAy)
End If
GpAy = O
End Function

Private Function GpAyzRmvRmk(A() As Gp) As Gp()
Dim J%, O() As Gp, M As Gp
For J = 0 To UB(A)
    Set M = RmvRmkzGp(A(J))
    If Sz(M.LnxAy) > 0 Then
        PushObj O, M
    End If
Next
GpAyzRmvRmk = O
End Function

Private Function RmvRmkzGp(A As Gp) As Gp
Dim B() As Lnx: B = A.LnxAy
Dim M As Lnx
Dim J&, O() As Lnx
For J = 0 To UB(B)
    M = B(J)
    If Not IsDDRmkLin(M.Lin) Then
        PushObj O, M
    End If
Next
Set RmvRmkzGp = Gp(O)
End Function

Private Function Blk(A As Gp) As Blk
Set Blk = New Blk
With Blk
    .BlkTy = BlkTyzGp(A)
    Set .Gp = A
End With
End Function

Private Function BlkTy$(Ly$())
Dim O$
Select Case True
Case IsPmLy(Ly): O = "PM"
Case IsSwLy(Ly): O = "SW"
Case IsRmLy(Ly): O = "RM"
Case IsSqLy(Ly): O = "SQ"
Case Else: O = "ER"
End Select
BlkTy = O
End Function

Private Function IsPmLy(A$()) As Boolean
IsPmLy = HasMajPfx(A, "%")
End Function

Private Function IsRmLy(A$()) As Boolean
IsRmLy = Sz(A) = 0
End Function

Private Function IsSqLy(A$()) As Boolean
If Sz(A) <> 0 Then Exit Function
Dim L$: L = A(0)
Dim Sy$(): Sy = SySsl("?SEL SEL ?SELDIS SELDIS UPD DRP")
If HitPfxAy(L, Sy) Then IsSqLy = True: Exit Function
End Function

Private Function IsSwLy(Ly$()) As Boolean
IsSwLy = HasMajPfx(Ly, "?")
End Function


Private Sub ZZ()
Dim A() As Blk
Dim B$
Dim C As Gp
Dim D() As Gp
Dim E() As Lnx
Dim F%()
Dim G$()
Dim XX
ErzBlkAy A
End Sub


