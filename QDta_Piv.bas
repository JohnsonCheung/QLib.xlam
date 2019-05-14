Attribute VB_Name = "QDta_Piv"
Option Explicit
Private Const Asm$ = "QDta"
Private Const CMod$ = "MDta_Piv."
Public Type KKCntMulItmColDry
    NKK As Integer
    Dry() As Variant
End Type
    
Enum EmAgg
    EiSum
    EiCnt
    EiAvg
End Enum
Function DryBlk(A, KIx%, GIx%) As Variant()
If Si(A) = 0 Then Exit Function
Dim J%, O, K, Blk(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In A
    K = Dr(KIx)
    Gp = Dr(GIx)
    O_Ix = IxzAy(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryBlk = O
End Function
Private Function KKDrIx&(KKDr, FstColIsKKDrDry)
Dim Ix&, CurKKDr
For Each CurKKDr In Itr(FstColIsKKDrDry)
    If IsEqAy(KKDr, CurKKDr) Then
        KKDrIx = Ix
        Exit Function
    End If
    Ix = Ix + 1
Next
Ix = -1
End Function
Private Function KKDrToItmAyDualColDry(Dry(), KKColIx&(), ItmColIx&) As Variant()
Dim Dr, Ix&, KKDr(), Itm
Dim O() 'KKDr_To_ItmAy_DualColDry
For Each Dr In Itr(Dry)
    KKDr = AywIxy(Dry, KKColIx)
    Itm = Dr(ItmColIx)
    Ix = KKDrIx(KKDr, O)
    If Ix = -1 Then
        PushI O, Array(KKDr, Array(Itm))
    Else
        O(Ix)(1) = AddAy(O(Ix)(1), Itm)
    End If
Next
KKDrToItmAyDualColDry = O
End Function
Function KKCntMulItmColDry(Dry(), KKColIx&(), ItmColIx&) As Variant()
Dim A(): A = KKDrToItmAyDualColDry(Dry, KKColIx, ItmColIx)
KKCntMulItmColDry = KKCntMulItmColDryD(A)
End Function
Private Function KKCntMulItmColDryD(KKDrToItmAyDualColDry()) As Variant()

End Function
Function GpDic(A As Drs, KK$, G$) As Dictionary
Dim Fny$()
Dim KeyIxy&(), GIx%
    Fny = TermAy(KK)
    KeyIxy = Ixy(A.Fny, Fny)
    PushI Fny, G & "_Gp"
    GIx = IxzAy(Fny, G)
Set GpDic = DryGpDic(A.Dry, KeyIxy, GIx)
End Function
Function DryzDotLyz2Col(DotLy$()) As Variant()
Dim O(), I, S$
For Each I In Itr(DotLy)
    S = I
    With Brk1(S, ".")
       Push O, Sy(.S1, .S2)
   End With
Next
DryzDotLyz2Col = O
End Function

Function DryzDotLy(DotLy$()) As Variant()
Dim I
For Each I In Itr(DotLy)
    PushI DryzDotLy, SplitDot(I)
Next
End Function
Function DryzDotLyzTwoCol(DotLy$()) As Variant()
Dim I
For Each I In Itr(DotLy)
    With Brk1Dot(CStr(I))
    PushI DryzDotLyzTwoCol, Array(.S1, .S2)
    End With
Next
End Function

Function DryzLyWithColon(LyWithColon$()) As Variant()
Dim I
For Each I In Itr(LyWithColon)
    PushI DryzLyWithColon, SplitColon(CStr(I))
Next
End Function

Function DryGpDic(A, KeyIxy, G) As Dictionary
Const CSub$ = CMod & "DryGpDic"
'If K < 0 Or G < 0 Then
'    Thw CSub, "K-Idx and G-Idx should both >= 0", "K-Idx G-Idx", K, G
'End If
'Dim Dr, U&, O As New Dictionary, KK, GG, Ay()
'U = UB(A): If U = -1 Then Exit Function
'For Each Dr In A
'    KK = Dr(K)
'    GG = Dr(G)
'    If O.Exists(KK) Then
'        Ay = O(KK)
'        PushI Ay, GG
'        O(KK) = Ay
'    Else
'        O.Add KK, Array(GG)
'    End If
'Next
'Set DryGpDic = O
End Function

Function DrszFbt(Fb, T) As Drs
DrszFbt = DrszT(Db(Fb), T)
End Function

Function KE24Drs() As Drs
KE24Drs = DrszFbt(SampFbzDutyDta, "KE24")
End Function

