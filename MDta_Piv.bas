Attribute VB_Name = "MDta_Piv"
Option Explicit
Const CMod$ = "MDta__Piv."
Public Type KKCntMulItmColDry
    NKK As Integer
    Dry() As Variant
End Type
    
Public Enum eAgg
    eSum
    eCnt
    eAvg
End Enum
Function DryGpAy(A, KIx%, GIx%) As Variant()
If Si(A) = 0 Then Exit Function
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
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
DryGpAy = O
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
Private Function KKDrToItmAyDualColDry(Dry(), KKColIx%(), ItmColIx%) As Variant()
Dim Dr, Ix&, KKDr(), Itm
Dim O() 'KKDr_To_ItmAy_DualColDry
For Each Dr In Itr(Dry)
    KKDr = AywIxAy(Dry, KKColIx)
    Itm = Dr(ItmColIx)
    Ix = KKDrIx(KKDr, O)
    If Ix = -1 Then
        PushI O, Array(KKDr, Array(Itm))
    Else
        O(Ix)(1) = AyAdd(O(Ix)(1), Itm)
    End If
Next
KKDrToItmAyDualColDry = O
End Function
Function KKCntMulItmColDry(Dry(), KKColIx%(), ItmColIx%) As Variant()
Dim A(): A = KKDrToItmAyDualColDry(Dry, KKColIx, ItmColIx)
KKCntMulItmColDry = KKCntMulItmColDryD(A)
End Function
Private Function KKCntMulItmColDryD(KKDrToItmAyDualColDry()) As Variant()

End Function
Function GpDicDKG(A As Drs, KK, G$) As Dictionary
Dim Fny$()
Dim KeyIxAy&(), GIx%
    Fny = NyzNN(KK)
    KeyIxAy = IxAy(A.Fny, Fny)
    PushI Fny, G & "_Gp"
    GIx = IxzAy(Fny, G)
Set GpDicDKG = DryGpDic(A.Dry, KeyIxAy, GIx)
End Function

Function DryDotAy(DotAy) As Variant()
Dim I
For Each I In Itr(DotAy)
    PushI DryDotAy, SplitDot(I)
Next
End Function

Function DryzLyWithColon(LyWithColon$()) As Variant()
Dim I
For Each I In Itr(LyWithColon)
    PushI DryzLyWithColon, SplitColon(I)
Next
End Function

Function DryGpDic(A, KeyIxAy, G) As Dictionary
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
Set DrszFbt = DrszT(Db(Fb), T)
End Function
Function KE24Drs() As Drs
Set KE24Drs = DrszFbt(SampFbzDutyDta, "KE24")
End Function

