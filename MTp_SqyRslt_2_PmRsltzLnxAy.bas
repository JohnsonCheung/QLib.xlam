Attribute VB_Name = "MTp_SqyRslt_2_PmRsltzLnxAy"
Option Explicit
Type PmRslt: Er() As String: Pm As Dictionary: End Type
Type LnxAyRslt: Er() As String: LnxAy() As Lnx: End Type
Function LnxAyRsltzEr(LnxAy() As Lnx, Er$()) As LnxAyRslt
LnxAyRsltzEr.Er = Er
LnxAyRsltzEr.LnxAy = LnxAy
End Function

Private Function PmRsltzEr(Pm As Dictionary, Er$()) As PmRslt
PmRsltzEr.Er = Er
Set PmRsltzEr.Pm = Pm
End Function
Function PmRsltzLnxAy(A() As Lnx) As PmRslt
With PmLyRslt(A)
    PmRsltzLnxAy = PmRsltzEr(Dic(.Ly), .Er)
End With
End Function
Private Function PmLyRslt(A() As Lnx) As LyRslt
Dim R1 As LnxAyRslt
Dim R2 As LnxAyRslt

    R1 = LnxAyRsltzDupKey(A)
    R2 = LnxAyRsltzPercentagePfx(R1)
PmLyRslt = LyRslt(SyAddAp(R1.Er, R2.Er), LyzLnxAy(R2.LnxAy))
End Function

Private Function LnxAyRsltzDupKey(A() As Lnx) As LnxAyRslt

End Function
Private Function LnxAyRsltzPercentagePfx(A As LnxAyRslt) As LnxAyRslt

End Function

