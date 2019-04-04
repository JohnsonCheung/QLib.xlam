Attribute VB_Name = "MIde_Mth_Ix_FT"
Option Explicit

Function MthFTIxAyzSrcMth(Src$(), MthNm, Optional WiTopRmk As Boolean) As FTIx()
Dim FmIx&, ToIx&, Ix
For Each Ix In Itr(MthIxAyzNm(Src, MthNm))
    If WiTopRmk Then
        FmIx = Ix
    Else
        FmIx = MthTopRmkIx(Src, Ix)
    End If
   PushObj MthFTIxAyzSrcMth, FTIx(FmIx, MthToIx(Src, FmIx))
Next
End Function

Function MthFTIxAyzMth(A As CodeModule, MthNm, Optional WiTopRmk As Boolean) As FTIx()
MthFTIxAyzMth = MthFTIxAyzSrcMth(Src(A), MthNm, WiTopRmk)
End Function

Function MthFTIxAy(Src$(), Optional WiTopRmk As Boolean) As FTIx()
Dim Ix, FmIx&, ToIx&
For Each Ix In MthIxItr(Src)
    If WiTopRmk Then
        FmIx = Ix
    Else
        FmIx = MthTopRmkIx(Src, Ix)
    End If
    PushObj MthFTIxAy, FTIx(FmIx, ToIx)
Next
End Function
