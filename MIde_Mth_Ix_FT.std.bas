Attribute VB_Name = "MIde_Mth_Ix_FT"
Option Explicit

Function MthFTIxAyMth(Src$(), MthNm, Optional WithTopRmk As Boolean) As FTIx()
Dim FmIx&, ToIx&, Ix
For Each Ix In Itr(MthIxAyMth(Src, MthNm))
    If WithTopRmk Then
        FmIx = Ix
    Else
        FmIx = MthTopRmkIx(Src, Ix)
    End If
   PushObj MthFTIxAyMth, FTIx(FmIx, MthToIx(Src, FmIx))
Next
End Function

Function MthFTIxAyMdMth(A As CodeModule, MthNm, Optional WithTopRmk As Boolean) As FTIx()
MthFTIxAyMdMth = MthFTIxAyMth(Src(A), MthNm, WithTopRmk)
End Function


Function MthFTIxAy(Src$(), Optional WithTopRmk As Boolean) As FTIx()
Dim Ix, FmIx&, ToIx&
For Each Ix In MthIxItr(Src)
    If WithTopRmk Then
        FmIx = Ix
    Else
        FmIx = MthTopRmkIx(Src, Ix)
    End If
    PushObj MthFTIxAy, FTIx(FmIx, ToIx)
Next
End Function
