Attribute VB_Name = "QIde_Mth_Ix_FT"
Option Explicit
Private Const CMod$ = "MIde_Mth_Ix_FT."
Private Const Asm$ = "QIde"

Function MthFTIxAyzSrcMth(Src$(), MthNm$, Optional WiTopRmk As Boolean) As FTIx()
Dim FmIx&, ToIx&, Ix&, I
For Each I In Itr(MthIxAyzNm(Src, MthNm))
    Ix = I
    If WiTopRmk Then
        FmIx = Ix
    Else
        FmIx = MthTopRmkIx(Src, Ix)
    End If
   PushObj MthFTIxAyzSrcMth, FTIx(FmIx, MthToIx(Src, FmIx))
Next
End Function

Function MthFTIxAyzMth(A As CodeModule, MthNm$, Optional WiTopRmk As Boolean) As FTIx()
MthFTIxAyzMth = MthFTIxAyzSrcMth(Src(A), MthNm, WiTopRmk)
End Function

Function MthFTIxAy(Src$(), Optional WiTopRmk As Boolean) As FTIx()
Dim Ix&, FmIx&, ToIx&, I
For Each I In MthIxItr(Src)
    Ix = I
    If WiTopRmk Then
        FmIx = Ix
    Else
        FmIx = MthTopRmkIx(Src, Ix)
    End If
    PushObj MthFTIxAy, FTIx(FmIx, ToIx)
Next
End Function
