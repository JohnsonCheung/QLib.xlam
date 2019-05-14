Attribute VB_Name = "QIde_Mth_Ix_FT"
Option Explicit
Private Const CMod$ = "MIde_Mth_Ix_FT."
Private Const Asm$ = "QIde"

Function MthFEIxszSN(Src$(), Mthn, Optional WiTopRmk As Boolean) As FEIxs
Dim FmIx&, EIx&, Ix&, I
For Each I In Itr(MthIxyzSN(Src, Mthn))
    Ix = I
    If WiTopRmk Then
        FmIx = Ix
    Else
        FmIx = TopRmkIx(Src, Ix)
    End If
   PushFEIx MthFEIxszSN, FEIx(FmIx, MthEIx(Src, FmIx))
Next
End Function

Function MthFEIxszMN(A As CodeModule, Mthn, Optional WiTopRmk As Boolean) As FEIxs
MthFEIxszMN = MthFEIxszSN(Src(A), Mthn, WiTopRmk)
End Function

Function MthFEIxs(Src$(), Optional WiTopRmk As Boolean) As FEIxs
Dim Ix&, FmIx&, EIx&, I
For Each I In MthIxItr(Src)
    Ix = I
    If WiTopRmk Then
        FmIx = Ix
    Else
        FmIx = TopRmkIx(Src, Ix)
    End If
    PushFEIx MthFEIxs, FEIx(FmIx, EIx)
Next
End Function
