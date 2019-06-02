Attribute VB_Name = "QIde_Mth_Ix_FT"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Ix_FT."
Private Const Asm$ = "QIde"

Function MthFeiszSN(Src$(), Mthn, Optional WiTopRmk As Boolean) As Feis
Dim FmIx&, EIx&, Ix&, I
For Each I In Itr(MthIxyzSN(Src, Mthn))
    Ix = I
    If WiTopRmk Then
        FmIx = Ix
    Else
        FmIx = TopRmkIx(Src, Ix)
    End If
   PushFei MthFeiszSN, Fei(FmIx, EndLix(Src, FmIx))
Next
End Function

Function MthFeiszMN(A As CodeModule, Mthn, Optional WiTopRmk As Boolean) As Feis
MthFeiszMN = MthFeiszSN(Src(A), Mthn, WiTopRmk)
End Function

Function MthFeis(Src$(), Optional WiTopRmk As Boolean) As Feis
Dim Ix&, FmIx&, EIx&, I
For Each I In MthIxItr(Src)
    Ix = I
    If WiTopRmk Then
        FmIx = Ix
    Else
        FmIx = TopRmkIx(Src, Ix)
    End If
    PushFei MthFeis, Fei(FmIx, EIx)
Next
End Function
