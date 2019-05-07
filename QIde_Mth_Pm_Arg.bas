Attribute VB_Name = "QIde_Mth_Pm_Arg"
Option Explicit
Private Const CMod$ = "MIde_Mth_Pm_Arg."
Private Const Asm$ = "QIde"
Public Const DoczArgStr$ = "It is splitting of MthPm"
Public Const DoczArgSy$ = "It Array of ArgStr"
Public Const DoczSset$ = "String-Aset"
Function MthPm$(MthLin$)
MthPm = BetBktMust(MthLin, CSub)
End Function


Property Get ArgAsetInPj() As Aset
Set ArgAsetInPj = ArgAsetzPj(CurPj)
End Property

Function ArgAsetzPj(A As VBProject) As Aset
Set ArgAsetzPj = New Aset
Dim L$, I
For Each I In MthLinSyzPj(A)
    L = I
    ArgAsetzPj.PushAy ArgSy(L)
Next
End Function

Private Sub Z_ArgAsetInPj()
ArgAsetInPj.Srt.Vc
End Sub

Function DimItmzArg$(Arg$)
DimItmzArg = BefOrAll(RmvPfxSpc(RmvPfxSpc(Arg, "Optional"), "ParamArray"), " =")
End Function

Function SfxzArg$(Arg$)
SfxzArg = RmvNm(DimItmzArg(Arg))
End Function


