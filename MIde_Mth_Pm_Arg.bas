Attribute VB_Name = "MIde_Mth_Pm_Arg"
Option Explicit
Public Const DocOfArgStr$ = "It is splitting of MthPm"
Public Const DocOfArgSy$ = "It Array of ArgStr"
Public Const DocOfSset$ = "String-Aset"
Function MthPm$(MthLin)
MthPm = BetBktMust(MthLin, CSub)
End Function


Property Get ArgAsetOfPj() As Aset
Set ArgAsetOfPj = ArgAsetzPj(CurPj)
End Property

Function ArgAsetzPj(A As VBProject) As Aset
Set ArgAsetzPj = New Aset
Dim L
For Each L In MthLinAyzPj(A)
    ArgAsetzPj.PushAy ArgAy(L)
Next
End Function

Private Sub Z_ArgAsetOfPj()
ArgAsetOfPj.Srt.Vc
End Sub

Function DimItmzArg$(Arg)
DimItmzArg = StrBefOrAll(RmvPfxSpc(RmvPfxSpc(Arg, "Optional"), "ParamArray"), " =")
End Function

Function ArgSfx$(Arg)
ArgSfx = RmvNm(DimItmzArg(Arg))
End Function


