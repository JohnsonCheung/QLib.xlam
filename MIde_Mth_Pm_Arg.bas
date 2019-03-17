Attribute VB_Name = "MIde_Mth_Pm_Arg"
Option Explicit
Const DoczArgStr$ = "It is splitting of MthPm"
Const DoczArgSy$ = "It Array of ArgStr"
Const DoczSset$ = "String-Aset"
Function MthPm$(Lin)
If IsMthLin(Lin) Then MthPm = StrBetBkt(Lin)
End Function

Function MthArgSy(Lin) As String()
MthArgSy = SplitCommaSpc(MthPm(Lin))
End Function

Property Get ArgStrSetPj() As Aset
Set ArgStrSetPj = ArgStrSetzPj(CurPj)
End Property
Function ArgStrSetzPj(A As VBProject) As Aset
Set ArgStrSetzPj = New Aset
Dim L
For Each L In MthLinAyzPj(A)
    ArgStrSetzPj.PushAy MthArgSy(L)
Next
End Function

Private Sub Z_ArgStrSetPj()
ArgStrSetPj.Srt.Brw
End Sub

Function DimItmzArgStr$(ArgStr)
DimItmzArgStr = StrBefOrAll(RmvPfxSpc(RmvPfxSpc(ArgStr, "Optional"), "ParamArray"), " =")
End Function

Function ArgTy$(ArgStr)
ArgTy = RmvNm(DimItmzArgStr(ArgStr))
End Function


