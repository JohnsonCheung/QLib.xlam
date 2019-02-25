Attribute VB_Name = "MIde_Mth_Pm_Arg"
Option Explicit
Function MthPm$(Lin)
If IsMthLin(Lin) Then
    MthPm = TakBetBkt(Lin)
End If
End Function
Function MthArgAy(Lin) As String()
MthArgAy = AyTrim(SplitComma(MthPm(Lin)))
End Function

Function ArgAyzPj(A As VBProject) As String()
Dim O$(), L
For Each L In MthLinAyzPj(A)
    PushIAy O, MthArgAy(L)
Next
ArgAyzPj = AywDist(O)
End Function

Private Sub Z_ArgAyzPj()
Brw AyQSrt(ArgAyzPj(CurPj))
End Sub

Function ArgSfx$(Arg)
Dim B$
B = RmvNm(RmvPfxSpc(RmvPfxSpc(Arg, "Optional"), "ParamArray"))
ArgSfx = RTrim(TakBefOrAll(B, "=", NoTrim:=True))
End Function


Function ArgTy$(Arg)

End Function


