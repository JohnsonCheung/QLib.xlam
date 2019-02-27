Attribute VB_Name = "MIde_Mth_Dcl"
Option Explicit

Property Get CurMthLinAyzMd() As String()
CurMthLinAyzMd = MthLinAyzMd(CurMd)
End Property

Function MthLinAyzSrcNm(A$(), MthNm$) As String()
Dim Ix
For Each Ix In Itr(MthIxAyMth(A, MthNm))
    PushI MthLinAyzSrcNm, ContLin(A, Ix)
Next
End Function

Private Sub Z_Src_PthMthLinAy()
Dim MthNy$(), Src$()
Src = SrcMd
MthNy = Sy("Src_MthDclDry", "Mth_MthDclLin")
Ept = Sy("Function Mth_MthDclLin$(A As Mth)", "Function Src_MthDclDry(A$()) As Variant()")
GoSub Tst
Exit Sub
Tst:
    Act = MthLinAyzSrc(Src)
    C
    Return
End Sub
