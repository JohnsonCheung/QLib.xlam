Attribute VB_Name = "QIde_Mth_Dcl"
Option Explicit
Private Const CMod$ = "MIde_Mth_Dcl."
Private Const Asm$ = "QIde"

Property Get MthLinSyM() As String()
MthLinSyM = MthLinSyzMd(CurMd)
End Property

Function MthLinSyzNmSrc(Src$(), MthNm$) As String()
Dim Ix
For Each Ix In Itr(MthIxAyzNm(Src, MthNm))
    PushI MthLinSyzNmSrc, ContLin(Src, Ix)
Next
End Function

Private Sub Z_Src_PthMthLinSy()
Dim MthNy$(), Src$()
Src = CurSrc
MthNy = Sy("Src_MthDclDry", "Mth_MthDclLin")
Ept = Sy("Function Mth_MthDclLin$(A As Mth)", "Function Src_MthDclDry(A$()) As Variant()")
GoSub Tst
Exit Sub
Tst:
    Act = MthLinSyzSrc(Src)
    C
    Return
End Sub
