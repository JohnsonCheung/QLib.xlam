Attribute VB_Name = "QIde_Mth_Dcl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Dcl."
Private Const Asm$ = "QIde"

Property Get MthLinyM() As String()
MthLinyM = MthLinyzMd(CMd)
End Property

Function MthLinyzNmSrc(Src$(), Mthn) As String()
Dim Ix
'For Each Ix In Itr(MthIxyzNm(Src, Mthn))
    PushI MthLinyzNmSrc, ContLin(Src, Ix)
'Next
End Function

Private Sub Z_Src_PthMthLiny()
Dim Mthny$(), Src$()
Src = CSrc
Mthny = Sy("Src_MthDclDry", "Mth_MthDclLin")
Ept = Sy("Function Mth_MthDclLin(A As Mth)", "Function Src_MthDclDry(A$()) As Variant()")
GoSub Tst
Exit Sub
Tst:
    Act = MthLinyzS(Src)
    C
    Return
End Sub
