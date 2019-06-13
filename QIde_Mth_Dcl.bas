Attribute VB_Name = "QIde_Mth_Dcl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Dcl."
Private Const Asm$ = "QIde"

Function MthLinAyM() As String()
MthLinAyM = MthLinAyzM(CMd)
End Function
Function MthLinAyzM(M As CodeModule) As String()
MthLinAyzM = MthLinAyzS(Src(M))
End Function
Function MthLinAyzS(Src$()) As String()
Dim Ix: For Each Ix In Itr(MthIxy(Src))
    PushI MthLinAyzS, ContLin(Src, Ix)
Next
End Function

Function MthLinAyzSN(Src$(), Mthn) As String()
Dim Ix
For Each Ix In Itr(MthIxyzSN(Src, Mthn))
    PushI MthLinAyzSN, ContLin(Src, Ix)
Next
End Function

Private Sub Z_Src_PthMthLinAy()
Dim Mthny$(), Src$()
Src = CSrc
Mthny = Sy("Src_MthDclDry", "Mth_MthDclLin")
Ept = Sy("Function Mth_MthDclLin(A As Mth)", "Function Src_MthDclDry(A$()) As Variant()")
GoSub Tst
Exit Sub
Tst:
    Act = MthLinAyzS(Src)
    C
    Return
End Sub
