Attribute VB_Name = "QIde_Fun_FmtMulLinSrc"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Fun_FmtMulLinSrc."
Private Const Asm$ = "QIde"
Private Function DryzMulStmtSrc(MulStmtSrc$()) As Variant()
Dim ODry(): ODry = DryzLyWithColon(MulStmtSrc)
Dim Dr, J%, I&
For Each Dr In ODry
    For J = 0 To UB(Dr) - 1
        Dr(J) = Dr(J) & ":"
    Next
    ODry(I) = Dr
    I = I + 1
Next
DryzMulStmtSrc = ODry
End Function

Function FmtMulStmtSrc(MulStmtSrc$()) As String()
FmtMulStmtSrc = FmtDryAsJnSep(DryzMulStmtSrc(MulStmtSrc), MaxColWdt:=200)
End Function


