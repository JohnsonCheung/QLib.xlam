Attribute VB_Name = "QIde_Fun_FmtMulLinSrc"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Fun_FmtMulLinSrc."
Private Const Asm$ = "QIde"
Private Function DyoMulStmtSrc(MulStmtSrc$()) As Variant()
Dim ODy(): ODy = DyoLyWithColon(MulStmtSrc)
Dim Dr, J%, I&
For Each Dr In ODy
    For J = 0 To UB(Dr) - 1
        Dr(J) = Dr(J) & ":"
    Next
    ODy(I) = Dr
    I = I + 1
Next
DyoMulStmtSrc = ODy
End Function

Function FmtMulStmtSrc(MulStmtSrc$()) As String()
FmtMulStmtSrc = FmtDy(DyoMulStmtSrc(MulStmtSrc), MaxColWdt:=200)
End Function


